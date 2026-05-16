import os, base64
from datetime import date, datetime, timedelta
from email.mime.text import MIMEText
from functools import wraps

from dotenv import load_dotenv
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

load_dotenv()
os.environ.setdefault("OAUTHLIB_INSECURE_TRANSPORT", "1")  # localhost HTTP only

app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "dev-secret")
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///conges.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

SCOPES = [
    "openid",
    "https://www.googleapis.com/auth/userinfo.email",
    "https://www.googleapis.com/auth/userinfo.profile",
    "https://www.googleapis.com/auth/gmail.send",
]

MONTHS_FR = ["Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre"]

HOLIDAYS_FIXED = [(1,1),(1,11),(5,1),(7,30),(8,14),(8,20),(8,21),(11,6),(11,18)]

def is_holiday(d):
    return (d.month, d.day) in HOLIDAYS_FIXED

def is_workday(d):
    return d.weekday() < 5 and not is_holiday(d)

def count_workdays(start, end):
    if start > end:
        raise ValueError("La date début doit être inférieure ou égale à la date fin.")
    total = 0
    cur = start
    while cur <= end:
        if is_workday(cur):
            total += 1
        cur += timedelta(days=1)
    return total

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    email = db.Column(db.String(180))
    google_email = db.Column(db.String(180))

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    days = db.Column(db.Float, nullable=False)
    period = db.Column(db.String(120), nullable=False)
    status = db.Column(db.String(30), default="En attente")
    created_at = db.Column(db.DateTime, default=datetime.now)

class MonthlyBalance(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    year = db.Column(db.Integer, nullable=False)
    month = db.Column(db.Integer, nullable=False)
    credit = db.Column(db.Float, default=1.5)
    extra_days = db.Column(db.Float, default=0)
    taken = db.Column(db.Float, default=0)
    period = db.Column(db.String(255), default="")
    balance = db.Column(db.Float, default=0)
    __table_args__ = (db.UniqueConstraint("year", "month", name="uix_year_month"),)

def init_seed():
    db.create_all()
    if not User.query.filter_by(username="admin").first():
        db.session.add(User(username="admin", password_hash=generate_password_hash("admin123"), email=""))
    # Seed from Août 2023 to Décembre 2027 using the values provided
    data = [
        (2023,8,1.5,0,0,"",1.5),(2023,9,1.5,0,0,"",3),(2023,10,1.5,0,0,"",4.5),(2023,11,1.5,0,0,"",6),(2023,12,1.5,0,0,"",7.5),
        (2024,1,1.5,0,3,"08-12/01/2024",6),(2024,2,1.5,0,0,"",7.5),(2024,3,1.5,0,0,"",9),(2024,4,1.5,0,4,"05-12/04/2024",6.5),(2024,5,1.5,0,0,"",8),(2024,6,1.5,0,3,"17-21/06/2024",6.5),(2024,7,1.5,0,0,"",8),(2024,8,1.5,0,3,"19-25/08/2024",6.5),(2024,9,1.5,0,5,"23-27/09/2024",3),(2024,10,1.5,0,0,"",4.5),(2024,11,1.5,1,1,"19/11/2024",6),(2024,12,1.5,1,0,"",8.5),
        (2025,1,1.5,1,4,"28-31/01/2025",7),(2025,2,1.5,0,0,"",8.5),(2025,3,1.5,0,0,"",10),(2025,4,1.5,0,5,"27/03/2025-07/04/2025",6.5),(2025,5,1.5,0,0,"",8),(2025,6,1.5,0,7,"05-13/06/2025",2.5),(2025,7,1.5,0,3,"09-11/07/2025",1),(2025,8,1.5,0,5,"15-29/08/2025",-2.5),(2025,9,1.5,0,0,"",-1),(2025,10,1.5,0,0,"",0.5),(2025,11,1.5,0,0,"",2),(2025,12,1.5,0,0,"",3.5),
        (2026,1,1.5,0,1,"02/01/2026",4),(2026,2,1.5,0,0,"",5.5),(2026,3,1.5,0,2,"09,23/03/2026",5),(2026,4,1.5,0,0,"",6.5),(2026,5,1.5,0,4,"25/05/2026-01/06/2026",4),(2026,6,1.5,0,0,"",5.5),(2026,7,1.5,0,0,"",7),(2026,8,1.5,0,1,"24/08/2026",7.5),(2026,9,1.5,0,0,"",9),(2026,10,1.5,0,0,"",10.5),(2026,11,1.5,0,4,"02-06/11/2026",8),(2026,12,1.5,0,0,"",9.5),
        (2027,1,1.5,0,3,"11-15/01/2027",8),(2027,2,1.5,0,0,"",9.5),(2027,3,1.5,0,3,"08-12/03/2027",8),(2027,4,1.5,0,0,"",9.5),(2027,5,1.5,0,3,"17-21/05/2027",8),(2027,6,1.5,0,0,"",9.5),(2027,7,1.5,0,0,"",11),(2027,8,1.5,0,0,"",12.5),(2027,9,1.5,0,0,"",14),(2027,10,1.5,0,0,"",15.5),(2027,11,1.5,0,0,"",17),(2027,12,1.5,0,0,"",18.5)
    ]
    for y,m,c,e,t,p,b in data:
        if not MonthlyBalance.query.filter_by(year=y, month=m).first():
            db.session.add(MonthlyBalance(year=y, month=m, credit=c, extra_days=e, taken=t, period=p, balance=b))
    db.session.commit()

def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(url_for("login"))
        return fn(*args, **kwargs)
    return wrapper

def google_flow():
    cid = os.getenv("GOOGLE_CLIENT_ID")
    secret = os.getenv("GOOGLE_CLIENT_SECRET")
    redirect_uri = os.getenv("GOOGLE_REDIRECT_URI", "http://127.0.0.1:5000/oauth2callback")
    if not cid or not secret:
        return None
    return Flow.from_client_config(
        {"web": {"client_id": cid, "client_secret": secret, "auth_uri": "https://accounts.google.com/o/oauth2/auth", "token_uri": "https://oauth2.googleapis.com/token", "redirect_uris": [redirect_uri]}},
        scopes=SCOPES,
        redirect_uri=redirect_uri,
    )

def credentials_from_session():
    c = session.get("google_credentials")
    if not c: return None
    return Credentials(token=c.get("token"), refresh_token=c.get("refresh_token"), token_uri=c.get("token_uri"), client_id=c.get("client_id"), client_secret=c.get("client_secret"), scopes=c.get("scopes"))

def send_gmail_oauth(to, subject, body):
    creds = credentials_from_session()
    if not creds:
        raise RuntimeError("Connecte-toi avec Google avant d'envoyer un email.")
    service = build("gmail", "v1", credentials=creds)
    msg = MIMEText(body, "plain", "utf-8")
    msg["To"] = to
    msg["From"] = "me"
    msg["Subject"] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        u = User.query.filter_by(username=request.form.get("username","").strip()).first()
        if u and check_password_hash(u.password_hash, request.form.get("password","")):
            session["user_id"] = u.id
            session["username"] = u.username
            return redirect(url_for("index"))
        flash("Login ou mot de passe incorrect.", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

@app.route("/login/google")
def login_google():
    flow = google_flow()
    if not flow:
        flash("Google OAuth n'est pas configuré. Vérifie GOOGLE_CLIENT_ID et GOOGLE_CLIENT_SECRET dans .env", "danger")
        return redirect(url_for("login"))
    auth_url, state = flow.authorization_url(access_type="offline", include_granted_scopes="true", prompt="consent")
    session["state"] = state
    return redirect(auth_url)

@app.route("/oauth2callback")
def oauth2callback():
    flow = google_flow()
    if not flow:
        flash("Google OAuth non configuré.", "danger")
        return redirect(url_for("login"))
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    session["google_credentials"] = {"token": creds.token, "refresh_token": creds.refresh_token, "token_uri": creds.token_uri, "client_id": creds.client_id, "client_secret": creds.client_secret, "scopes": creds.scopes}
    user = User.query.filter_by(username="admin").first()
    session["user_id"] = user.id
    session["username"] = user.username
    flash("Connexion Google réussie. L'envoi Gmail API est activé.", "success")
    return redirect(url_for("index"))

@app.route("/")
@login_required
def index():
    rows = MonthlyBalance.query.order_by(MonthlyBalance.year, MonthlyBalance.month).all()
    requests = LeaveRequest.query.order_by(LeaveRequest.created_at.desc()).all()
    return render_template("index.html", rows=rows, requests=requests, months=MONTHS_FR)

@app.route("/request", methods=["POST"])
@login_required
def make_request():
    try:
        sd = datetime.strptime(request.form["start_date"], "%Y-%m-%d").date()
        ed = datetime.strptime(request.form["end_date"], "%Y-%m-%d").date()
        days = count_workdays(sd, ed)
        if days <= 0:
            raise ValueError("Cette période ne contient aucun jour ouvrable : weekend ou jour férié uniquement.")
        existing = LeaveRequest.query.filter(
            LeaveRequest.status.in_(["En attente", "Approuvé"]),
            LeaveRequest.start_date <= ed,
            LeaveRequest.end_date >= sd,
        ).first()
        if existing:
            raise ValueError(f"Cette période chevauche déjà une demande {existing.status.lower()} : {existing.period}.")
        period = sd.strftime("%d/%m/%Y") if sd == ed else f"{sd.strftime('%d/%m/%Y')}-{ed.strftime('%d/%m/%Y')}"
        lr = LeaveRequest(start_date=sd, end_date=ed, days=days, period=period, status="En attente")
        db.session.add(lr); db.session.commit()
        if request.form.get("send_email"):
            to = request.form.get("email_to","").strip()
            if not to: raise ValueError("Destinataire email obligatoire.")
            subject = f"Demande de congé - {period}"
            body = request.form.get("message") or f"Bonjour,\n\nJe souhaite demander un congé pour la période {period}.\nNombre de jours ouvrables : {days}.\n\nCordialement"
            send_gmail_oauth(to, subject, body)
            flash("Demande enregistrée et email envoyé via Gmail API OAuth.", "success")
        else:
            flash("Demande créée en attente d'approbation.", "success")
    except Exception as e:
        flash(str(e), "danger")
    return redirect(url_for("index"))

@app.route("/approve/<int:req_id>")
@login_required
def approve(req_id):
    lr = LeaveRequest.query.get_or_404(req_id)
    if lr.status == "Approuvé":
        flash("Cette demande est déjà approuvée.", "warning")
        return redirect(url_for("index"))
    lr.status = "Approuvé"
    mb = MonthlyBalance.query.filter_by(year=lr.start_date.year, month=lr.start_date.month).first()
    if mb:
        mb.taken += lr.days
        mb.period = (mb.period + ", " if mb.period else "") + lr.period
        # recompute from this row forward
        rows = MonthlyBalance.query.order_by(MonthlyBalance.year, MonthlyBalance.month).all()
        prev = 0
        for r in rows:
            r.balance = prev + r.credit + r.extra_days - r.taken
            prev = r.balance
    db.session.commit()
    flash("Congé approuvé : jours pris, période, solde et X calendrier mis à jour.", "success")
    return redirect(url_for("index"))

@app.route("/reject/<int:req_id>", methods=["POST", "GET"])
@login_required
def reject(req_id):
    lr = LeaveRequest.query.get_or_404(req_id)
    if lr.status == "Approuvé":
        flash("Impossible de refuser une demande déjà approuvée : les jours ont déjà été comptabilisés.", "danger")
        return redirect(url_for("index"))
    if lr.status == "Refusé":
        flash("Cette demande est déjà refusée.", "warning")
        return redirect(url_for("index"))
    lr.status = "Refusé"
    db.session.commit()
    flash("Demande refusée : aucun X ajouté et le solde n'est pas modifié.", "success")
    return redirect(url_for("index"))

@app.route("/database-info")
@login_required
def database_info():
    db_path = os.path.join(app.instance_path, "conges.db")
    return f"Base SQLite : {db_path}"

with app.app_context():
    init_seed()

if __name__ == "__main__":
    app.run(debug=True)
