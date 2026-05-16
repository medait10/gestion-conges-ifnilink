const MB_TRANSLATIONS = {
  en: {
    "Gestion intelligente des congés et absences": "Smart leave and absence management",
    "Connecte-toi à ton espace sécurisé MEDAIT-BOQAL.": "Sign in to your secure MEDAIT-BOQAL workspace.",
    "Créer un compte": "Create account",
    "Se connecter": "Sign in",
    "Mot de passe": "Password",
    "Nom utilisateur": "Username",
    "Nouvelle demande de congé": "New leave request",
    "Choisis le type, la période et le destinataire. Les week-ends et jours fériés ne sont pas comptés.": "Choose the type, period and recipient. Weekends and holidays are not counted.",
    "Paiement sécurisé": "Secure payment",
    "Abonnement MEDAIT-BOQAL": "MEDAIT-BOQAL subscription",
    "7 jours d’essai": "7-day trial",
    "À propos de nous": "About us",
    "Droits d’auteur": "Copyright",
    "Secure leave management platform for teams and individuals.": "Secure leave management platform for teams and individuals.",
    "Payments are processed by Stripe Checkout. Card data is not stored here.": "Payments are processed by Stripe Checkout. Card data is not stored here."
  },
  de: {
    "Gestion intelligente des congés et absences": "Intelligente Verwaltung von Urlaub und Abwesenheiten",
    "Connecte-toi à ton espace sécurisé MEDAIT-BOQAL.": "Melde dich in deinem sicheren MEDAIT-BOQAL-Bereich an.",
    "Créer un compte": "Konto erstellen",
    "Se connecter": "Anmelden",
    "Mot de passe": "Passwort",
    "Nom utilisateur": "Benutzername",
    "Nouvelle demande de congé": "Neuer Urlaubsantrag",
    "Paiement sécurisé": "Sichere Zahlung",
    "Abonnement MEDAIT-BOQAL": "MEDAIT-BOQAL Abonnement",
    "7 jours d’essai": "7 Tage Testphase",
    "À propos de nous": "Über uns",
    "Droits d’auteur": "Urheberrecht",
    "Secure leave management platform for teams and individuals.": "Sichere Urlaubsplattform für Teams und Einzelpersonen.",
    "Payments are processed by Stripe Checkout. Card data is not stored here.": "Zahlungen werden über Stripe Checkout verarbeitet. Kartendaten werden hier nicht gespeichert."
  },
  es: {
    "Gestion intelligente des congés et absences": "Gestión inteligente de permisos y ausencias",
    "Connecte-toi à ton espace sécurisé MEDAIT-BOQAL.": "Inicia sesión en tu espacio seguro MEDAIT-BOQAL.",
    "Créer un compte": "Crear cuenta",
    "Se connecter": "Iniciar sesión",
    "Mot de passe": "Contraseña",
    "Nom utilisateur": "Usuario",
    "Nouvelle demande de congé": "Nueva solicitud de permiso",
    "Paiement sécurisé": "Pago seguro",
    "Abonnement MEDAIT-BOQAL": "Suscripción MEDAIT-BOQAL",
    "7 jours d’essai": "Prueba de 7 días",
    "À propos de nous": "Sobre nosotros",
    "Droits d’auteur": "Derechos de autor",
    "Secure leave management platform for teams and individuals.": "Plataforma segura de permisos para equipos y usuarios.",
    "Payments are processed by Stripe Checkout. Card data is not stored here.": "Los pagos se procesan con Stripe Checkout. Los datos de tarjeta no se almacenan aquí."
  },
  ar: {
    "Gestion intelligente des congés et absences": "منصة ذكية لتدبير العطل والغيابات",
    "Connecte-toi à ton espace sécurisé MEDAIT-BOQAL.": "سجّل الدخول إلى فضائك الآمن في MEDAIT-BOQAL.",
    "Créer un compte": "إنشاء حساب",
    "Se connecter": "تسجيل الدخول",
    "Mot de passe": "كلمة المرور",
    "Nom utilisateur": "اسم المستخدم",
    "Nouvelle demande de congé": "طلب عطلة جديد",
    "Paiement sécurisé": "دفع آمن",
    "Abonnement MEDAIT-BOQAL": "اشتراك MEDAIT-BOQAL",
    "7 jours d’essai": "تجربة لمدة 7 أيام",
    "À propos de nous": "من نحن",
    "Droits d’auteur": "حقوق النشر",
    "Secure leave management platform for teams and individuals.": "منصة آمنة لتدبير العطل للفرق والأفراد.",
    "Payments are processed by Stripe Checkout. Card data is not stored here.": "تتم معالجة المدفوعات عبر Stripe Checkout ولا يتم تخزين بيانات البطاقة هنا."
  }
};

function translateTextNode(node, lang){
  if(lang === "fr") return;
  const dict = MB_TRANSLATIONS[lang] || {};
  const original = node.nodeValue.trim();
  if(!original) return;
  if(dict[original]){
    node.nodeValue = node.nodeValue.replace(original, dict[original]);
  }
}

function applyAutoTranslations(lang){
  if(lang === "fr") return;
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
    acceptNode(node){
      const parent = node.parentElement;
      if(!parent) return NodeFilter.FILTER_REJECT;
      if(["SCRIPT","STYLE","TEXTAREA","INPUT","SELECT"].includes(parent.tagName)) return NodeFilter.FILTER_REJECT;
      const text = node.nodeValue.trim();
      return text.length > 1 ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT;
    }
  });
  let node;
  while(node = walker.nextNode()){
    translateTextNode(node, lang);
  }
}
