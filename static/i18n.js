
const V35_FALLBACK = {"en": {"Gestion congés secteur privé Maroc": "Morocco private sector leave management", "Bilan mensuel": "Monthly summary", "Synchronisation depuis Google Calendar": "Sync from Google Calendar", "Toutes les demandes": "All requests", "Nouvelle demande de congé": "New leave request", "Créer la demande": "Create request", "Dernières demandes": "Latest requests", "Centre d’aide": "Help center", "Backups base de données": "Database backups", "Mini panneau Admin DB": "Mini Admin DB panel", "Historique des demandes": "Request history", "Profil": "Profile", "Enregistrer": "Save", "Annuler": "Cancel"}, "de": {"Gestion congés secteur privé Maroc": "Urlaubsverwaltung Privatsektor Marokko", "Bilan mensuel": "Monatsübersicht", "Synchronisation depuis Google Calendar": "Synchronisierung aus Google Calendar", "Toutes les demandes": "Alle Anträge", "Nouvelle demande de congé": "Neuer Urlaubsantrag", "Créer la demande": "Antrag erstellen", "Dernières demandes": "Letzte Anträge", "Centre d’aide": "Hilfezentrum", "Backups base de données": "Datenbank-Backups", "Mini panneau Admin DB": "Mini Admin-DB-Panel", "Historique des demandes": "Antragshistorie", "Profil": "Profil", "Enregistrer": "Speichern", "Annuler": "Stornieren"}, "es": {"Gestion congés secteur privé Maroc": "Gestión de permisos sector privado Marruecos", "Bilan mensuel": "Resumen mensual", "Synchronisation depuis Google Calendar": "Sincronización desde Google Calendar", "Toutes les demandes": "Todas las solicitudes", "Nouvelle demande de congé": "Nueva solicitud de permiso", "Créer la demande": "Crear solicitud", "Dernières demandes": "Últimas solicitudes", "Centre d’aide": "Centro de ayuda", "Backups base de données": "Backups de base de datos", "Mini panneau Admin DB": "Panel Admin DB", "Historique des demandes": "Historial de solicitudes", "Profil": "Perfil", "Enregistrer": "Guardar", "Annuler": "Cancelar"}, "ar": {"Gestion congés secteur privé Maroc": "تدبير العطل للقطاع الخاص بالمغرب", "Bilan mensuel": "الحصيلة الشهرية", "Synchronisation depuis Google Calendar": "مزامنة من Google Calendar", "Toutes les demandes": "كل الطلبات", "Nouvelle demande de congé": "طلب عطلة جديد", "Créer la demande": "إنشاء الطلب", "Dernières demandes": "آخر الطلبات", "Centre d’aide": "مركز المساعدة", "Backups base de données": "نسخ قاعدة البيانات", "Mini panneau Admin DB": "لوحة إدارة قاعدة البيانات", "Historique des demandes": "سجل الطلبات", "Profil": "الملف الشخصي", "Enregistrer": "حفظ", "Annuler": "إلغاء"}};
function v35Translate(lang){
  if(lang === "fr") return;
  const d = V35_FALLBACK[lang] || {};
  const keys = Object.keys(d).sort((a,b)=>b.length-a.length);
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
    acceptNode(n){ const p=n.parentElement; if(!p || ["SCRIPT","STYLE","TEXTAREA"].includes(p.tagName)) return NodeFilter.FILTER_REJECT; return n.nodeValue.trim().length>1 ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_REJECT; }
  });
  let n;
  while(n=walker.nextNode()) {
    let s=n.nodeValue;
    keys.forEach(k=>{ if(s.includes(k)) s=s.split(k).join(d[k]); });
    n.nodeValue=s;
  }
  document.querySelectorAll("input,textarea,option,button").forEach(el=>{
    ["placeholder","value","title"].forEach(a=>{let v=el.getAttribute(a); if(v){keys.forEach(k=>{if(v.includes(k)) v=v.split(k).join(d[k]);}); el.setAttribute(a,v);}});
    if(el.tagName==="OPTION"){let v=el.textContent; keys.forEach(k=>{if(v.includes(k)) v=v.split(k).join(d[k]);}); el.textContent=v;}
  });
  if(lang==="ar") document.documentElement.dir="rtl";
}
window.addEventListener("DOMContentLoaded",()=>v35Translate(window.MB_LANG || "fr"));
