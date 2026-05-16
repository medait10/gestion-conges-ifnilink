
const UI_TRANSLATIONS = {
  de: {
    "Dashboard":"Dashboard",
    "Guide":"Benutzerhandbuch",
    "Subscription":"Abonnement",
    "Leave Request":"Urlaubsantrag",
    "History":"Verlauf",
    "Calendar":"Kalender",
    "Profile":"Profil",
    "About":"Über uns",
    "Copyright":"Urheberrecht",
    "Professional leave management, beautifully simple.":"Professionelles Urlaubsmanagement – einfach und modern.",
    "Start 7-day trial":"7 Tage testen",
    "Sign in":"Anmelden",
    "Secure workspace":"Sicherer Arbeitsbereich",
    "Monthly":"Monatlich",
    "Annual":"Jährlich",
    "Trial":"Testversion",
    "Use trial":"Testversion nutzen",
    "Subscribe monthly":"Monatlich abonnieren",
    "Subscribe annually":"Jährlich abonnieren",
    "Payment setup for owner":"Zahlungseinrichtung für den Besitzer"
  },
  fr: {},
  es: {
    "Dashboard":"Panel",
    "Guide":"Guía",
    "Subscription":"Suscripción",
    "Leave Request":"Solicitud",
    "History":"Historial",
    "Calendar":"Calendario",
    "Profile":"Perfil"
  },
  ar: {
    "Dashboard":"لوحة التحكم",
    "Guide":"الدليل",
    "Subscription":"الاشتراك",
    "Leave Request":"طلب عطلة",
    "History":"السجل",
    "Calendar":"التقويم",
    "Profile":"الملف الشخصي"
  }
};

function deepTranslate(lang){
  if(lang === "fr") return;
  const dict = UI_TRANSLATIONS[lang] || {};
  document.querySelectorAll("*").forEach(el=>{
    if(el.children.length === 0){
      let txt = el.textContent.trim();
      if(dict[txt]){
        el.textContent = dict[txt];
      }
    }
    if(el.placeholder && dict[el.placeholder]){
      el.placeholder = dict[el.placeholder];
    }
  });
}

window.addEventListener("DOMContentLoaded", ()=>{
  deepTranslate(window.MB_LANG || "fr");
});
