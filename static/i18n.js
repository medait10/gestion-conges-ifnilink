
const MEDFLOW_EXTRA = {"en": {"Congé annuel payé": "Paid annual leave", "Naissance fils/fille": "Birth of child", "Décès parent / ascendant": "Death of parent / ascendant", "Décès conjoint / enfant": "Death of spouse / child", "Décès frère/sœur/beau-parent": "Death of sibling / parent-in-law", "Mariage du salarié": "Employee marriage", "Mariage d'un enfant": "Child marriage", "Circoncision": "Circumcision", "Opération conjoint/enfant à charge": "Spouse/dependent child operation", "Repos maladie": "Sick leave", "Bonjour, merci de bien vouloir approuver ma demande de congé.": "Hello, please approve my leave request.", "directeur@ifnilink.ma": "manager@example.com"}, "de": {"Congé annuel payé": "Bezahlter Jahresurlaub", "Naissance fils/fille": "Geburt eines Kindes", "Décès parent / ascendant": "Tod Elternteil / Vorfahr", "Décès conjoint / enfant": "Tod Ehepartner / Kind", "Décès frère/sœur/beau-parent": "Tod Geschwister / Schwiegereltern", "Mariage du salarié": "Heirat des Mitarbeiters", "Mariage d'un enfant": "Heirat eines Kindes", "Circoncision": "Beschneidung", "Opération conjoint/enfant à charge": "Operation Ehepartner/Kind", "Repos maladie": "Krankschreibung", "Bonjour, merci de bien vouloir approuver ma demande de congé.": "Hallo, bitte genehmigen Sie meinen Urlaubsantrag.", "directeur@ifnilink.ma": "manager@example.com"}, "es": {"Congé annuel payé": "Vacaciones anuales pagadas", "Naissance fils/fille": "Nacimiento de hijo/a", "Décès parent / ascendant": "Fallecimiento padre/madre/ascendiente", "Décès conjoint / enfant": "Fallecimiento cónyuge / hijo", "Décès frère/sœur/beau-parent": "Fallecimiento hermano/a / suegro/a", "Mariage du salarié": "Matrimonio del empleado", "Mariage d'un enfant": "Matrimonio de un hijo", "Circoncision": "Circuncisión", "Opération conjoint/enfant à charge": "Operación cónyuge/hijo a cargo", "Repos maladie": "Baja médica", "Bonjour, merci de bien vouloir approuver ma demande de congé.": "Hola, por favor apruebe mi solicitud de permiso.", "directeur@ifnilink.ma": "manager@example.com"}, "ar": {"Congé annuel payé": "عطلة سنوية مدفوعة", "Naissance fils/fille": "ازدياد ابن/ابنة", "Décès parent / ascendant": "وفاة أحد الوالدين / الأصول", "Décès conjoint / enfant": "وفاة الزوج/الزوجة أو الابن", "Décès frère/sœur/beau-parent": "وفاة أخ/أخت أو أحد الأصهار", "Mariage du salarié": "زواج الموظف", "Mariage d'un enfant": "زواج ابن/ابنة", "Circoncision": "الختان", "Opération conjoint/enfant à charge": "عملية للزوج/الطفل المكفول", "Repos maladie": "راحة مرضية", "Bonjour, merci de bien vouloir approuver ma demande de congé.": "مرحبا، المرجو الموافقة على طلب العطلة الخاص بي.", "directeur@ifnilink.ma": "manager@example.com"}};

function medflowTranslateOptions(lang){
  if(lang === "fr") return;
  const dict = MEDFLOW_EXTRA[lang] || {};
  document.querySelectorAll("option").forEach(o=>{
    const t=o.textContent.trim();
    if(dict[t]) o.textContent=dict[t];
  });
}

function medflowTranslateInputs(lang){
  if(lang === "fr") return;
  const dict = MEDFLOW_EXTRA[lang] || {};
  document.querySelectorAll("input, textarea").forEach(el=>{
    if(el.placeholder && dict[el.placeholder]) el.placeholder = dict[el.placeholder];
    if(el.value && dict[el.value]) el.value = dict[el.value];
  });
}

function applyAutoTranslations(lang){
  medflowTranslateOptions(lang);
  medflowTranslateInputs(lang);
  if(lang === "ar") document.documentElement.dir = "rtl";
}
window.addEventListener("DOMContentLoaded",()=>applyAutoTranslations(window.MB_LANG||"fr"));
