/**
 * Cockpit Conciergerie — étape 3 : API de l'application mobile (ménages).
 * Déploiement : Déployer > Application Web, exécuter en tant que MOI, accès « Tout le monde » (le code personnel protège l'accès).
 * Dépend de Ical.gs et Conciergerie.gs.
 */
var SESSION_TTL = 21600;      // 6 h (maximum du cache de script)
var MAX_FAILS = 10;           // essais de code ratés tolérés par fenêtre de 10 min
var MAX_PHOTO_BYTES = 3 * 1024 * 1024;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Cockpit Conciergerie')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Accès depuis l'interface hébergée hors Google (GitHub Pages…) : POST texte brut {fn:'login'|'api', args:[…]}.
 *  Seules ces deux fonctions sont joignables ; le jeton de session et le code personnel protègent tout le reste. */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents), a = req.args || [];
    if (req.fn === 'login') out = login(a[0]);
    else if (req.fn === 'api') out = api(a[0], a[1] || {});
    else out = { ok: false, error: 'Fonction inconnue' };
  } catch (err) { out = { ok: false, error: 'Requête invalide' }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- Connexion par code personnel ----------
function login(code) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: 'Trop d\'essais. Réessayez dans 10 minutes.' };
  var p = authenticate(String(code || '').trim());
  if (!p) { cache.put('fails', String(fails + 1), 600); return { ok: false, error: 'Code invalide' }; }
  var token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s:' + token, JSON.stringify(p), SESSION_TTL);
  journal_(p.id, 'connexion', 'Codes', p.id, '');
  return { ok: true, token: token, me: p };
}
function session_(token) {
  var v = token && CacheService.getScriptCache().get('s:' + token);
  if (!v) throw new Error('SESSION');
  var me = JSON.parse(v);
  // Un code révoqué coupe aussi les sessions en cours.
  var p = ref_('Codes').rows.filter(function (r) { return r['ID personne'] === me.id; })[0];
  if (!p || p['Actif'] !== 'Oui' || p['Révoqué le']) throw new Error('SESSION');
  return me;
}

// ---------- Point d'entrée unique ----------
/** op = { opId, action, payload }. Les écritures sont dédoublonnées par opId : rejouer une saisie après une coupure ne l'applique pas deux fois. */
function api(token, op) {
  try {
    var me = session_(token);
    var h = READS[op.action], w = WRITES[op.action];
    if (h) return { ok: true, data: h(me, op.payload || {}) };
    if (!w) throw new Error('Action inconnue');
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) throw new Error('Serveur occupé, réessayez');
    try {
      if (!op.opId) throw new Error('Opération sans identifiant');
      if (done_(op.opId)) return { ok: true, dup: true };
      var data = w(me, op.payload || {});
      SpreadsheetApp.getActive().getSheetByName('OpsTraitees').appendRow([op.opId, new Date(), me.id, op.action]);
      return { ok: true, data: data };
    } finally { lock.releaseLock(); }
  } catch (e) {
    var m = String(e && e.message || e);
    return { ok: false, error: m === 'SESSION' ? 'Session expirée' : m, session: m === 'SESSION' };
  }
}
function done_(opId) {
  var sh = SpreadsheetApp.getActive().getSheetByName('OpsTraitees');
  var n = sh.getLastRow() - 1;
  if (n < 1) return false;
  return sh.getRange(2, 1, n, 1).getValues().some(function (r) { return r[0] === opId; });
}

// ---------- Aides ----------
function need_(me, roles) { if (roles.indexOf(me.role) < 0) throw new Error('Accès refusé'); }
function logementOf_(id) { return ref_('Logements').rows.filter(function (l) { return l['ID'] === id; })[0]; }
function points_(modele) {
  return ref_('ModelesListe').rows.filter(function (p) { return p['Modèle'] === modele; })
    .sort(function (a, b) { return a['Ordre'] - b['Ordre']; });
}
function parseJson_(s) { try { return s ? JSON.parse(s) : []; } catch (e) { return []; } }
function ids_(s) { return String(s || '').split(',').filter(function (x) { return x; }); }
/** Vue d'un ménage SANS aucune donnée voyageur : l'agent voit adresse, accès, date, liste. */
function menageView_(m, withDetail) {
  var log = logementOf_(m['ID logement']) || {};
  var pts = points_(m['Liste de contrôle']), coches = parseJson_(m['Cochés (JSON)']);
  var v = { id: m['ID'], date: day_(m['Date (suit la réservation)']), jour: m['Jour (déduit)'], statut: m['Statut'], agent: m['ID agent'],
    logement: { id: log['ID'], nom: log['Nom'], adresse: log['Adresse'] }, total: pts.length, coches: coches.length,
    photos: ids_(m['Photos (ids Drive)']).length };
  if (withDetail) {
    v.consignes = log['Consignes d\'accès'] || '';
    v.points = pts.map(function (p) { return { id: p['ID point'], texte: p['Point à contrôler'], coche: coches.indexOf(p['ID point']) > -1 }; });
    v.photoIds = ids_(m['Photos (ids Drive)']);
    v.linge = m['Linge manquant'] || '';
    v.debut = m['Début'] ? new Date(m['Début']).toISOString() : null;
  }
  return v;
}
function menageFor_(me, id) {
  var t = table_('Menages');
  var m = t.rows.filter(function (r) { return r['ID'] === id; })[0];
  if (!m) throw new Error('Ménage inconnu');
  if (me.role !== 'Gestion' && m['ID agent'] !== me.id) throw new Error('Accès refusé');
  return { t: t, m: m };
}
function touch_(t, m) { setCell_(t, m.__row, 'Dernière modif', new Date()); }

// ---------- Lectures ----------
var READS = {
  myMenages: function (me) {
    need_(me, ['Ménage']);
    var limit = new Date(); limit.setDate(limit.getDate() - 7);
    var min = Utilities.formatDate(limit, TZ, 'yyyy-MM-dd');
    var list = table_('Menages').rows.filter(function (m) { return m['ID agent'] === me.id && m['Statut'] !== 'Annulé' && day_(m['Date (suit la réservation)']) >= min; })
      .map(function (m) { return menageView_(m, false); })
      .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    return { today: today_(), menages: list, logements: myLogements_(me) };
  },
  menageDetail: function (me, p) { need_(me, ['Ménage', 'Gestion']); return menageView_(menageFor_(me, p.id).m, true); },
  listMenages: function (me) {   // gestion : tout le planning + alertes
    need_(me, ['Gestion']);
    var agents = ref_('Codes').rows.filter(function (c) { return c['Rôle'] === 'Ménage' && c['Actif'] === 'Oui'; })
      .map(function (c) { return { id: c['ID personne'], nom: c['Nom'] }; });
    var list = table_('Menages').rows.filter(function (m) { return m['Statut'] !== 'Annulé'; }).map(function (m) { return menageView_(m, false); })
      .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
    return { today: today_(), menages: list, agents: agents, alertes: alerts(), logements: allLogements_(), work: work_(), plan: planning_() };
  },
  /** Intervenant (Jardin et entretien) : ses interventions non closes, avec adresse et consignes d'accès. */
  myInterventions: function (me) {
    need_(me, ['Jardin et entretien']);
    var list = table_('Interventions').rows.filter(function (i) { return i['ID'] && i['ID agent'] === me.id && i['Statut'] !== 'Clôturée'; })
      .map(interventionView_)
      .sort(function (a, b) { return (a.urgence === 'Haute' ? 0 : 1) - (b.urgence === 'Haute' ? 0 : 1); });
    return { today: today_(), interventions: list, logements: allLogements_() };
  },
  getPhoto: function (me, p) {   // gestion seulement, et seulement une photo enregistrée dans le classeur
    need_(me, ['Gestion']);
    var ok = table_('Menages').rows.some(function (m) { return ids_(m['Photos (ids Drive)']).indexOf(p.fileId) > -1; }) ||
             table_('Interventions').rows.some(function (m) { return ids_(m['Photos (ids Drive)']).indexOf(p.fileId) > -1; });
    if (!ok) throw new Error('Photo inconnue');
    var b = DriveApp.getFileById(p.fileId).getBlob();
    return 'data:image/jpeg;base64,' + Utilities.base64Encode(b.getBytes());
  }
};
// ---------- Fiches logement et calendrier (vue Gestion) ----------
function planning_() {
  var t = today_(), lo = new Date(); lo.setDate(lo.getDate() - 7);
  var min = Utilities.formatDate(lo, TZ, 'yyyy-MM-dd');
  var resas = table_('Reservations').rows.filter(function (r) { return r['ID'] && r['Statut'] !== 'Annulée' && day_(r['Départ']) >= min; })
    .map(function (r) { return { id: r['ID'], logement: r['ID logement'], start: day_(r['Arrivée']), end: day_(r['Départ']), canal: r['Canal'], voyageur: r['Voyageur (saisi à la main)'] || '', caution: r['Caution (statut)'] || '' }; });
  var logements = ref_('Logements').rows.map(function (l) {
    var s = l['Dernière synchro iCal'];
    return { id: l['ID'], nom: l['Nom'], adresse: l['Adresse'] || '', type: l['Type'] || '', proprio: l['ID propriétaire'] || '', consignes: l['Consignes d\'accès'] || '',
      ical: !!l['Lien iCal'], icalActif: !icalOff_(l), synchro: s instanceof Date ? Utilities.formatDate(s, TZ, 'dd/MM HH:mm') : '' };
  });
  return { resas: resas, logements: logements };
}
// ---------- Interventions : vue de l'intervenant ----------
function interventionView_(i) {
  var l = logementOf_(i['ID logement']) || {};
  return { id: i['ID'], logement: { id: l['ID'], nom: l['Nom'] || i['ID logement'], adresse: l['Adresse'] || '' }, consignes: l['Consignes d\'accès'] || '',
    type: i['Type'], description: i['Description'], statut: i['Statut'], urgence: i['Urgence'], photos: ids_(i['Photos (ids Drive)']).length };
}
function interventionFor_(me, id) {
  var x = rowById_('Interventions', id, 'Intervention');
  if (x.r['ID agent'] !== me.id) throw new Error('Accès refusé');
  return x;
}
function closeLinkedTask_(me, x) {      // une intervention close ne laisse pas de tâche ouverte
  if (!x.r['ID tâche']) return;
  var tt = rowById_('Taches', x.r['ID tâche'], 'Tâche');
  if (tt.r['Statut'] !== 'Faite') { setCell_(tt.t, tt.r.__row, 'Statut', 'Faite'); journal_(me.id, 'tache_statut', 'Taches', tt.r['ID'], 'Faite (intervention clôturée)'); }
}
// ---------- Interventions et tâches (vue Gestion) ----------
var STATUTS_INTERVENTION = ['Ouverte', 'Planifiée', 'En cours', 'Clôturée'];
var STATUTS_TACHE = ['À faire', 'En cours', 'Faite'];
var PRIORITES_TACHE = ['Haute', 'Normale', 'Basse'];
function safe_(v, max) { var t = String(v === undefined || v === null ? '' : v).trim().substring(0, max); return /^[=+\-@]/.test(t) ? "'" + t : t; }
function stamp_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v || ''); }
function people_() {
  return ref_('Codes').rows.filter(function (c) { return c['Actif'] === 'Oui' && (c['Rôle'] === 'Gestion' || c['Rôle'] === 'Jardin et entretien' || c['Rôle'] === 'Ménage'); })
    .map(function (c) { return { id: c['ID personne'], nom: c['Nom'], role: c['Rôle'] }; });
}
/** Personnes mentionnées par « @Nom » dans un texte (le plus long nom d'abord, mention en début de mot). */
function mentionIds_(txt) {
  var low = ' ' + String(txt || '').toLowerCase(), out = [];
  people_().slice().sort(function (a, b) { return b.nom.length - a.nom.length; }).forEach(function (p) {
    var tag = '@' + String(p.nom).toLowerCase(), i = low.indexOf(tag);
    while (i >= 0) {
      var nx = low.charAt(i + tag.length);
      if (!/[a-z0-9@À-ſ]/.test(low.charAt(i - 1)) && (nx === '' || !/[a-z0-9\u00c0-\u017f]/.test(nx))) {
        if (out.indexOf(p.id) < 0) out.push(p.id);
        low = low.substring(0, i) + new Array(tag.length + 1).join(' ') + low.substring(i + tag.length); break;
      }
      i = low.indexOf(tag, i + 1);
    }
  });
  return out;
}
function personName_(id) { var p = people_().filter(function (x) { return x.id === id; })[0]; return p ? p.nom : ''; }
/** Valide les champs d'une tâche. partial = modification : seuls les champs fournis sont contrôlés. */
function tacheFields_(p, partial) {
  var o = {}, has = function (k) { return p[k] !== undefined && p[k] !== null; };
  if (has('titre') || !partial) { var t = safe_(p.titre, 200); if (!t) throw new Error('Donnez un titre à la tâche'); o['Titre'] = t; }
  if (has('priorite')) { if (PRIORITES_TACHE.indexOf(p.priorite) < 0) throw new Error('Priorité inconnue'); o['Priorité'] = p.priorite; }
  if (has('statut')) { if (STATUTS_TACHE.indexOf(p.statut) < 0) throw new Error('Statut inconnu'); o['Statut'] = p.statut; }
  if (has('details')) o['Détails'] = safe_(p.details, 2000);
  if (has('contact')) o['Contact'] = safe_(p.contact, 120);
  if (has('tel')) { var tl = String(p.tel).trim(); if (!/^[0-9 +().\-]{0,40}$/.test(tl)) throw new Error('Téléphone invalide'); o['Téléphone'] = tl; }
  if (has('email')) { var em = String(p.email).trim(); if (em && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) throw new Error('E-mail invalide'); o['E-mail'] = em.substring(0, 120); }
  if (has('responsable')) { if (p.responsable && !people_().some(function (x) { return x.id === p.responsable; })) throw new Error('Personne inconnue'); o['ID responsable'] = p.responsable; }
  if (has('logement')) { if (p.logement && !logementOf_(p.logement)) throw new Error('Logement inconnu'); o['ID logement (facultatif)'] = p.logement; }
  if (has('echeance')) { if (p.echeance !== '' && !/^\d{4}-\d{2}-\d{2}$/.test(String(p.echeance))) throw new Error('Échéance invalide'); o['Échéance'] = p.echeance === '' ? '' : dateOf_(p.echeance); }
  return o;
}
function work_() {
  var intervenants = ref_('Codes').rows.filter(function (c) { return (c['Rôle'] === 'Jardin et entretien' || c['Rôle'] === 'Ménage') && c['Actif'] === 'Oui'; })
    .map(function (c) { return { id: c['ID personne'], nom: c['Nom'], role: c['Rôle'] }; });
  var interventions = table_('Interventions').rows.filter(function (i) { return i['ID']; }).map(function (i) {
    var l = logementOf_(i['ID logement']) || {};
    return { id: i['ID'], logement: { id: l['ID'], nom: l['Nom'] || i['ID logement'] }, type: i['Type'], description: i['Description'], statut: i['Statut'],
      urgence: i['Urgence'], agent: i['ID agent'], tache: i['ID tâche'], photos: ids_(i['Photos (ids Drive)']).length };
  });
  var coms = {};
  table_('Commentaires').rows.forEach(function (c) { (coms[c['ID tâche']] = coms[c['ID tâche']] || []).push({ id: c['ID'], par: personName_(c['Auteur']) || c['Auteur'], date: stamp_(c['Date']), txt: String(c['Texte']), ment: mentionIds_(c['Texte']) }); });
  var taches = table_('Taches').rows.filter(function (t) { return t['ID']; }).map(function (t) {
    var l = logementOf_(t['ID logement (facultatif)']) || {};
    return { id: t['ID'], titre: t['Titre'], echeance: t['Échéance'] ? day_(t['Échéance']) : '', responsable: t['ID responsable'], statut: t['Statut'],
      logement: l['Nom'] || '', logementId: l['ID'] || '', intervention: t['ID intervention'],
      priorite: t['Priorité'] || 'Normale', details: t['Détails'] || '', contact: t['Contact'] || '', tel: String(t['Téléphone'] || ''), email: t['E-mail'] || '',
      creePar: personName_(t['Créée par']) || t['Créée par'] || '', creeLe: stamp_(t['Créée le']), com: coms[t['ID']] || [] };
  });
  return { interventions: interventions, taches: taches, intervenants: intervenants, people: people_() };
}
function rowById_(tab, id, what) {
  var t = table_(tab), r = t.rows.filter(function (x) { return x['ID'] === id; })[0];
  if (!r) throw new Error(what + ' inconnu(e)');
  return { t: t, r: r };
}
function myLogements_(me) {
  var ids = {};
  table_('Menages').rows.forEach(function (m) { if (m['ID agent'] === me.id) ids[m['ID logement']] = true; });
  return ref_('Logements').rows.filter(function (l) { return ids[l['ID']]; }).map(function (l) { return { id: l['ID'], nom: l['Nom'] }; });
}
function allLogements_() { return ref_('Logements').rows.map(function (l) { return { id: l['ID'], nom: l['Nom'] }; }); }

// ---------- Écritures ----------
var WRITES = {
  startMenage: function (me, p) {
    need_(me, ['Ménage']);
    var x = menageFor_(me, p.id), s = x.m['Statut'];
    if (s === 'En cours') return { statut: s };
    if (s !== 'Affecté') throw new Error('Ce ménage ne peut pas être démarré (' + s + ')');
    setCell_(x.t, x.m.__row, 'Statut', 'En cours'); setCell_(x.t, x.m.__row, 'Début', new Date()); touch_(x.t, x.m);
    journal_(me.id, 'menage_demarre', 'Menages', p.id, '');
    return { statut: 'En cours' };
  },
  saveProgress: function (me, p) {
    need_(me, ['Ménage']);
    var x = menageFor_(me, p.id);
    if (x.m['Statut'] !== 'En cours') throw new Error('Démarrez le ménage avant de cocher la liste');
    var valid = points_(x.m['Liste de contrôle']).map(function (q) { return q['ID point']; });
    var coches = (p.coches || []).filter(function (c) { return valid.indexOf(c) > -1; });
    setCell_(x.t, x.m.__row, 'Cochés (JSON)', JSON.stringify(coches));
    setCell_(x.t, x.m.__row, 'Linge manquant', String(p.linge || '').substring(0, 500));
    touch_(x.t, x.m);
    journal_(me.id, 'liste_mise_a_jour', 'Menages', p.id, coches.length + '/' + valid.length + (p.linge ? ' ; linge : ' + p.linge : ''));
    return { coches: coches.length, total: valid.length };
  },
  addPhoto: function (me, p) {
    need_(me, ['Ménage']);
    var x = menageFor_(me, p.id);
    if (x.m['Statut'] !== 'En cours') throw new Error('Démarrez le ménage avant d\'ajouter des photos');
    var fid = savePhoto_(p.data, p.id);
    var all = ids_(x.m['Photos (ids Drive)']).concat([fid]);
    setCell_(x.t, x.m.__row, 'Photos (ids Drive)', all.join(',')); setCell_(x.t, x.m.__row, 'Nb photos', all.length); touch_(x.t, x.m);
    journal_(me.id, 'photo_ajoutee', 'Menages', p.id, fid);
    return { photos: all.length };
  },
  finishMenage: function (me, p) {
    need_(me, ['Ménage']);
    var x = menageFor_(me, p.id), s = x.m['Statut'];
    if (s === 'Terminé') return { statut: s };
    if (s !== 'En cours') throw new Error('Ce ménage n\'est pas en cours (' + s + ')');
    setCell_(x.t, x.m.__row, 'Statut', 'Terminé'); setCell_(x.t, x.m.__row, 'Fin', new Date()); touch_(x.t, x.m);
    var n = parseJson_(x.m['Cochés (JSON)']).length, tot = points_(x.m['Liste de contrôle']).length;
    journal_(me.id, 'menage_termine', 'Menages', p.id, n + '/' + tot + ' points');
    return { statut: 'Terminé', coches: n, total: tot };
  },
  assignMenage: function (me, p) {
    need_(me, ['Gestion']);
    var x = menageFor_(me, p.id), s = x.m['Statut'];
    if (s === 'Terminé' || s === 'Annulé' || s === 'En cours') throw new Error('Ce ménage ne peut plus être réaffecté (' + s + ')');
    if (p.agent) {
      var ok = ref_('Codes').rows.some(function (c) { return c['ID personne'] === p.agent && c['Rôle'] === 'Ménage' && c['Actif'] === 'Oui'; });
      if (!ok) throw new Error('Agent inconnu ou inactif');
    }
    setCell_(x.t, x.m.__row, 'ID agent', p.agent || ''); setCell_(x.t, x.m.__row, 'Statut', p.agent ? 'Affecté' : 'À affecter'); touch_(x.t, x.m);
    journal_(me.id, 'menage_affecte', 'Menages', p.id, p.agent || 'désaffecté');
    return { statut: p.agent ? 'Affecté' : 'À affecter' };
  },
  assignIntervention: function (me, p) {
    need_(me, ['Gestion']);
    var x = rowById_('Interventions', p.id, 'Intervention'), st = x.r['Statut'];
    if (st === 'Clôturée') throw new Error('Intervention clôturée');
    if (p.agent) {
      var ok = ref_('Codes').rows.some(function (c) { return c['ID personne'] === p.agent && (c['Rôle'] === 'Jardin et entretien' || c['Rôle'] === 'Ménage') && c['Actif'] === 'Oui'; });
      if (!ok) throw new Error('Intervenant inconnu ou inactif');
    }
    setCell_(x.t, x.r.__row, 'ID agent', p.agent || '');
    if (p.agent && st === 'Ouverte') setCell_(x.t, x.r.__row, 'Statut', 'Planifiée');
    if (!p.agent && st === 'Planifiée') setCell_(x.t, x.r.__row, 'Statut', 'Ouverte');
    journal_(me.id, 'intervention_affectee', 'Interventions', p.id, p.agent || 'désaffectée');
    return { ok: true };
  },
  setIntervention: function (me, p) {
    need_(me, ['Gestion']);
    if (STATUTS_INTERVENTION.indexOf(p.statut) < 0) throw new Error('Statut inconnu');
    var x = rowById_('Interventions', p.id, 'Intervention');
    setCell_(x.t, x.r.__row, 'Statut', p.statut);
    journal_(me.id, 'intervention_statut', 'Interventions', p.id, p.statut);
    if (p.statut === 'Clôturée' && x.r['ID tâche']) {      // la tâche liée suit : une intervention close ne laisse pas de tâche ouverte
      var tt = rowById_('Taches', x.r['ID tâche'], 'Tâche');
      if (tt.r['Statut'] !== 'Faite') { setCell_(tt.t, tt.r.__row, 'Statut', 'Faite'); journal_(me.id, 'tache_statut', 'Taches', tt.r['ID'], 'Faite (intervention clôturée)'); }
    }
    return { ok: true };
  },
  setTache: function (me, p) {
    need_(me, ['Gestion']);
    var o = tacheFields_(p, true), keys = Object.keys(o);
    if (!keys.length) throw new Error('Rien à modifier');
    var x = rowById_('Taches', p.id, 'Tâche');
    keys.forEach(function (k) { setCell_(x.t, x.r.__row, k, o[k]); });
    var autres = keys.filter(function (k) { return k !== 'Statut' && k !== 'Échéance'; });
    if (o['Statut'] !== undefined) journal_(me.id, 'tache_statut', 'Taches', p.id, o['Statut']);
    if (o['Échéance'] !== undefined) journal_(me.id, 'tache_report', 'Taches', p.id, p.echeance);
    if (autres.length) journal_(me.id, 'tache_modifiee', 'Taches', p.id, autres.join(', '));
    return { ok: true };
  },
  addTache: function (me, p) {
    need_(me, ['Gestion']);
    var q = {}; Object.keys(p).forEach(function (k) { q[k] = p[k]; });
    if (q.echeance === undefined || q.echeance === null) q.echeance = today_();
    var o = tacheFields_(q, false);
    var tt = table_('Taches'), tid = nextId_(tt, 'T', 3);
    o['ID'] = tid; if (o['Statut'] === undefined) o['Statut'] = 'À faire';
    if (o['ID responsable'] === undefined) o['ID responsable'] = (p.responsable === undefined ? me.id : '');
    if (o['Priorité'] === undefined) o['Priorité'] = 'Normale';
    o['Créée par'] = me.id; o['Créée le'] = new Date();
    appendRow_(tt, o);
    journal_(me.id, 'tache_creee', 'Taches', tid, o['Titre'].substring(0, 80));
    return { tache: tid };
  },
  addCommentaire: function (me, p) {
    need_(me, ['Gestion']);
    var txt = safe_(p.txt, 2000);
    if (!txt) throw new Error('Écrivez un commentaire');
    rowById_('Taches', p.tache, 'Tâche');
    var tc = table_('Commentaires'), id = nextId_(tc, 'C', 4);
    appendRow_(tc, { 'ID': id, 'ID tâche': p.tache, 'Date': new Date(), 'Auteur': me.id, 'Texte': txt });
    journal_(me.id, 'tache_commentaire', 'Taches', p.tache, txt.substring(0, 80) + (mentionIds_(txt).length ? ' · mentionne ' + mentionIds_(txt).join(', ') : ''));
    return { commentaire: id };
  },
  setIcal: function (me, p) {
    need_(me, ['Gestion']);
    var t = table_('Logements'), l = t.rows.filter(function (x) { return x['ID'] === p.id; })[0];
    if (!l) throw new Error('Logement inconnu');
    if (typeof p.actif !== 'boolean') throw new Error('Valeur invalide');
    if (t.h.indexOf('Synchro iCal active') < 0) {            // colonne créée à la première utilisation
      t.sh.getRange(1, t.h.length + 1).setValue('Synchro iCal active'); t.h.push('Synchro iCal active');
    }
    setCell_(t, l.__row, 'Synchro iCal active', p.actif ? 'Oui' : 'Non');
    journal_(me.id, p.actif ? 'ical_active' : 'ical_desactive', 'Logements', p.id, p.actif ? 'synchronisation réactivée' : 'synchronisation suspendue');
    return { ok: true };
  },
  setConsignes: function (me, p) {
    need_(me, ['Gestion']);
    var t = table_('Logements'), l = t.rows.filter(function (x) { return x['ID'] === p.id; })[0];
    if (!l) throw new Error('Logement inconnu');
    setCell_(t, l.__row, 'Consignes d\'accès', String(p.consignes || '').substring(0, 1000));
    if (p.adresse !== undefined && String(p.adresse).trim()) setCell_(t, l.__row, 'Adresse', String(p.adresse).trim().substring(0, 200));
    journal_(me.id, 'logement_modifie', 'Logements', p.id, 'consignes/adresse');
    return { ok: true };
  },
  startIntervention: function (me, p) {
    need_(me, ['Jardin et entretien']);
    var x = interventionFor_(me, p.id), s = x.r['Statut'];
    if (s === 'En cours') return { statut: s };
    if (s === 'Clôturée') throw new Error('Intervention déjà clôturée');
    setCell_(x.t, x.r.__row, 'Statut', 'En cours');
    journal_(me.id, 'intervention_demarree', 'Interventions', p.id, '');
    return { statut: 'En cours' };
  },
  addInterventionPhoto: function (me, p) {
    need_(me, ['Jardin et entretien']);
    var x = interventionFor_(me, p.id);
    if (x.r['Statut'] !== 'En cours') throw new Error('Démarrez l\'intervention avant d\'ajouter des photos');
    var fid = savePhoto_(p.data, p.id);
    var all = ids_(x.r['Photos (ids Drive)']).concat([fid]);
    setCell_(x.t, x.r.__row, 'Photos (ids Drive)', all.join(',')); setCell_(x.t, x.r.__row, 'Nb photos', all.length);
    journal_(me.id, 'photo_ajoutee', 'Interventions', p.id, fid);
    return { photos: all.length };
  },
  finishIntervention: function (me, p) {
    need_(me, ['Jardin et entretien']);
    var x = interventionFor_(me, p.id), s = x.r['Statut'];
    if (s === 'Clôturée') return { statut: s };
    if (s !== 'En cours') throw new Error('Cette intervention n\'est pas en cours (' + s + ')');
    setCell_(x.t, x.r.__row, 'Statut', 'Clôturée');
    journal_(me.id, 'intervention_terminee', 'Interventions', p.id, String(p.note || '').substring(0, 300));
    closeLinkedTask_(me, x);
    return { statut: 'Clôturée' };
  },
  /** Signalement d'incident : logement, photo, description. Crée l'intervention et une tâche pour la gestion. */
  reportIncident: function (me, p) {
    need_(me, ['Ménage', 'Jardin et entretien', 'Gestion']);
    var desc = String(p.description || '').trim();
    if (!desc) throw new Error('Décrivez l\'incident');
    if (!logementOf_(p.logement)) throw new Error('Logement inconnu');
    if (me.role === 'Ménage' && !myLogements_(me).some(function (l) { return l.id === p.logement; })) throw new Error('Accès refusé');
    var photos = p.photo ? [savePhoto_(p.photo, 'incident')] : [];
    var ti = table_('Interventions'), id = nextId_(ti, 'I', 3);
    appendRow_(ti, { 'ID': id, 'ID logement': p.logement, 'Type': 'Incident', 'Description': desc.substring(0, 1000), 'Nb photos': photos.length,
      'ID agent': '', 'Statut': 'Ouverte', 'Urgence': p.urgent ? 'Haute' : 'Normale', 'ID tâche': '', 'Photos (ids Drive)': photos.join(','), 'Signalé par': me.id });
    var gestion = ref_('Codes').rows.filter(function (c) { return c['Rôle'] === 'Gestion' && c['Actif'] === 'Oui'; })[0];
    var tt = table_('Taches'), tid = nextId_(tt, 'T', 3);
    appendRow_(tt, { 'ID': tid, 'Titre': 'Traiter l\'incident ' + id + ' (' + (logementOf_(p.logement)['Nom']) + ')', 'Échéance': dateOf_(today_()),
      'ID responsable': gestion ? gestion['ID personne'] : '', 'Statut': 'À faire', 'ID logement (facultatif)': p.logement, 'ID intervention': id });
    var ti2 = table_('Interventions'); var r = ti2.rows.filter(function (x) { return x['ID'] === id; })[0];
    setCell_(ti2, r.__row, 'ID tâche', tid);
    journal_(me.id, 'incident_signale', 'Interventions', id, p.logement + ' : ' + desc.substring(0, 80));
    return { intervention: id, tache: tid };
  }
};

function savePhoto_(dataUrl, label) {
  var m = /^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/.exec(String(dataUrl || ''));
  if (!m) throw new Error('Photo invalide (JPEG attendu)');
  var bytes = Utilities.base64Decode(m[1]);
  if (bytes.length > MAX_PHOTO_BYTES) throw new Error('Photo trop lourde');
  var props = PropertiesService.getScriptProperties(), fid = props.getProperty('PHOTO_FOLDER'), folder;
  try { folder = fid ? DriveApp.getFolderById(fid) : null; } catch (e) { folder = null; }
  if (!folder) { folder = DriveApp.createFolder('Cockpit Conciergerie — photos'); props.setProperty('PHOTO_FOLDER', folder.getId()); }
  var name = label + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss') + '.jpg';
  return folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', name)).getId();
}

// ---------- Gestion des codes (depuis le menu du classeur) ----------
function generateCode_(personId) {
  var t = ref_('Codes'), r = t.rows.filter(function (x) { return x['ID personne'] === personId; })[0];
  if (!r) throw new Error('Personne inconnue');
  var alpha = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789', code = '';
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid() + Utilities.getUuid(), Utilities.Charset.UTF_8);
  for (var i = 0; i < 10; i++) code += alpha.charAt((bytes[i] + 256) % alpha.length);
  var h = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'cockpit|' + code, Utilities.Charset.UTF_8));
  setCell_(t, r.__row, 'Empreinte du code (SHA-256)', h); setCell_(t, r.__row, 'Actif', 'Oui');
  setCell_(t, r.__row, 'Créé le', dateOf_(today_())); setCell_(t, r.__row, 'Révoqué le', '');
  journal_(SYSTEM, 'code_genere', 'Codes', personId, '');
  return code;     // affiché une seule fois, jamais stocké en clair
}
function revokeCode_(personId) {
  var t = ref_('Codes'), r = t.rows.filter(function (x) { return x['ID personne'] === personId; })[0];
  if (!r) throw new Error('Personne inconnue');
  setCell_(t, r.__row, 'Actif', 'Non'); setCell_(t, r.__row, 'Révoqué le', dateOf_(today_()));
  journal_(SYSTEM, 'code_revoque', 'Codes', personId, '');
}
function menuGenerateCode() {
  var ui = SpreadsheetApp.getUi(), a = ui.prompt('Générer un code personnel', 'ID de la personne (ex. U-002). L\'ancien code cesse de fonctionner.', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  ui.alert('Nouveau code (à transmettre en main propre, il ne sera plus affiché) :\n\n' + generateCode_(a.getResponseText().trim()));
}
function menuRevokeCode() {
  var ui = SpreadsheetApp.getUi(), a = ui.prompt('Révoquer un code', 'ID de la personne (ex. U-002).', ui.ButtonSet.OK_CANCEL);
  if (a.getSelectedButton() !== ui.Button.OK) return;
  revokeCode_(a.getResponseText().trim()); ui.alert('Code révoqué. Les sessions en cours de cette personne s\'arrêtent.');
}
