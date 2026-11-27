/**
 * Cockpit Conciergerie — étape 2 : réservations (import iCal, saisie manuelle, chevauchement, état de synchro).
 * Lie les onglets par en-têtes (l'ordre des colonnes peut changer, pas leurs noms).
 * Dépend de Ical.gs.
 */
var TZ = 'Europe/Paris';
var SYSTEM = 'SYSTEM';
var CANAUX = ['Airbnb', 'Booking', 'Direct', 'Blocage'];

// ---------- Accès aux onglets ----------
function table_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Onglet introuvable : ' + name);
  var v = sh.getDataRange().getValues(), h = v[0], rows = [];
  for (var i = 1; i < v.length; i++) {
    if (v[i][0] === '') continue;
    var o = { __row: i + 1 };
    h.forEach(function (k, j) { o[k] = v[i][j]; });
    rows.push(o);
  }
  return { sh: sh, h: h, rows: rows };
}
/** Lecture mémorisée pour la durée d'une exécution : réservée aux onglets que l'API ne modifie jamais (Logements, ModelesListe, Codes). */
var MEMO_ = {};
function ref_(name) { return MEMO_[name] || (MEMO_[name] = table_(name)); }
function colLetter_(t, header) {
  var j = t.h.indexOf(header);
  if (j < 0) throw new Error('Colonne introuvable : ' + header);
  var n = j + 1, s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function day_(v) { return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v); }
function dateOf_(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nextId_(t, prefix, pad) {
  var max = 0;
  t.rows.forEach(function (r) { var n = parseInt(String(r['ID']).replace(prefix + '-', ''), 10); if (n > max) max = n; });
  var s = String(max + 1); while (s.length < pad) s = '0' + s;
  return prefix + '-' + s;
}
function appendRow_(t, obj) {
  t.sh.appendRow(t.h.map(function (k) { return obj.hasOwnProperty(k) ? obj[k] : ''; }));
  return t.sh.getLastRow();
}
function setCell_(t, row, header, value) { t.sh.getRange(row, t.h.indexOf(header) + 1).setValue(value); }

// ---------- Journal : chaque écriture est horodatée et attribuée ----------
function journal_(personId, action, tab, rowId, detail) {
  var sh = SpreadsheetApp.getActive().getSheetByName('Journal');
  if (!sh) throw new Error('Onglet introuvable : Journal');
  sh.appendRow([new Date(), personId, action, tab, rowId, detail || '']);
}

// ---------- Codes personnels ----------
/** Retourne la personne si le code est valide et actif, sinon null. L'empreinte = SHA-256 de « cockpit|<code> ». */
function authenticate(code) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'cockpit|' + code, Utilities.Charset.UTF_8);
  var h = hex_(bytes);
  var found = table_('Codes').rows.filter(function (p) {
    return p['Empreinte du code (SHA-256)'] === h && p['Actif'] === 'Oui' && !p['Révoqué le'];
  })[0];
  return found ? { id: found['ID personne'], nom: found['Nom'], role: found['Rôle'] } : null;
}
function requireRole_(personId, roles) {
  var p = table_('Codes').rows.filter(function (r) { return r['ID personne'] === personId && r['Actif'] === 'Oui'; })[0];
  if (!p || roles.indexOf(p['Rôle']) < 0) throw new Error('Accès refusé');
  return p;
}

// ---------- Formules des lignes ajoutées ----------
function reservationFormulas_(t, row) {
  var L = function (h) { return colLetter_(t, h); };
  var B = L('ID logement'), C = L('Arrivée'), D = L('Départ'), S = L('Statut');
  var f = '=IF(' + S + row + '<>"Active",0,COUNTIFS($' + B + '$2:$' + B + '$200,' + B + row + ',$' + S + '$2:$' + S + '$200,"Active",$' + C + '$2:$' + C + '$200,"<"&' + D + row + ',$' + D + '$2:$' + D + '$200,">"&' + C + row + ')-1)';
  t.sh.getRange(row, t.h.indexOf('Chevauchements') + 1).setFormula(f);
  t.sh.getRange(row, t.h.indexOf('Alerte') + 1).setFormula('=IF(' + L('Chevauchements') + row + '>0,"CHEVAUCHEMENT","")');
}
function menageFormulas_(t, row) {
  var L = function (h) { return colLetter_(t, h); };
  t.sh.getRange(row, t.h.indexOf('Date (suit la réservation)') + 1).setFormula(
    '=INDEX(Reservations!$D$2:$D$200,MATCH(' + L('ID réservation') + row + ',Reservations!$A$2:$A$200,0))');
  t.sh.getRange(row, t.h.indexOf('Jour (déduit)') + 1).setFormula(
    '=CHOOSE(WEEKDAY(' + L('Date (suit la réservation)') + row + ',2),"lundi","mardi","mercredi","jeudi","vendredi","samedi","dimanche")');
  t.sh.getRange(row, t.h.indexOf('Alerte') + 1).setFormula(
    '=IF(' + L('Statut') + row + '="À affecter","MÉNAGE NON AFFECTÉ","")');
}

/** Un ménage est créé pour chaque départ (hors blocage). Un seul par réservation. */
function ensureMenage_(resId, logId, personId) {
  var t = table_('Menages');
  if (t.rows.some(function (m) { return m['ID réservation'] === resId; })) return null;
  var id = nextId_(t, 'M', 3);
  var row = appendRow_(t, { 'ID': id, 'ID logement': logId, 'ID réservation': resId, 'Statut': 'À affecter', 'Liste de contrôle': 'Liste standard', 'Nb photos': 0 });
  menageFormulas_(t, row);
  journal_(personId, 'menage_cree', 'Menages', id, 'départ ' + resId);
  return id;
}
function cancelMenage_(resId, personId) {
  var t = table_('Menages');
  t.rows.forEach(function (m) {
    if (m['ID réservation'] === resId && m['Statut'] !== 'Terminé' && m['Statut'] !== 'En cours' && m['Statut'] !== 'Annulé') {
      setCell_(t, m.__row, 'Statut', 'Annulé');
      journal_(personId, 'menage_annule', 'Menages', m['ID'], 'réservation ' + resId + ' annulée');
    }
  });
}
function reviveMenage_(resId, logId, personId) {
  var t = table_('Menages');
  var m = t.rows.filter(function (x) { return x['ID réservation'] === resId; })[0];
  if (m && m['Statut'] === 'Annulé') { setCell_(t, m.__row, 'Statut', 'À affecter'); journal_(personId, 'menage_rouvert', 'Menages', m['ID'], 'réservation ' + resId + ' réactivée'); }
  else if (!m) ensureMenage_(resId, logId, personId);
}

// ---------- Import iCal ----------
function syncAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return [{ statut: 'occupé' }];
  try {
    var report = table_('Logements').rows.map(syncLogement_);
    SpreadsheetApp.flush();
    return report;
  } finally { lock.releaseLock(); }
}

/** Un lien iCal est actif sauf si la colonne « Synchro iCal active » vaut « Non ». */
function icalOff_(log) { return String(log['Synchro iCal active'] == null ? '' : log['Synchro iCal active']).trim().toLowerCase() === 'non'; }

function syncLogement_(log) {
  var id = log['ID'], url = String(log['Lien iCal'] || '');
  if (icalOff_(log)) return { logement: id, statut: 'ignoré (iCal désactivé)' };
  if (!/^https?:\/\//.test(url) || url.indexOf('exemple.invalid') > -1) return { logement: id, statut: 'ignoré (pas de lien iCal)' };
  var text;
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    text = resp.getContentText();
    if (resp.getResponseCode() !== 200 || text.indexOf('BEGIN:VCALENDAR') < 0) throw new Error('HTTP ' + resp.getResponseCode());
  } catch (e) {
    // La synchro n'est PAS horodatée : l'alerte « synchro en panne depuis X heures » finit par se déclencher.
    journal_(SYSTEM, 'sync_echec', 'Logements', id, String(e));
    return { logement: id, statut: 'échec : ' + e };
  }
  var events = parseIcal_(text), today = today_();
  var tr = table_('Reservations');
  var existing = tr.rows.filter(function (r) { return r['ID logement'] === id && r['Source'] === 'iCal'; })
    .map(function (r) { return { id: r['ID'], uid: r['UID iCal'], start: day_(r['Arrivée']), end: day_(r['Départ']), status: r['Statut'], row: r.__row }; });
  var d = diffEvents_(existing, events, today);
  var byUid = {}; existing.forEach(function (r) { byUid[r.uid] = r; });
  var rowsAll = tr; var n = { add: 0, update: 0, cancel: 0, reactivate: 0 };

  d.add.forEach(function (e) {
    rowsAll = table_('Reservations');                      // recharge pour obtenir le bon prochain identifiant
    var rid = nextId_(rowsAll, 'R', 3), canal = channelOf_(url, e.summary);
    var row = appendRow_(rowsAll, { 'ID': rid, 'ID logement': id, 'Arrivée': dateOf_(e.start), 'Départ': dateOf_(e.end), 'Canal': canal,
      'Voyageur (saisi à la main)': '', 'Caution (statut)': '', 'Source': 'iCal', 'UID iCal': e.uid, 'Statut': 'Active' });
    reservationFormulas_(rowsAll, row);
    if (canal !== 'Blocage') ensureMenage_(rid, id, SYSTEM);
    journal_(SYSTEM, 'reservation_importee', 'Reservations', rid, canal + ' ' + e.start + ' → ' + e.end);
    n.add++;
  });
  d.update.forEach(function (u) {
    var r = byUid[u.event.uid];
    setCell_(tr, r.row, 'Arrivée', dateOf_(u.event.start)); setCell_(tr, r.row, 'Départ', dateOf_(u.event.end));
    journal_(SYSTEM, 'reservation_modifiee', 'Reservations', r.id, u.before.start + '→' + u.before.end + ' devient ' + u.event.start + '→' + u.event.end);
    n.update++;
  });
  d.cancel.forEach(function (r) {
    setCell_(tr, r.row, 'Statut', 'Annulée'); cancelMenage_(r.id, SYSTEM);
    journal_(SYSTEM, 'reservation_annulee', 'Reservations', r.id, 'absente du flux iCal'); n.cancel++;
  });
  d.reactivate.forEach(function (e) {
    var r = byUid[e.uid];
    setCell_(tr, r.row, 'Statut', 'Active'); reviveMenage_(r.id, id, SYSTEM);
    journal_(SYSTEM, 'reservation_reactivee', 'Reservations', r.id, 'de retour dans le flux iCal'); n.reactivate++;
  });

  var tl = table_('Logements');
  setCell_(tl, log.__row, 'Dernière synchro iCal', new Date());
  return { logement: id, statut: 'ok', ajouts: n.add, modifiees: n.update, annulees: n.cancel, reactivees: n.reactivate };
}

// ---------- Saisie manuelle (réservations directes, blocages propriétaire) ----------
/**
 * p = { logement:'L-001', arrivee:'2026-11-12', depart:'2026-11-16', canal:'Direct'|'Blocage'|..., voyageur:'', caution:'À contrôler'|'Libérée'|'' }
 * Un chevauchement n'empêche pas la saisie (la réalité prime) : il est signalé dans le retour et par l'alerte du classeur.
 */
function addManualReservation(personId, p) {
  requireRole_(personId, ['Gestion']);
  var logs = table_('Logements').rows;
  if (!logs.some(function (l) { return l['ID'] === p.logement; })) throw new Error('Logement inconnu : ' + p.logement);
  if (CANAUX.indexOf(p.canal) < 0) throw new Error('Canal invalide : ' + p.canal);
  if (!p.arrivee || !p.depart || p.arrivee >= p.depart) throw new Error('Le départ doit être après l\'arrivée');
  var t = table_('Reservations');
  var id = nextId_(t, 'R', 3);
  var row = appendRow_(t, { 'ID': id, 'ID logement': p.logement, 'Arrivée': dateOf_(p.arrivee), 'Départ': dateOf_(p.depart), 'Canal': p.canal,
    'Voyageur (saisi à la main)': p.canal === 'Blocage' ? '' : (p.voyageur || ''), 'Caution (statut)': p.canal === 'Blocage' ? '' : (p.caution || ''),
    'Source': 'Manuelle', 'UID iCal': '', 'Statut': 'Active' });
  reservationFormulas_(t, row);
  if (p.canal !== 'Blocage') ensureMenage_(id, p.logement, personId);
  journal_(personId, 'reservation_creee', 'Reservations', id, p.canal + ' ' + p.arrivee + ' → ' + p.depart);
  var conflicts = alerts().overlaps.filter(function (pr) { return pr.indexOf(id) > -1; });
  return { id: id, chevauchement: conflicts.length ? conflicts[0] : null };
}

/** Changer les dates d'une réservation manuelle ; le ménage suit (sa date est lue dans la réservation). */
function updateManualReservationDates(personId, resId, arrivee, depart) {
  requireRole_(personId, ['Gestion']);
  var t = table_('Reservations');
  var r = t.rows.filter(function (x) { return x['ID'] === resId; })[0];
  if (!r) throw new Error('Réservation inconnue');
  if (r['Source'] !== 'Manuelle') throw new Error('Une réservation iCal se modifie dans la plateforme d\'origine');
  if (arrivee >= depart) throw new Error('Le départ doit être après l\'arrivée');
  setCell_(t, r.__row, 'Arrivée', dateOf_(arrivee)); setCell_(t, r.__row, 'Départ', dateOf_(depart));
  journal_(personId, 'reservation_modifiee', 'Reservations', resId, arrivee + ' → ' + depart);
}

// ---------- Alertes (consommées par Ma journée, le calendrier et la fiche logement) ----------
function alerts() {
  var res = table_('Reservations').rows.map(function (r) {
    return { id: r['ID'], logement: r['ID logement'], start: day_(r['Arrivée']), end: day_(r['Départ']), status: r['Statut'] };
  });
  var seuil = Number(SpreadsheetApp.getActive().getSheetByName('Parametres').getRange('B4').getValue()) || 12;
  var now = Date.now();
  var stale = table_('Logements').rows.filter(function (l) { return l['Lien iCal'] && !icalOff_(l) && l['Dernière synchro iCal']; }).map(function (l) {
    return { logement: l['ID'], heures: Math.round(hoursSince_(new Date(l['Dernière synchro iCal']).getTime(), now) * 10) / 10 };
  }).filter(function (x) { return x.heures > seuil; });
  var unassigned = table_('Menages').rows.filter(function (m) { return m['Statut'] === 'À affecter'; }).map(function (m) { return m['ID']; });
  return { overlaps: findOverlaps_(res), staleSyncs: stale, unassignedMenages: unassigned };
}
function showAlerts() {
  var a = alerts();
  SpreadsheetApp.getUi().alert('Chevauchements : ' + (a.overlaps.map(function (p) { return p.join(' / '); }).join(', ') || 'aucun') +
    '\nSynchros en panne : ' + (a.staleSyncs.map(function (s) { return s.logement + ' (' + s.heures + ' h)'; }).join(', ') || 'aucune') +
    '\nMénages non affectés : ' + (a.unassignedMenages.join(', ') || 'aucun'));
}

// ---------- Menu et déclencheur ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Conciergerie')
    .addItem('Synchroniser les iCal maintenant', 'syncAll')
    .addItem('Installer la synchro automatique (toutes les heures)', 'installTrigger')
    .addItem('Afficher les alertes', 'showAlerts')
    .addSeparator()
    .addItem('Générer un code personnel', 'menuGenerateCode')
    .addItem('Révoquer un code', 'menuRevokeCode').addToUi();
}
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
}
