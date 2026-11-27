/**
 * Logique pure (sans Google) : lecture iCal, canal, différences, chevauchements, empreinte.
 * Testable hors Google : voir le test Node livré à part.
 * Les dates sont des chaînes « AAAA-MM-JJ » (jour civil), jamais des Date, pour éviter les décalages de fuseau.
 */

function unfold_(text) {
  return String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n[ \t]/g, '');
}

function icalDay_(value) {
  var m = /^(\d{4})(\d{2})(\d{2})/.exec(String(value).trim());
  return m ? m[1] + '-' + m[2] + '-' + m[3] : null;
}

/** Lit un flux iCal. Retourne [{uid, start, end, summary, status}] ; ignore les événements sans UID ou sans dates. */
function parseIcal_(text) {
  var lines = unfold_(text).split('\n');
  var events = [], cur = null;
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (line === 'BEGIN:VEVENT') { cur = {}; continue; }
    if (line === 'END:VEVENT') {
      if (cur && cur.uid && cur.start && cur.end) events.push(cur);
      cur = null; continue;
    }
    if (!cur) continue;
    var idx = line.indexOf(':');
    if (idx < 0) continue;
    var name = line.substring(0, idx).split(';')[0].toUpperCase();
    var val = line.substring(idx + 1);
    if (name === 'UID') cur.uid = val.trim();
    else if (name === 'DTSTART') cur.start = icalDay_(val);
    else if (name === 'DTEND') cur.end = icalDay_(val);
    else if (name === 'SUMMARY') cur.summary = val.trim();
    else if (name === 'STATUS') cur.status = val.trim().toUpperCase();
  }
  return events.filter(function (e) { return e.status !== 'CANCELLED' && e.start < e.end; });
}

/** Canal d'après l'adresse du flux. Airbnb signale un blocage par « Not available » ; Booking exporte tout pareil. */
function channelOf_(url, summary) {
  var u = String(url).toLowerCase(), s = String(summary || '');
  if (u.indexOf('airbnb') > -1) return /not available|indisponible|blocked/i.test(s) ? 'Blocage' : 'Airbnb';
  if (u.indexOf('booking.com') > -1) return 'Booking';
  return /not available|indisponible|blocked/i.test(s) ? 'Blocage' : 'Direct';
}

/**
 * existing : réservations iCal du logement [{uid, start, end, status}] ; events : flux lu ; today : « AAAA-MM-JJ ».
 * - nouvelle UID -> ajout ; dates changées -> mise à jour (le ménage suit) ;
 * - UID disparue du flux et départ pas encore passé -> annulation (une réservation passée qui sort du flux est laissée telle quelle).
 * - une réservation annulée qui revient dans le flux est réactivée.
 */
function diffEvents_(existing, events, today) {
  var byUid = {}, seen = {}, out = { add: [], update: [], cancel: [], reactivate: [] };
  existing.forEach(function (r) { byUid[r.uid] = r; });
  events.forEach(function (e) {
    seen[e.uid] = true;
    var r = byUid[e.uid];
    if (!r) out.add.push(e);
    else {
      if (r.status === 'Annulée') out.reactivate.push(e);
      if (r.start !== e.start || r.end !== e.end) out.update.push({ event: e, before: r });
    }
  });
  existing.forEach(function (r) {
    if (!seen[r.uid] && r.status === 'Active' && r.end > today) out.cancel.push(r);
  });
  return out;
}

/** Deux réservations actives du même logement se chevauchent si l'une arrive avant le départ de l'autre (départ = arrivée suivante : pas de conflit). */
function findOverlaps_(reservations) {
  var pairs = [];
  var act = reservations.filter(function (r) { return r.status === 'Active'; });
  for (var i = 0; i < act.length; i++)
    for (var j = i + 1; j < act.length; j++) {
      var a = act[i], b = act[j];
      if (a.logement === b.logement && a.start < b.end && b.start < a.end) pairs.push([a.id, b.id]);
    }
  return pairs;
}

function hoursSince_(lastSyncMs, nowMs) { return (nowMs - lastSyncMs) / 3600000; }

function hex_(bytes) {
  return bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length < 2 ? '0' + v : v; }).join('');
}

if (typeof module !== 'undefined') module.exports = { parseIcal_: parseIcal_, channelOf_: channelOf_, diffEvents_: diffEvents_, findOverlaps_: findOverlaps_, hoursSince_: hoursSince_, hex_: hex_ };
