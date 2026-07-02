/* Fantasy Cycling scoring engine — reads riders CSV + results XLSX at runtime
   and derives the same data structures the dashboards render. No dependencies. */
(function () {
  'use strict';

  /* ---------- XLSX reader (zip + DEFLATE via DecompressionStream) ---------- */
  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const w = ds.writable.getWriter(); w.write(bytes); w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
  }
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
  async function unzip(buf) {
    const b = new Uint8Array(buf);
    let eo = -1;
    for (let i = b.length - 22; i >= 0; i--) { if (u32(b, i) === 0x06054b50) { eo = i; break; } }
    if (eo < 0) throw new Error('not a zip');
    const cdOff = u32(b, eo + 16), cdCnt = u16(b, eo + 10);
    const files = {}; let p = cdOff;
    for (let i = 0; i < cdCnt; i++) {
      const comp = u16(b, p + 10), csize = u32(b, p + 20), nlen = u16(b, p + 28), elen = u16(b, p + 30), clen = u16(b, p + 32), lho = u32(b, p + 42);
      const name = new TextDecoder().decode(b.subarray(p + 46, p + 46 + nlen));
      const lnlen = u16(b, lho + 26), lelen = u16(b, lho + 28);
      const dstart = lho + 30 + lnlen + lelen;
      files[name] = { comp, raw: b.subarray(dstart, dstart + csize) };
      p += 46 + nlen + elen + clen;
    }
    const out = {};
    for (const n of Object.keys(files)) { const f = files[n]; out[n] = f.comp === 0 ? f.raw.slice() : await inflateRaw(f.raw); }
    return out;
  }
  const dec = u => new TextDecoder().decode(u);
  function parseShared(xml) {
    const arr = []; const re = /<si>([\s\S]*?)<\/si>/g; let m;
    while ((m = re.exec(xml))) {
      const t = m[1].replace(/<rPh[\s\S]*?<\/rPh>/g, '');
      const ts = [...t.matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(x => x[1]);
      arr.push(ts.join('').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'").replace(/&quot;/g, '"'));
    }
    return arr;
  }
  const colNum = ref => { let n = 0; for (const c of ref.replace(/[0-9]/g, '')) n = n * 26 + (c.charCodeAt(0) - 64); return n - 1; };
  function parseSheet(xml, shared) {
    const rows = []; const rowRe = /<row[^>]*?>([\s\S]*?)<\/row>/g; let rm;
    while ((rm = rowRe.exec(xml))) {
      const cells = []; const cRe = /<c[^>]*?r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let cm;
      while ((cm = cRe.exec(rm[1]))) {
        const ci = colNum(cm[1]); const attrs = cm[2] || '', inner = cm[3] || '';
        const tM = /t="([^"]+)"/.exec(attrs); const t = tM ? tM[1] : '';
        let val = '';
        if (t === 'inlineStr') { const im = /<t[^>]*>([\s\S]*?)<\/t>/.exec(inner); val = im ? im[1] : ''; }
        else { const v = /<v>([\s\S]*?)<\/v>/.exec(inner); val = t === 's' ? (v ? shared[+v[1]] : '') : (v ? v[1] : ''); }
        cells[ci] = (val || '').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&#39;/g, "'").replace(/&quot;/g, '"');
      }
      rows.push(cells);
    }
    return rows;
  }
  function readXlsxBuffer(files) {
    const shared = files['xl/sharedStrings.xml'] ? parseShared(dec(files['xl/sharedStrings.xml'])) : [];
    const sn = Object.keys(files).find(n => /xl\/worksheets\/sheet1\.xml$/.test(n));
    return parseSheet(dec(files[sn]), shared);
  }
  async function xlsxRows(arrayBuffer) { return readXlsxBuffer(await unzip(arrayBuffer)); }

  function parseCsv(text) {
    const rows = []; let i = 0, f = '', row = [], q = false;
    const pf = () => { row.push(f); f = ''; }; const pr = () => { rows.push(row); row = []; };
    while (i < text.length) {
      const c = text[i];
      if (q) { if (c === '"') { if (text[i + 1] === '"') { f += '"'; i += 2; continue; } q = false; i++; continue; } f += c; i++; continue; }
      if (c === '"') { q = true; i++; continue; }
      if (c === ',') { pf(); i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { pf(); pr(); i++; continue; }
      f += c; i++;
    }
    if (f.length || row.length) { pf(); pr(); }
    return rows;
  }
  function rowsToObjs(rows) {
    if (!rows || !rows.length) return [];
    const h = rows[0].map(x => (x || '').trim());
    return rows.slice(1).filter(r => r.length && r.some(x => x !== '' && x != null))
      .map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] !== undefined ? r[i] : ''); return o; });
  }

  /* ---------- shared helpers ---------- */
  const norm = s => (typeof s !== 'string' ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/-/g, ' ').trim();
  const serialToDate = v => new Date(Date.UTC(1899, 11, 30) + (+v) * 86400000);
  const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const MON_IDX = {}; MON.forEach((m, i) => MON_IDX[m] = i);
  // Returns a UTC-midnight epoch (ms) for Excel serials, ISO, M/D/YYYY, or "Mon D[ YYYY]".
  function parseDate(s) {
    if (s == null || s === '') return null;
    s = String(s).trim();
    if (/^\d+(\.\d+)?$/.test(s)) return +serialToDate(s);
    let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3]);
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return Date.UTC(+m[3], +m[1] - 1, +m[2]);
    m = s.match(/^([A-Za-z]{3})[a-z]*\s+(\d{1,2})(?:\s+(\d{4}))?$/);
    if (m && MON_IDX[m[1]] != null) return Date.UTC(m[3] ? +m[3] : 2026, MON_IDX[m[1]], +m[2]);
    const d = new Date(s);
    return isNaN(d) ? null : +d;
  }
  const r1 = x => Math.round(x * 10) / 10;
  const fmtDate = serial => { const d = serialToDate(serial); return MON[d.getUTCMonth()] + ' ' + d.getUTCDate(); };
  const fmtMs = ms => { const d = new Date(ms); return MON[d.getUTCMonth()] + ' ' + d.getUTCDate(); };

  /* ---------- Grand Tour scoring ---------- */
  const GT_SCORING = {
    'GC Standing': { 1: 40, 2: 35, 3: 30, 4: 25, 5: 20, 6: 18, 7: 16, 8: 14, 9: 12, 10: 10 },
    'Stage Result': { 1: 12, 2: 10, 3: 8, 4: 6, 5: 5, 6: 4, 7: 3, 8: 2, 9: 1, 10: 1 },
    'Jersey': { 1: 12, 2: 8, 3: 4 }
  };
  // TdF-only stage scoring (adopted for 2026 TdF only — Giro/Vuelta keep the default above)
  // 12 places deep, per league ruling.
  const TDF_STAGE_SCORING = { 1: 12, 2: 10, 3: 9, 4: 8, 5: 7, 6: 6, 7: 5, 8: 4, 9: 3, 10: 2, 11: 1, 12: 1 };
  const REPL = { 1: 1, 2: 1, 3: 1, 4: .9, 5: .9, 6: .8, 7: .8, 8: .7, 9: .7, 10: .6, 11: .6, 12: .6, 13: .5, 14: .5, 15: .5, 16: .5, 17: .5, 18: .5, 19: .5, 20: .5 };
  // Replacement-penalty cutoff: subs added ON OR AFTER this date score reduced GC + zero
  // jersey. Per-race (passed via cfg.repcut); this is only the fallback default.
  const DEFAULT_REPCUT = Date.UTC(2026, 4, 16);
  const STAGE_COLS = ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th', '10th', '11th', '12th'];

  function computeGrandTour(ridersCsvText, resultsRows, cfg) {
    cfg = cfg || {};
    const REPCUT = cfg.repcut != null ? cfg.repcut : DEFAULT_REPCUT;
    // riders
    const riders = rowsToObjs(parseCsv(ridersCsvText)).map((r, idx) => ({
      rider_name: r.rider_name, owner: r.owner, add_date: r.add_date,
      drop_date: (r.drop_date || '').trim() || null,
      is_replacement: ['true', '1', 'yes'].includes(String(r.is_replacement).trim().toLowerCase()),
      rider_role: r.rider_role || 'N/A', replaces_rider: (r.replaces_rider || '').trim() || '',
      match_name: norm(r.rider_name), _i: idx
    }));
    const owners = [...new Set(riders.map(r => r.owner))];
    // team_pick: draft order by add_date (stable), replacements inherit replaced rider's slot
    const sorted = riders.slice().sort((a, b) => (parseDate(a.add_date) - parseDate(b.add_date)) || a._i - b._i);
    const cum = {}, pickMap = {};
    for (const r of sorted) {
      const ow = r.owner.toLowerCase().trim(), rd = r.rider_name.toLowerCase().trim();
      if (!r.is_replacement) { cum[ow] = (cum[ow] || 0) + 1; r.team_pick = cum[ow]; pickMap[ow + '|' + rd] = cum[ow]; }
      else { const rep = r.replaces_rider.toLowerCase().trim(); r.team_pick = pickMap[ow + '|' + rep] != null ? pickMap[ow + '|' + rep] : 99; pickMap[ow + '|' + rd] = r.team_pick; }
    }

    // raw results -> flat rows
    const resObjs = rowsToObjs(resultsRows);
    const raw = [], stageMeta = [];
    for (const sd of resObjs) {
      const s = +sd.Stage; if (!s || !sd['1st'] || String(sd['1st']).trim() === '') continue;
      stageMeta.push({ s, date: sd.Date, winner: sd['1st'] });
      STAGE_COLS.forEach((c, i) => { if (sd[c] && String(sd[c]).trim() !== '') raw.push({ Stage: s, res_rider: sd[c], rank: i + 1, Category: 'Stage Result', date: sd.Date }); });
      for (let i = 1; i <= 10; i++) { const c = 'GC #' + i; if (sd[c] && String(sd[c]).trim() !== '') raw.push({ Stage: s, res_rider: sd[c], rank: i, Category: 'GC Standing', date: sd.Date }); }
      [['Points #', 'Points Jersey'], ['Mountain #', 'Mountain Jersey'], ['Youth #', 'Youth Jersey']].forEach(([pre, cat]) => {
        for (let i = 1; i <= 3; i++) { const c = pre + i; if (sd[c] && String(sd[c]).trim() !== '') raw.push({ Stage: s, res_rider: sd[c], rank: i, Category: cat, date: sd.Date }); }
      });
    }
    raw.forEach(r => r.match_name = norm(r.res_rider));
    const allStages = [...new Set(stageMeta.map(x => x.s))].sort((a, b) => a - b);
    const latest = allStages.length ? Math.max(...allStages) : null;

    const byMatch = {}; riders.forEach(r => { (byMatch[r.match_name] = byMatch[r.match_name] || []).push(r); });
    const STAGE_SCORE = cfg.stageScoring || GT_SCORING['Stage Result'];
    const calc = (cat, rank, r) => {
      const tbl = cat === 'Stage Result' ? STAGE_SCORE : (GT_SCORING[cat] || GT_SCORING.Jersey); const base = tbl[rank] || 0; const isJ = cat.indexOf('Jersey') >= 0;
      if (r.is_replacement && parseDate(r.add_date) >= REPCUT) {
        if (cat === 'Stage Result') return base;
        if (isJ) return 0;
        if (cat === 'GC Standing') return base * (REPL[r.team_pick] != null ? REPL[r.team_pick] : 0.5);
      }
      return base;
    };
    // full proc (validity-filtered, owner-attributed)
    const procFull = [];
    for (const e of raw) {
      const ms = byMatch[e.match_name]; if (!ms) continue; const sdate = parseDate(e.date);
      for (const r of ms) {
        const ad = parseDate(r.add_date), dd = r.drop_date ? parseDate(r.drop_date) : null;
        if (sdate >= ad && (dd == null || sdate <= dd)) {
          procFull.push({ Stage: e.Stage, Category: e.Category, rank: e.rank, owner: r.owner, rider_name: r.rider_name, team_pick: r.team_pick, role: r.rider_role, add_date: r.add_date, drop_date: r.drop_date, pts: calc(e.Category, e.rank, r), dispCat: e.Category.indexOf('Jersey') >= 0 ? 'Jerseys' : e.Category });
        }
      }
    }
    // snapshot: GC + jersey only count at the latest stage; stage results accumulate
    const proc = procFull.filter(p => p.Category === 'Stage Result' || p.Stage === latest);

    // totals + meta
    const tot = {}; owners.forEach(o => tot[o] = 0); proc.forEach(p => tot[p.owner] += p.pts);
    const dT = r1(tot.Daniel || 0), tT = r1(tot.Tanner || 0);
    const lastMeta = latest != null ? stageMeta.find(x => x.s === latest) : null;
    const meta = { race: cfg.race || '', year: cfg.year || 2026, stagesDone: latest || 0, lastDate: lastMeta ? fmtDate(lastMeta.date) : 'Pre-race', daniel: dT, tanner: tT, leader: dT >= tT ? 'Daniel' : 'Tanner', lead: Math.abs(r1(dT - tT)) };

    // trajectory + duels
    const stageRes = {}, snap = {}; allStages.forEach(s => { stageRes[s] = { Daniel: 0, Tanner: 0 }; snap[s] = { Daniel: 0, Tanner: 0 }; });
    procFull.forEach(p => { if (p.Category === 'Stage Result') stageRes[p.Stage][p.owner] += p.pts; else snap[p.Stage][p.owner] += p.pts; });
    const trajectory = [], duels = []; let cD = 0, cT = 0;
    for (const s of allStages) {
      cD += stageRes[s].Daniel; cT += stageRes[s].Tanner;
      trajectory.push({ s, d: r1(cD + snap[s].Daniel), t: r1(cT + snap[s].Tanner) });
      duels.push({ s, d: r1(stageRes[s].Daniel), t: r1(stageRes[s].Tanner) });
    }

    // active-status per (name|owner): a stint with no drop_date counts as currently
    // rostered only if no later replacement row targets that rider (handles re-added
    // riders who were implicitly dropped without a drop_date being filled in).
    const replTargets = {}; // owner|loweredName -> latest replacement add_date
    riders.forEach(r => {
      if (r.is_replacement && r.replaces_rider) {
        const key = r.owner + '|' + r.replaces_rider.toLowerCase().trim();
        const ad = parseDate(r.add_date);
        if (replTargets[key] == null || ad > replTargets[key]) replTargets[key] = ad;
      }
    });
    const activeNO = {};
    riders.forEach(r => {
      if (r.drop_date) return;
      const repAfter = replTargets[r.owner + '|' + r.rider_name.toLowerCase().trim()];
      if (repAfter != null && repAfter > parseDate(r.add_date)) return; // replaced later -> not active
      activeNO[r.rider_name + '|' + r.owner] = true;
    });

    // leaderboard: merge stints by name|owner|role
    const lbMap = {}, lbOrder = [];
    for (const p of proc) {
      const k = p.rider_name + '|' + p.owner + '|' + p.role;
      let o = lbMap[k]; if (!o) { o = lbMap[k] = { name: p.rider_name, owner: p.owner, role: p.role, sr: 0, gc: 0, jer: 0 }; lbOrder.push(k); }
      if (p.dispCat === 'Stage Result') o.sr += p.pts; else if (p.dispCat === 'GC Standing') o.gc += p.pts; else o.jer += p.pts;
    }
    const leaderboard = lbOrder.map(k => { const o = lbMap[k]; return { name: o.name, owner: o.owner, role: o.role, sr: r1(o.sr), gc: r1(o.gc), jer: r1(o.jer), total: r1(o.sr + o.gc + o.jer), dropped: !activeNO[o.name + '|' + o.owner] }; })
      .sort((a, b) => b.total - a.total);

    // rosters: per stint pts, ordered by slot then base-before-subs then add_date
    const stintPts = {}; proc.forEach(p => { const k = p.rider_name + '|' + p.owner + '|' + (p.add_date || '') + '|' + (p.drop_date || ''); stintPts[k] = (stintPts[k] || 0) + p.pts; });
    const rostersByOwner = {}; owners.forEach(o => rostersByOwner[o] = []);
    for (const r of riders) {
      const k = r.rider_name + '|' + r.owner + '|' + (r.add_date || '') + '|' + (r.drop_date || '');
      rostersByOwner[r.owner].push({ slot: r.team_pick, rider: r.rider_name, role: r.rider_role, pts: r1(stintPts[k] || 0), isRep: r.is_replacement, replaces: r.replaces_rider, dropped: !!r.drop_date, add: r.add_date, _i: r._i });
    }
    owners.forEach(o => rostersByOwner[o].sort((a, b) => (a.slot - b.slot) || (a.isRep - b.isRep) || (parseDate(a.add) - parseDate(b.add)) || (a._i - b._i)).forEach(x => delete x._i));

    // topPerf: group by name+owner, top 10 per owner
    const np = {}; owners.forEach(o => np[o] = {});
    proc.forEach(p => { const o = np[p.owner]; const g = o[p.rider_name] || (o[p.rider_name] = { name: p.rider_name, pts: 0, dropped: false }); g.pts += p.pts; if (p.drop_date) g.dropped = true; });
    const topPerf = {}; owners.forEach(o => topPerf[o] = Object.values(np[o]).map(x => ({ name: x.name, pts: r1(x.pts), dropped: !activeNO[x.name + '|' + o] })).sort((a, b) => b.pts - a.pts).slice(0, 10));

    // free agents: unowned riders, same snapshot rule, base scoring
    const owned = new Set(riders.map(r => r.match_name));
    const faMap = {};
    raw.filter(e => (e.Category === 'Stage Result' || e.Stage === latest) && !owned.has(e.match_name))
      .forEach(e => { const tbl = e.Category === 'Stage Result' ? STAGE_SCORE : (GT_SCORING[e.Category] || GT_SCORING.Jersey); faMap[e.res_rider] = (faMap[e.res_rider] || 0) + (tbl[e.rank] || 0); });
    const freeAgents = Object.keys(faMap).map(name => ({ name, pts: r1(faMap[name]) })).filter(x => x.pts > 0).sort((a, b) => b.pts - a.pts);

    // stages
    const ownerAtStage = (mn, sdate) => { const ms = byMatch[mn]; if (!ms) return null; for (const r of ms) { const ad = parseDate(r.add_date), dd = r.drop_date ? parseDate(r.drop_date) : null; if (sdate >= ad && (dd == null || sdate <= dd)) return r.owner; } return null; };
    const stages = stageMeta.map(sm => ({ s: sm.s, date: fmtDate(sm.date), winner: sm.winner, winnerOwner: ownerAtStage(norm(sm.winner), parseDate(sm.date)) }));

    // breakdown keyed name|owner
    const breakdown = {};
    proc.forEach(p => { const k = p.rider_name + '|' + p.owner; (breakdown[k] = breakdown[k] || []).push({ stage: p.Stage, cat: p.Category, rank: p.rank, pts: p.pts }); });

    return { meta, trajectory, duels, leaderboard, rostersByOwner, topPerf, freeAgents, stages, breakdown };
  }

  /* ---------- Annual league scoring (tier-based) ---------- */
  const ANNUAL_TIER = {
    'Tier 1': { 1: 30, 2: 27, 3: 24, 4: 21, 5: 18, 6: 15, 7: 12, 8: 9, 9: 6, 10: 3 },
    'Tier 2': { 1: 20, 2: 18, 3: 16, 4: 14, 5: 12, 6: 10, 7: 8, 8: 6, 9: 4, 10: 2 },
    'Tier 3': { 1: 10, 2: 9, 3: 8, 4: 7, 5: 6, 6: 5, 7: 4, 8: 3, 9: 2, 10: 1 }
  };
  const FAR_FUTURE = Date.UTC(2026, 11, 31);

  function computeAnnual(ridersCsvText, scheduleCsvText, resultsRows) {
    // riders
    const riders = rowsToObjs(parseCsv(ridersCsvText)).map((r, i) => ({
      rider_name: r.rider_name, owner: r.owner,
      add: parseDate(r.add_date) != null ? parseDate(r.add_date) : Date.UTC(2026, 0, 1),
      drop: parseDate(r.drop_date),
      replFor: (r.replacement_for || '').trim(),
      match: norm(r.rider_name), _i: i
    }));
    const owners = [...new Set(riders.map(r => r.owner))];
    const cum = {}; riders.forEach(r => { if (!r.replFor) { cum[r.owner] = (cum[r.owner] || 0) + 1; r.slot = cum[r.owner]; } });
    const ownerByMatch = {}; riders.forEach(r => { if (ownerByMatch[r.match] == null) ownerByMatch[r.match] = r.owner; });
    const baseByON = {}; riders.forEach(r => { if (!r.replFor) baseByON[r.owner + '|' + r.match] = r; });
    const effSlot = r => {
      let cur = r, g = 0;
      while (cur.replFor && g++ < 20) {
        const nb = baseByON[cur.owner + '|' + norm(cur.replFor)] || riders.find(x => x.owner === cur.owner && norm(x.rider_name) === norm(cur.replFor));
        if (!nb) break; if (!nb.replFor) return nb.slot; cur = nb;
      }
      return r.slot;
    };
    riders.forEach(r => r.effSlot = r.replFor ? effSlot(r) : r.slot);

    // schedule
    const schedRows = rowsToObjs(parseCsv(scheduleCsvText));
    const schedInfo = {}; schedRows.forEach(s => schedInfo[(s.race_name || '').trim()] = { tier: (s.tier || '').trim(), type: (s.race_type || '').trim(), date: (s.date || '').trim() });

    // results -> field
    const objs = rowsToObjs(resultsRows);
    const field = [];
    for (const o of objs) {
      const race = (o['Race Name'] || '').trim(); if (!race || !o['1st'] || String(o['1st']).trim() === '') continue;
      const dms = parseDate(o.Date); const tier = (schedInfo[race] || {}).tier;
      STAGE_COLS.forEach((c, i) => { const raw = (o[c] || '').trim(); if (!raw) return; field.push({ date: dms, race, stage: (o.Stage || '').trim(), rider: raw, match: norm(raw), place: i + 1, tier, pts: (ANNUAL_TIER[tier] || {})[i + 1] || 0 }); });
    }
    if (!field.length) return null;

    // ownership window
    const rosterByMatch = {}; riders.forEach(r => { (rosterByMatch[r.match] = rosterByMatch[r.match] || []).push(r); });
    const ownerStint = (match, dms) => { const list = rosterByMatch[match]; if (!list) return null; for (const r of list) { const dr = r.drop != null ? r.drop : FAR_FUTURE; if (dms >= r.add && dms <= dr) return r; } return null; };
    const proc = []; field.forEach(f => { const st = ownerStint(f.match, f.date); if (st) proc.push(Object.assign({}, f, { owner: st.owner })); });

    // timeline + totals
    const dates = [...new Set(proc.map(p => p.date))].sort((a, b) => a - b);
    const bdo = {}; proc.forEach(p => { const k = p.date + '|' + p.owner; bdo[k] = (bdo[k] || 0) + p.pts; });
    let cd = 0, ct = 0;
    const timeline = dates.map(dt => { cd += bdo[dt + '|Daniel'] || 0; ct += bdo[dt + '|Tanner'] || 0; return { date: fmtMs(dt), d: cd, t: ct }; });
    const daniel = cd, tanner = ct;
    const meta = { leader: daniel >= tanner ? 'Daniel' : 'Tanner', daniel, tanner, lead: Math.abs(daniel - tanner), races: new Set(field.map(f => f.race)).size, lastDate: dates.length ? fmtMs(Math.max.apply(null, dates)) : '' };

    // team pts per owner|match
    const teamPts = {}; proc.forEach(p => { const k = p.owner + '|' + p.match; teamPts[k] = (teamPts[k] || 0) + p.pts; });

    // rosters: bases by slot, subs nested beneath their base
    const buildRoster = owner => {
      const mine = riders.filter(r => r.owner === owner);
      const out = [];
      mine.filter(r => !r.replFor).sort((a, b) => a.slot - b.slot).forEach(b => {
        out.push({ slot: b.slot, rider: b.rider_name, pts: teamPts[owner + '|' + b.match] || 0, isSub: false, dropped: b.drop != null });
        mine.filter(r => r.replFor && r.effSlot === b.slot).sort((a, c) => a.add - c.add || a._i - c._i)
          .forEach(s => out.push({ slot: b.slot, rider: s.rider_name, pts: teamPts[owner + '|' + s.match] || 0, isSub: true, dropped: s.drop != null }));
      });
      return out;
    };
    const rostersByOwner = {}; owners.forEach(o => rostersByOwner[o] = buildRoster(o));

    // leaderboard: whole field by season points; team = points captured while owned
    const seasonByName = {}, order = [];
    field.forEach(f => { if (seasonByName[f.rider] == null) { seasonByName[f.rider] = 0; order.push(f.rider); } seasonByName[f.rider] += f.pts; });
    const teamByMatch = {}; proc.forEach(p => { teamByMatch[p.match] = (teamByMatch[p.match] || 0) + p.pts; });
    const lb = order.map(name => { const mt = norm(name); const owner = ownerByMatch[mt] || 'Free Agent'; return { rider: name, owner, season: seasonByName[name], team: owner === 'Free Agent' ? 0 : (teamByMatch[mt] || 0) }; })
      .sort((a, b) => b.season - a.season);
    const leaderboard = lb.map((r, i) => ({ rank: i + 1, rider: r.rider, owner: r.owner, season: r.season, team: r.team }));
    const freeAgents = lb.filter(r => r.owner === 'Free Agent').map((r, i) => ({ rank: i + 1, rider: r.rider, pts: r.season }));

    // history: every owned scoring result, newest first (date desc, place asc, pts desc)
    const blankStage = s => /one day/i.test(s) ? '' : s;
    const history = proc.slice().sort((a, b) => b.date - a.date || a.race.localeCompare(b.race) || a.place - b.place)
      .map(p => ({ date: fmtMs(p.date), race: p.race, stage: blankStage(p.stage), rider: p.rider, owner: p.owner, place: p.place, pts: p.pts }));

    // stats
    const stat = owner => { const ps = proc.filter(p => p.owner === owner); return { wins: ps.filter(p => p.place === 1).length, podiums: ps.filter(p => p.place <= 3).length, top10: ps.filter(p => p.place <= 10).length }; };
    const stats = { daniel: stat('Daniel'), tanner: stat('Tanner') };

    // tiers: points by draft-pick group of 5 (subs counted under their base slot)
    const grp = [[1, 5], [6, 10], [11, 15], [16, 20], [21, 25], [26, 30]];
    const tiers = grp.map(([a, b]) => {
      let d = 0, t = 0;
      riders.forEach(r => { if (r.effSlot >= a && r.effSlot <= b) { const v = teamPts[r.owner + '|' + r.match] || 0; if (r.owner === 'Daniel') d += v; else if (r.owner === 'Tanner') t += v; } });
      return { label: 'Picks ' + a + '\u2013' + b, t, d };
    });

    // top scorers per owner (top 10 by team points)
    const top = owner => {
      const nameOf = {}; riders.filter(r => r.owner === owner).forEach(r => { if (nameOf[r.match] == null) nameOf[r.match] = r.rider_name; });
      const agg = {}, ord = [];
      proc.filter(p => p.owner === owner).forEach(p => { if (agg[p.match] == null) { agg[p.match] = 0; ord.push(p.match); } agg[p.match] += p.pts; });
      return ord.map(m => ({ rider: nameOf[m] || m, pts: agg[m] })).sort((a, b) => b.pts - a.pts).slice(0, 10);
    };

    // schedule
    const scoredRaces = new Set(field.map(f => f.race));
    const schedule = schedRows.map(s => ({ date: (s.date || '').trim(), race: (s.race_name || '').trim(), tier: (s.tier || '').trim(), type: (s.race_type || '').trim(), scored: scoredRaces.has((s.race_name || '').trim()) }));

    return { meta, timeline, rostersByOwner, topDaniel: top('Daniel'), topTanner: top('Tanner'), leaderboard, freeAgents, history, stats, tiers, schedule };
  }

  /* ---------- loaders ---------- */
  async function fetchText(url) { const r = await fetch(url, { cache: 'no-store' }); if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + url); return r.text(); }
  async function fetchBuf(url) { const r = await fetch(url, { cache: 'no-store' }); if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + url); return r.arrayBuffer(); }

  const GT_CONFIG = {
    giro: { riders: 'giro-riders.csv', results: 'giro-results.xlsx', race: "Giro d'Italia", year: 2026, repcut: Date.UTC(2026, 4, 16) },
    tdf: { riders: 'tdf-riders.csv', results: 'tdf-results.xlsx', race: 'Tour de France', year: 2026, repcut: Date.UTC(2026, 6, 14), stageScoring: TDF_STAGE_SCORING },
    vuelta: { riders: 'vuelta-riders.csv', results: 'vuelta-results.xlsx', race: 'Vuelta a España', year: 2026, repcut: Date.UTC(2026, 8, 1) }
  };

  // Returns derived data object, or null if this race has no data files yet.
  async function loadGrandTour(race) {
    const cfg = GT_CONFIG[race]; if (!cfg) return null;
    let csv;
    try { csv = await fetchText(cfg.riders); } catch (e) { return null; }
    let rows = [];
    try { const buf = await fetchBuf(cfg.results); rows = await xlsxRows(buf); } catch (e) {}
    try {
      return computeGrandTour(csv, rows, cfg);
    } catch (e) { console.error('[FantasyEngine] grand tour compute failed', e); return null; }
  }

  const ANNUAL_CONFIG = { riders: 'annual-riders.csv', schedule: 'annual-schedule.csv', results: 'annual-results.xlsx' };
  async function loadAnnual() {
    let ridersCsv, schedCsv, buf;
    try { [ridersCsv, schedCsv, buf] = await Promise.all([fetchText(ANNUAL_CONFIG.riders), fetchText(ANNUAL_CONFIG.schedule), fetchBuf(ANNUAL_CONFIG.results)]); }
    catch (e) { return null; }
    try { return computeAnnual(ridersCsv, schedCsv, await xlsxRows(buf)); }
    catch (e) { console.error('[FantasyEngine] annual compute failed', e); return null; }
  }

  // Build a riders CSV (same shape as tdf-riders.csv) from live Firebase data:
  // draft picks (rooms/<room>/picks) + transfers (transfers/<room>/pending).
  // All draft picks share DRAFT_DATE so their team_pick slot follows draft order.
  function buildTdfRosterCsv(picks, transfers, draftDate) {
    draftDate = draftDate || '2026-07-03';
    const C = ['rider_name', 'owner', 'add_date', 'drop_date', 'is_replacement', 'rider_role', 'replaces_rider'];
    const rows = [];
    Object.keys(picks || {}).map(k => picks[k]).filter(Boolean).sort((a, b) => (a.ts || 0) - (b.ts || 0))
      .forEach(p => rows.push({ rider_name: p.rider, owner: p.owner, add_date: draftDate, drop_date: '', is_replacement: 'False', rider_role: p.role || '', replaces_rider: '' }));
    Object.keys(transfers || {}).map(k => transfers[k]).filter(Boolean).sort((a, b) => (a.ts || 0) - (b.ts || 0))
      .forEach(sw => {
        const orig = rows.find(r => norm(r.rider_name) === norm(sw.out) && r.owner === sw.owner && !(r.drop_date || '').trim());
        if (orig) orig.drop_date = sw.restDay;
        rows.push({ rider_name: sw.inName, owner: sw.owner, add_date: sw.restDay, drop_date: '', is_replacement: 'True', rider_role: sw.outRole || '', replaces_rider: sw.out });
      });
    const esc = x => /[",\n]/.test(String(x == null ? '' : x)) ? '"' + String(x).replace(/"/g, '""') + '"' : String(x == null ? '' : x);
    return C.join(',') + '\n' + rows.map(r => C.map(c => esc(r[c])).join(',')).join('\n');
  }

  // Rider archetypes (used by the dev-tool 2024/2025 replay). Names are
  // normalized with norm() at build time. Anyone not listed and not finishing
  // top-10 on GC is treated as a Stage Hunter.
  const SPRINTERS = new Set([
    'Jasper Philipsen', 'Jonathan Milan', 'Tim Merlier', 'Kaden Groves', 'Biniam Girmay',
    'Dylan Groenewegen', 'Jordi Meeus', 'Phil Bauhaus', 'Pascal Ackermann', 'Bryan Coquard',
    'Paul Penhoet', 'Pavel Bittner', 'Tobias Lund Andresen', 'Alberto Dainese', 'Arnaud De Lie',
    'Soren Waerenskjold', 'Jake Stewart', 'Stian Fredheim', 'Marius Mayrhofer', 'Mads Pedersen',
    'Mark Cavendish', 'Alexander Kristoff', 'Fabio Jakobsen', 'Sam Bennett', 'Caleb Ewan',
    'Danny Van Poppel', 'Jordi Warlop', 'Amaury Capiot', 'Jenno Berckmoes', 'Marijn Van Den Berg',
    'Cees Bol', 'Fernando Gaviria', 'Sam Welsford', 'Olav Kooij'
  ].map(norm));
  const GC_RIDERS = new Set([
    'Tadej Pogacar', 'Jonas Vingegaard', 'Remco Evenepoel', 'Primoz Roglic', 'Joao Almeida',
    'Adam Yates', 'Simon Yates', 'Enric Mas', 'Carlos Rodriguez', 'Richard Carapaz',
    'Mattias Skjelmose Jensen', 'Felix Gall', 'Ben O\'connor', 'Aleksandr Vlasov', 'Oscar Onley',
    'Florian Lipowitz', 'Kevin Vauquelin', 'Tobias Halland Johannessen', 'Lenny Martinez',
    'Santiago Buitrago Sanchez', 'Sepp Kuss', 'Guillaume Martin', 'Romain Bardet', 'Derek Gee',
    'Edward Irl Dunbar', 'Thymen Arensman', 'Egan Bernal', 'Geraint Thomas', 'Mikel Landa',
    'David Gaudu', 'Jai Hindley', 'Juan Ayuso', 'Aleksander Vlasov', 'Cian Uijtdebroeks',
    'Mattias Skjelmose', 'Pello Bilbao', 'Wilco Kelderman', 'Gaudu David'
  ].map(norm));

  // Per-rider fantasy leaderboard from results alone (no roster). Scores every
  // rider with the GT system: stage points accumulate; GC + jersey count only at
  // the latest stage (snapshot). Returns ranked rows with a points breakdown.
  function computeRiderLeaderboard(resultsRows, cfg) {
    cfg = cfg || {};
    const STAGE_SCORE = cfg.stageScoring || GT_SCORING['Stage Result'];
    const resObjs = rowsToObjs(resultsRows);
    const raw = [], stageMeta = [];
    for (const sd of resObjs) {
      const s = +sd.Stage; if (!s || !sd['1st'] || String(sd['1st']).trim() === '') continue;
      stageMeta.push({ s, date: sd.Date, winner: sd['1st'] });
      STAGE_COLS.forEach((c, i) => { if (sd[c] && String(sd[c]).trim() !== '') raw.push({ Stage: s, rider: sd[c], rank: i + 1, cat: 'Stage Result' }); });
      for (let i = 1; i <= 10; i++) { const c = 'GC #' + i; if (sd[c] && String(sd[c]).trim() !== '') raw.push({ Stage: s, rider: sd[c], rank: i, cat: 'GC Standing' }); }
      [['Points #', 'Points Jersey'], ['Mountain #', 'Mountain Jersey'], ['Youth #', 'Youth Jersey']].forEach(([pre, cat]) => {
        for (let i = 1; i <= 3; i++) { const c = pre + i; if (sd[c] && String(sd[c]).trim() !== '') raw.push({ Stage: s, rider: sd[c], rank: i, cat }); }
      });
    }
    const allStages = [...new Set(stageMeta.map(x => x.s))].sort((a, b) => a - b);
    if (!allStages.length) return null;
    const latest = Math.max(...allStages);
    const by = {}; // displayName key -> agg
    for (const e of raw) {
      if (e.cat !== 'Stage Result' && e.Stage !== latest) continue; // snapshot for GC/jersey
      const tbl = e.cat === 'Stage Result' ? STAGE_SCORE : (GT_SCORING[e.cat] || GT_SCORING.Jersey); const pts = tbl[e.rank] || 0;
      if (!pts) continue;
      const k = norm(e.rider);
      const o = by[k] || (by[k] = { name: e.rider, sr: 0, gc: 0, jer: 0, sw: 0, podiums: 0, gcRank: 99, ptsJer: 0, komJer: 0, ythJer: 0 });
      o.name = e.rider;
      if (e.cat === 'Stage Result') { o.sr += pts; if (e.rank === 1) o.sw += 1; if (e.rank <= 3) o.podiums += 1; }
      else if (e.cat === 'GC Standing') { o.gc += pts; o.gcRank = Math.min(o.gcRank, e.rank); }
      else { o.jer += pts; if (e.cat === 'Points Jersey') o.ptsJer = Math.max(o.ptsJer, 4 - e.rank); else if (e.cat === 'Mountain Jersey') o.komJer = Math.max(o.komJer, 4 - e.rank); else o.ythJer = Math.max(o.ythJer, 4 - e.rank); }
    }
    // Role by rider identity, not by jersey results.
    //  1) top-10 final GC  -> GC
    //  2) known sprinter   -> Sprinter
    //  3) known GC rider    -> GC
    //  4) everyone else     -> Stage Hunter
    const roleOf = o => {
      if (o.gcRank <= 10) return 'GC';
      const n = norm(o.name);
      if (SPRINTERS.has(n)) return 'Sprinter';
      if (GC_RIDERS.has(n)) return 'GC';
      return 'Stage Hunter';
    };
    const lb = Object.keys(by).map(k => { const o = by[k]; return { name: o.name, role: roleOf(o), sr: r1(o.sr), gc: r1(o.gc), jer: r1(o.jer), total: r1(o.sr + o.gc + o.jer), wins: o.sw, podiums: o.podiums }; })
      .filter(x => x.total > 0).sort((a, b) => b.total - a.total).map((x, i) => Object.assign({ rank: i + 1 }, x));
    const lastMeta = stageMeta.find(x => x.s === latest);
    return { race: cfg.race || '', year: cfg.year || '', stages: latest, lastDate: lastMeta ? fmtDate(lastMeta.date) : '', leaderboard: lb };
  }

  window.FantasyEngine = Object.assign(window.FantasyEngine || {}, {
    loadGrandTour, computeGrandTour, loadAnnual, computeAnnual, buildTdfRosterCsv, computeRiderLeaderboard, xlsxRows, parseCsv, rowsToObjs, _helpers: { norm, parseDate, fmtDate, fmtMs, r1 }
  });
})();
