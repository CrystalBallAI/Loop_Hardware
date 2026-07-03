/* hw-integration.js — keep the reference hardware UI (central product image, score
   hero, building-block ribbon) and feed it REAL pipeline data by overriding the
   page's own data globals and re-running its native builders. Adds a compact
   fixed bar with the live grade/decision + Excel/Provenance/Back-to-Load. */
(function () {
  'use strict';

  var KEY_BY_MODULE = { drone: 'drone', base: 'basestation', gcp: 'gcp', checkpoint: 'checkpoint' };
  var LABEL = { drone: 'Drone', base: 'Base Station', gcp: 'Control Point', checkpoint: 'Check Point' };
  // the page's per-subsystem confidence engine for each hardware module
  var DS = { drone: 'dsDrone', base: 'dsBase', gcp: 'dsGcp', checkpoint: 'dsCp' };

  var cache = {};

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function scoreColor(s) { return s >= 90 ? '#64be94' : s >= 75 ? '#5596cc' : s >= 60 ? '#d2aa4e' : '#c86262'; }

  function fetchResult(key, cb) {
    if (cache[key] !== undefined) { cb(cache[key]); return; }
    fetch('/api/results/' + key)
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (d) { cache[key] = d; cb(d); })
      .catch(function () { cache[key] = null; cb(null); });
  }

  var SCORE_PREFIX = { drone: 'dn-', base: '', gcp: 'gp-', checkpoint: 'cp-' };

  function blockLevel(bb) {
    var lvl = 'good';
    (bb.indicators || []).forEach(function (i) {
      var l = (i.level || '').toLowerCase();
      if (l === 'resurvey' || l === 'critical') lvl = 'resurvey';
      else if ((l === 'review' || l === 'minor') && lvl !== 'resurvey') lvl = 'review';
    });
    if (bb.anomaly && lvl === 'good') lvl = 'review';
    return lvl;
  }
  function statusText(l) { return l === 'good' ? 'Nominal' : (l === 'resurvey' ? 'Action req.' : 'Review'); }

  // Patch the native hero with authoritative pipeline values (apex + real reason),
  // and remove the mock "+2.3%" delta (keep the real hard-gate message).
  function patchHero(module, res, overall) {
    var p = SCORE_PREFIX[module];
    var d = document.getElementById(p + 'scoreDelta');
    if (d) d.textContent = res.globalGate ? 'Hard gate — score forced to 0' : '';
    if (overall) {
      var s = document.getElementById(p + 'scoreNum');
      if (s && res.overallScore != null) s.innerHTML = res.overallScore + '<span class="pct">%</span>';
      var r = document.getElementById(p + 'mdReason');
      if (r && res.recommendation) r.textContent = res.recommendation;
    }
  }

  // DRONE only: the page rubric splits calibration into a 4th block, but the pipeline
  // (06) folds it into Image Capture. Render the block cards straight from 06's 3 blocks.
  var droneSelected = {};
  function renderDroneBlocks(res) {
    var host = document.getElementById('dn-bbStripHead');
    if (!host || !res.bbs || !res.bbs.length) return;
    droneSelected = {};
    host.innerHTML = res.bbs.map(function (bb) {
      var lvl = blockLevel(bb);
      var inds = bb.indicators || [];
      var f = inds.filter(function (i) { return (i.level || 'good').toLowerCase() !== 'good'; }).length;
      var imp = bb.weight >= 0.25 ? 'High' : (bb.weight >= 0.15 ? 'Medium' : 'Low');
      var diag = lvl === 'good' ? ('All ' + inds.length + ' indicators nominal') : (f + ' of ' + inds.length + ' indicators flagged');
      // identical markup to the native card: status reads OK / Review (not a number)
      return '<div class="bb-cell ' + lvl + '" id="dn-' + esc(bb.id) + '" data-bb="' + esc(bb.id) + '">'
        + '<span class="bb-accent"></span>'
        + '<div class="bb-cell-top"><span class="bb-stat"><i class="bb-led"></i>' + statusText(lvl) + '</span>'
        + (lvl !== 'good' ? '<span class="bb-impact">' + imp + ' impact</span>' : '') + '</div>'
        + '<div class="bb-cell-name">' + esc(bb.name) + '<b class="bb-cell-status">' + (lvl === 'good' ? 'OK' : 'Review') + '</b></div>'
        + '<div class="bb-cell-diag">' + diag + '</div>'
        + '<div class="bb-cell-act"><span class="bb-inspect">Details &rsaquo;</span></div>'
        + '</div>';
    }).join('');
    renderDroneOverlay(res);
    host.querySelectorAll('.bb-cell').forEach(function (c) {
      var id = c.getAttribute('data-bb');
      c.onclick = function () { toggleDroneBB(res, id); };           // card -> indicator overlay (like native)
      var insp = c.querySelector('.bb-inspect');
      if (insp) insp.onclick = function (e) { e.stopPropagation(); openDroneNativePanel(res, id); }; // Details -> right panel
    });
  }

  // The indicator overlay (bbq tables on the product image) — same markup/classes as
  // the native renderIndicators, driven by our 06 block data.
  function renderDroneOverlay(res) {
    var layer = document.getElementById('dn-indicatorLayer');
    if (!layer) return;
    var QPOS = ['bbq-tl', 'bbq-tr', 'bbq-bl', 'bbq-br'];
    var html = [];
    (res.bbs || []).forEach(function (bb, idx) {
      if (!droneSelected[bb.id]) return;
      var rows = (bb.indicators || []).map(function (i) {
        var lv = (i.level || '').toLowerCase();
        var sev = lv === 'good' ? '' : (lv === 'resurvey' || lv === 'critical' ? ' sev-resurvey' : (lv === 'minor' ? ' sev-minor' : ' sev-review'));
        return '<div class="bbq-ind' + sev + '"><span class="bbq-dot"></span><em>' + esc(i.name) + '</em><b>'
          + (i.currentScore == null ? '—' : Math.round(i.currentScore)) + '</b></div>';
      }).join('');
      html.push('<div class="bbq ' + QPOS[idx % 4] + '"><div class="bbq-head"><span class="bbq-name">' + esc(bb.name)
        + '</span><span class="bbq-score">' + bb.score + '</span></div><div class="bbq-list">' + rows + '</div></div>');
    });
    layer.innerHTML = html.join('');
    layer.className = 'indicator-layer grouped' + (html.length ? ' show' : '');
    (res.bbs || []).forEach(function (bb) {
      var el = document.getElementById('dn-' + bb.id);
      if (el) el.classList.toggle('active', !!droneSelected[bb.id]);
    });
  }
  function toggleDroneBB(res, id) { droneSelected[id] = !droneSelected[id]; renderDroneOverlay(res); }

  // Open the ORIGINAL right-side drawer (#dn-drawer / #dn-drawerBody), populated with
  // our 06 block data using the page's own drawer markup + classes — identical look to
  // the Base/GCP/Check Point panels (reuses dsDrone.setSection / toggleVerified).
  function openDroneNativePanel(res, bbId) {
    var bb = (res.bbs || []).filter(function (b) { return b.id === bbId; })[0];
    var body = document.getElementById('dn-drawerBody');
    if (!bb || !body) return;
    // match native openBBDetails: select this block exclusively + show its overlay
    droneSelected = {}; droneSelected[bbId] = true; renderDroneOverlay(res);
    var lvl = blockLevel(bb);
    function lvlOf(i) { var l = (i.level || '').toLowerCase(); return l === 'critical' ? 'resurvey' : l; }
    var actionable = [], noted = [], verified = [];
    (bb.indicators || []).forEach(function (i) {
      var l = lvlOf(i);
      if (l === 'resurvey' || l === 'review') actionable.push(i);
      else if (l === 'minor') noted.push(i);
      else verified.push(i);
    });
    function acc(i) {
      var l = lvlOf(i), scCls = l === 'resurvey' ? 'resurvey' : (l === 'review' ? 'review' : '');
      var inner = (l === 'good' || !l)
        ? '<div class="acc-state">' + esc(i.verifiedStatement || i.desc || 'Verified and in good standing.') + '</div>'
          + (i.bandLabel ? '<div class="acc-evi">Evidence &middot; ' + esc(i.bandLabel) + '</div>' : '')
        : '<div class="acc-state">' + esc(i.bandLabel || '') + '</div>'
          + (i.impact ? '<div class="d-ind-impact">' + esc(i.impact) + '</div>' : '')
          + ((i.actions && i.actions.length)
            ? '<ul class="d-acts">' + i.actions.map(function (a) { return '<li>' + esc(a) + '</li>'; }).join('') + '</ul>'
            : '')
          + (i.alert ? '<div class="d-ind-impact">' + esc(i.alert) + '</div>' : '');
      return '<div class="acc"><div class="acc-head" onclick="this.parentNode.classList.toggle(\'open\')">'
        + '<span class="acc-chev">&#9654;</span><span class="acc-name">' + esc(i.name) + '</span>'
        + '<span class="acc-right"><span class="acc-sc ' + scCls + '">' + (i.currentScore == null ? '--' : Math.round(i.currentScore)) + '</span></span></div>'
        + '<div class="acc-body"><div class="acc-inner">' + inner + '</div></div></div>';
    }
    var actHtml = actionable.length ? actionable.map(acc).join('') : '<div class="d-empty">Nothing to action in this block.</div>';
    var notedHtml = noted.map(acc).join('');
    var verHtml = verified.length ? verified.map(acc).join('') : '<div class="d-empty">No checks passed cleanly.</div>';
    var verBlock = verified.length
      ? '<div class="d-sec-row"><div style="display:flex;align-items:baseline;gap:10px;flex:1;min-width:0">'
        + '<span class="d-sec verified" style="margin:0;padding:0;border:0;flex-shrink:0">Verified</span>'
        + '<span class="d-empty" style="padding:0">' + verified.length + ' indicators verified and in good standing.</span></div>'
        + '<button class="d-ctrl" id="dn-verToggle" onclick="dsDrone.toggleVerified()" style="flex-shrink:0">+ More Details</button></div>'
        + '<div id="dn-verSec" style="display:none">' + verHtml + '</div>'
      : '<div class="d-sec-row"><div class="d-sec verified">Verified<span class="d-sec-count">0</span></div></div><div id="dn-verSec">' + verHtml + '</div>';
    body.innerHTML =
      '<h2>' + esc(bb.name) + '</h2>'
      + '<div class="d-headrow"><div class="d-headmetric"><div class="d-score">' + bb.score + '<span>%</span></div>'
      + '<div class="d-verdict ' + lvl + '">' + statusText(lvl) + '</div></div>'
      + '<span class="d-info-wrap"><button class="d-info" type="button" aria-label="About this block">i</button>'
      + '<span class="d-info-pop">' + esc(bb.desc || bb.name) + '</span></span></div>'
      + '<div class="d-sec-row"><div class="d-sec actionable">Actionables<span class="d-sec-count">' + actionable.length + '</span></div>'
      + '<div class="d-ctrls"><button class="d-ctrl" onclick="dsDrone.setSection(\'#dn-actSec\',true)">Expand all</button>'
      + '<button class="d-ctrl" onclick="dsDrone.setSection(\'#dn-actSec\',false)">Collapse all</button></div></div>'
      + '<div id="dn-actSec">' + actHtml + notedHtml + '</div>'
      + verBlock;
    var d = document.getElementById('dn-drawer');
    if (d) d.classList.add('open');
  }

  // Feed REAL per-indicator scores into the page's own confidence engine and let it
  // re-render the native hero / building blocks / indicators (keeps the reference UI).
  function injectNative(module, res) {
    if (!res) return;
    var ds = window[DS[module]];
    if (!ds || typeof ds.setLive !== 'function') return;
    try {
      ds.setLive(res.scores || {}, res.points || [], res.nulls || []);
      patchHero(module, res, true);
      if (module === 'drone') renderDroneBlocks(res);
    } catch (e) { /* leave native mock if the engine rejects the data */ }
  }

  // ---- compact fixed live-result bar --------------------------------------
  var bar;
  function ensureBar() {
    if (bar) return bar;
    bar = document.createElement('div');
    bar.className = 'cbmi-bar hidden';
    document.body.appendChild(bar);
    return bar;
  }
  function updateBar(module, res) {
    var key = KEY_BY_MODULE[module];
    var b = ensureBar();
    if (!res) {
      b.className = 'cbmi-bar';
      b.innerHTML = '<span class="cbmi-bar-live">No live run</span>'
        + '<span class="cbmi-bar-note">Run ' + esc(LABEL[module]) + ' on the Load page</span>'
        + '<span class="cbmi-bar-actions"><a class="cbmi-bar-btn" href="/">&#8592; Load</a></span>';
      return;
    }
    var dec = (res.verdict || 'review');
    b.className = 'cbmi-bar';
    b.innerHTML = '<span class="cbmi-bar-live">Live result</span>'
      + '<span class="cbmi-bar-score" style="color:' + scoreColor(res.overallScore || 0) + '">' + (res.overallScore == null ? '--' : res.overallScore) + '</span>'
      + '<span class="cbmi-bar-meta"><span class="cbmi-bar-grade">' + esc(res.tier || res.grade || '') + '</span>'
      + '<span class="cbmi-bar-dec ' + dec + '">' + esc(res.decisionLabel || res.decision || '') + (res.globalGate ? ' · GATE' : '') + '</span></span>'
      + '<span class="cbmi-bar-actions">'
      + '<a class="cbmi-bar-btn primary" href="/api/download/' + key + '/xlsx" target="_blank">Excel</a>'
      + '<a class="cbmi-bar-btn" href="/api/download/' + key + '/provenance" target="_blank">Provenance</a>'
      + '<a class="cbmi-bar-btn" href="/">&#8592; Load</a>'
      + '</span>';
  }

  function renderForModule(module) {
    if (!KEY_BY_MODULE[module]) return;
    fetchResult(KEY_BY_MODULE[module], function (res) {
      injectNative(module, res);
    });
  }

  function rewireLoadBadge() {
    var b = document.getElementById('nbadge-load');
    if (b) { b.onclick = function () { location.href = '/'; }; b.classList.remove('locked'); }
  }

  function boot() {
    try { window.originComplete = function () { return true; }; } catch (e) {}

    // FIX (pre-existing SPA bug): the gcp/checkpoint engines reference
    // `computeBlockScore`, which the bundle only defines inside the base/drone
    // engines — so for gcp/cp it's an undefined free variable that throws inside
    // renderIndicators and aborts the right panel. Provide a global fallback
    // (base/drone keep their own scoped version, which shadows this).
    if (typeof window.computeBlockScore !== 'function') {
      window.computeBlockScore = function (blockId, scores) {
        scores = scores || {};
        for (var k in cache) {
          var res = cache[k];
          if (!res || !res.bbs) continue;
          for (var i = 0; i < res.bbs.length; i++) {
            var bb = res.bbs[i];
            if (bb.bbId === blockId) {
              var vals = (bb.indicators || []).map(function (x) { return scores[x.id]; })
                .filter(function (v) { return typeof v === 'number'; });
              if (vals.length) return vals.reduce(function (a, b) { return a + b; }, 0) / vals.length;
              return bb.score;
            }
          }
        }
        var all = Object.keys(scores).map(function (id) { return scores[id]; })
          .filter(function (v) { return typeof v === 'number'; });
        return all.length ? all.reduce(function (a, b) { return a + b; }, 0) / all.length : 0;
      };
    }

    rewireLoadBadge();

    // pre-fetch all results so injection after switchModule is instant (no mock flash lingering)
    Object.keys(KEY_BY_MODULE).forEach(function (m) { fetchResult(KEY_BY_MODULE[m], function () {}); });

    if (typeof window.switchModule === 'function' && !window.__cbmiWrapped) {
      var orig = window.switchModule;
      window.switchModule = function (m) {
        var r = orig.apply(this, arguments);
        if (KEY_BY_MODULE[m]) { rewireLoadBadge(); renderForModule(m); }
        return r;
      };
      window.__cbmiWrapped = true;
    }

    // gcp/checkpoint re-render on point change — re-apply our hero patches afterwards
    ['gcp', 'checkpoint'].forEach(function (module) {
      var ds = window[DS[module]];
      if (ds && typeof ds.selectPoint === 'function' && !ds.__cbmiPt) {
        var origSel = ds.selectPoint;
        ds.__cbmiPt = true;
        ds.selectPoint = function (val) {
          var r = origSel.apply(this, arguments);
          var res = cache[KEY_BY_MODULE[module]];
          if (res) patchHero(module, res, (!val || val === 'overall'));
          return r;
        };
      }
    });

    var h = (location.hash || '').replace('#', '') || 'drone';
    if (!KEY_BY_MODULE[h]) h = 'drone';
    if (typeof window.switchModule === 'function') window.switchModule(h);
    else renderForModule(h);

    window.addEventListener('hashchange', function () {
      var m = (location.hash || '').replace('#', '');
      if (KEY_BY_MODULE[m] && typeof window.switchModule === 'function') window.switchModule(m);
    });
  }

  if (document.readyState === 'complete') setTimeout(boot, 250);
  else window.addEventListener('load', function () { setTimeout(boot, 250); });
})();
