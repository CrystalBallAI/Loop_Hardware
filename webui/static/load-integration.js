/* load-integration.js — wire the Load page to the backend.
   - captures the real File objects the user picks (the page only keeps metadata)
   - overrides ldSubmit() to POST them to /api/run and show a live progress panel
   - polls /api/run/{id} and links each finished subsystem to its hardware page
   Targets only the 4 hardware subsystems; other Load sections are ignored. */
(function () {
  'use strict';

  var SYS_KEY = { 'Drone': 'drone', 'Base Station': 'basestation', 'Control Point': 'gcp', 'Check Point': 'checkpoint' };
  var KEY_MODULE = { drone: 'drone', basestation: 'base', gcp: 'gcp', checkpoint: 'checkpoint' };
  var KEY_LABEL = { drone: 'Drone', basestation: 'Base Station', gcp: 'Control Point', checkpoint: 'Check Point' };

  var CAP = {};        // itemId -> [File]     (drone_*/base_* single-slot items)
  var CAP_POINT = {};  // sys -> idx -> fileId -> [File]   (Control/Check point per-point)

  function keyForItem(id) {
    if (id.indexOf('drone_') === 0) return 'drone';
    if (id.indexOf('base_') === 0) return 'basestation';
    return null;
  }

  // ---- capture real files (page's handler only stores name/size) ----------
  function installCapture() {
    var fi = document.getElementById('ld-fileinput');
    if (!fi) return;
    fi.addEventListener('change', function (e) {
      var files = e.target.files;
      if (!files || !files.length) return;
      var arr = Array.prototype.slice.call(files);
      var pt = window.LOAD_PENDING_POINT, id = window.LOAD_PENDING_ID;
      if (pt) {
        CAP_POINT[pt.sys] = CAP_POINT[pt.sys] || {};
        CAP_POINT[pt.sys][pt.idx] = CAP_POINT[pt.sys][pt.idx] || {};
        CAP_POINT[pt.sys][pt.idx][pt.fileId] = arr;
      } else if (id) {
        CAP[id] = arr;
      } else if (typeof window.ldMatchInput === 'function') {
        arr.forEach(function (f) {
          var mid = window.ldMatchInput(f);
          if (mid) { CAP[mid] = (CAP[mid] || []); CAP[mid].push(f); }
        });
      }
    }, true); // capture phase: runs before the page's own change handler

    // drag-and-drop path goes through ldIngestFiles, not the input
    if (typeof window.ldIngestFiles === 'function') {
      var orig = window.ldIngestFiles;
      window.ldIngestFiles = function (files) {
        if (typeof window.ldMatchInput === 'function') {
          Array.prototype.slice.call(files).forEach(function (f) {
            var mid = window.ldMatchInput(f);
            if (mid) { CAP[mid] = (CAP[mid] || []); CAP[mid].push(f); }
          });
        }
        return orig.apply(this, arguments);
      };
    }
  }

  // ---- build the multipart payload ---------------------------------------
  function buildRun() {
    var fd = new FormData(), subs = {};
    Object.keys(CAP).forEach(function (itemId) {
      var key = keyForItem(itemId);
      if (!key) return;
      CAP[itemId].forEach(function (f) { fd.append(key + '__' + itemId, f, f.name); });
      subs[key] = 1;
    });
    Object.keys(CAP_POINT).forEach(function (sys) {
      var key = SYS_KEY[sys];
      if (!key) return;
      var rows = CAP_POINT[sys];
      Object.keys(rows).forEach(function (idx) {
        var slots = rows[idx];
        Object.keys(slots).forEach(function (fileId) {
          slots[fileId].forEach(function (f) { fd.append(key + '__point' + idx + '__' + fileId, f, f.name); });
        });
      });
      subs[key] = 1;
    });
    try { fd.append('forms', JSON.stringify(window.FORM_STATE || {})); } catch (e) {}
    return { fd: fd, subs: Object.keys(subs) };
  }

  // ---- progress overlay ---------------------------------------------------
  var ov, lastRunId;
  function ensureOverlay() {
    if (ov) return ov;
    ov = document.createElement('div');
    ov.className = 'cbmi-run-ov hidden';
    ov.innerHTML = '<div class="cbmi-run">'
      + '<div class="cbmi-run-head"><span class="cbmi-run-title">Validate &amp; Run Scoring</span>'
      + '<button class="cbmi-run-close" title="Close">&times;</button></div>'
      + '<div class="cbmi-run-sub">Running each subsystem pipeline — process · score · recommend · excel</div>'
      + '<div class="cbmi-run-body"></div></div>';
    document.body.appendChild(ov);
    ov.querySelector('.cbmi-run-close').onclick = function () { ov.classList.add('hidden'); };
    return ov;
  }
  function openOverlay() { ensureOverlay().classList.remove('hidden'); }

  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function scoreColor(s) { return s >= 90 ? '#64be94' : s >= 75 ? '#5596cc' : s >= 60 ? '#d2aa4e' : '#c86262'; }

  function renderRun(run) {
    var body = ensureOverlay().querySelector('.cbmi-run-body');
    var html = '';
    Object.keys(run.subsystems).forEach(function (key) {
      var s = run.subsystems[key];
      var steps = s.steps.map(function (st) {
        var mark = st.status === 'done' ? '&#10003;' : (st.status === 'failed' ? '&times;' : '');
        return '<div class="cbmi-step ' + st.status + '"><span class="cbmi-step-dot">' + mark + '</span>'
          + '<span class="cbmi-step-lbl">' + esc(st.label) + '</span></div>';
      }).join('');
      var badge = s.status === 'done' ? 'done' : (s.status === 'failed' ? 'failed' : (s.status === 'running' ? 'running' : ''));
      var result = '';
      if (s.status === 'done') {
        var mod = KEY_MODULE[key];
        result = '<div class="cbmi-sub-result show">'
          + '<span class="cbmi-res-score" style="color:' + scoreColor(s.score || 0) + '">' + (s.score == null ? '--' : s.score) + '</span>'
          + '<span class="cbmi-res-meta">' + esc(s.grade || '') + (s.decision ? ' · ' + esc(s.decision) : '') + '</span>'
          + '<span class="cbmi-res-actions">'
          + '<a class="cbmi-res-btn primary" href="/hw#' + mod + '" target="_blank">View on hardware page &#8594;</a>'
          + '<a class="cbmi-res-btn" href="/api/download/' + key + '/xlsx" target="_blank">Excel</a>'
          + '</span></div>';
      } else if (s.status === 'failed') {
        result = '<div class="cbmi-sub-result show"><span class="cbmi-res-err">&#9888; ' + esc(s.error || 'failed') + '</span></div>';
      }
      html += '<div class="cbmi-sub-row">'
        + '<div class="cbmi-sub-top"><span class="cbmi-sub-name">' + esc(s.label) + '</span>'
        + '<span class="cbmi-sub-badge ' + badge + '">' + esc(s.status) + '</span></div>'
        + '<div class="cbmi-steps">' + steps + '</div>'
        + result + '</div>';
    });
    body.innerHTML = html;
  }

  function poll(runId) {
    fetch('/api/run/' + runId).then(function (r) { return r.json(); }).then(function (run) {
      if (run && run.subsystems) renderRun(run);
      if (run && run.status !== 'done') setTimeout(function () { poll(runId); }, 1200);
      else if (typeof window.cbmiOnRunDone === 'function') window.cbmiOnRunDone();
    }).catch(function () { setTimeout(function () { poll(runId); }, 2000); });
  }

  function startRun(fd) {
    openOverlay();
    ensureOverlay().querySelector('.cbmi-run-body').innerHTML =
      '<div class="cbmi-run-sub">Uploading &amp; starting…</div>';
    fetch('/api/run', { method: 'POST', body: fd }).then(function (r) { return r.json(); })
      .then(function (d) {
        if (d.error) {
          ensureOverlay().querySelector('.cbmi-run-body').innerHTML =
            '<div class="cbmi-res-err" style="padding:14px">&#9888; ' + esc(d.error) + '</div>';
          return;
        }
        lastRunId = d.runId;
        poll(d.runId);
      })
      .catch(function (e) {
        ensureOverlay().querySelector('.cbmi-run-body').innerHTML =
          '<div class="cbmi-res-err" style="padding:14px">&#9888; ' + esc(e) + '</div>';
      });
  }

  // ---- overrides + sample button -----------------------------------------
  function overrideSubmit() {
    window.ldSubmit = function () {
      var msg = document.getElementById('ld-submit-msg');
      var built = buildRun();
      if (!built.subs.length) {
        if (msg) {
          msg.className = 'ld-submit-msg err';
          msg.textContent = 'Upload files for at least one of: Drone, Base Station, Control Point, Check Point.';
        }
        return;
      }
      built.fd.append('subsystems', JSON.stringify(built.subs));
      if (msg) { msg.className = 'ld-submit-msg ok'; msg.textContent = 'Submitting ' + built.subs.length + ' subsystem(s)…'; }
      startRun(built.fd);
    };
  }

  // ---- app badge (version + beta expiry) ---------------------------------
  var MODULE_BY_KEY = { drone: 'drone', basestation: 'base', gcp: 'gcp', checkpoint: 'checkpoint' };
  var LABEL_BY_KEY = { drone: 'Drone', basestation: 'Base', gcp: 'Control Pt', checkpoint: 'Check Pt' };

  function injectAppBadge() {
    fetch('/api/app').then(function (r) { return r.json(); }).then(function (a) {
      var el = document.getElementById('cbmi-appbadge') || document.createElement('div');
      el.id = 'cbmi-appbadge';
      el.className = 'cbmi-appbadge' + (a.ok ? '' : ' expired');
      var right = a.note ? '<span class="cbmi-appbadge-note">' + esc(a.note) + '</span>' : '';
      el.innerHTML = '<span class="cbmi-appbadge-v">CBMI Loop ' + esc(a.version || '') + '</span>'
        + (a.tester && a.tester !== 'dev' ? '<span class="cbmi-appbadge-t">' + esc(a.tester) + '</span>' : '')
        + right;
      if (!el.parentNode) document.body.appendChild(el);
    }).catch(function () {});
  }

  // ---- previous runs panel ------------------------------------------------
  function fmtTime(epoch) {
    if (!epoch) return '';
    try { return new Date(epoch * 1000).toLocaleString(); } catch (e) { return ''; }
  }

  function renderHistory(runs) {
    var panel = document.getElementById('cbmi-history');
    if (!panel) return;
    if (!runs || !runs.length) {
      panel.innerHTML = '<div class="cbmi-hist-head">Previous runs</div>'
        + '<div class="cbmi-hist-empty">No runs yet — upload inputs and click Validate &amp; Run Scoring.</div>';
      return;
    }
    var rows = runs.map(function (run) {
      var subs = Object.keys(run.subsystems || {}).map(function (k) {
        var s = run.subsystems[k];
        var done = s.status === 'done';
        var sc = (s.score == null ? '—' : s.score);
        return '<span class="cbmi-hist-sub ' + (done ? 'ok' : 'bad') + '">'
          + '<b>' + esc(LABEL_BY_KEY[k] || k) + '</b> ' + sc
          + (s.decision ? ' · ' + esc(s.decision) : '') + '</span>';
      }).join('');
      return '<div class="cbmi-hist-row">'
        + '<div class="cbmi-hist-when">' + esc(fmtTime(run.finishedAt) || run.runId) + '</div>'
        + '<div class="cbmi-hist-subs">' + subs + '</div>'
        + '<button class="cbmi-hist-open" data-run="' + esc(run.runId) + '">Open results ›</button>'
        + '</div>';
    }).join('');
    panel.innerHTML = '<div class="cbmi-hist-head">Previous runs'
      + '<button class="cbmi-hist-diag" title="Export a diagnostics bundle to send to support">Export diagnostics</button>'
      + '</div>' + rows;
    Array.prototype.forEach.call(panel.querySelectorAll('.cbmi-hist-open'), function (b) {
      b.onclick = function () {
        var rid = b.getAttribute('data-run');
        fetch('/api/runs/' + rid + '/activate', { method: 'POST' })
          .then(function (r) { return r.json(); })
          .then(function (d) {
            var first = (d.activated && d.activated[0]) || 'drone';
            window.open('/hw#' + (MODULE_BY_KEY[first] || first), '_blank');
          }).catch(function () {});
      };
    });
    var diag = panel.querySelector('.cbmi-hist-diag');
    if (diag) diag.onclick = function () { window.open('/api/diagnostics', '_blank'); };
  }

  function refreshHistory() {
    fetch('/api/runs').then(function (r) { return r.json(); })
      .then(function (d) { renderHistory(d.runs || []); }).catch(function () {});
  }

  function injectHistoryPanel() {
    if (document.getElementById('cbmi-history')) return;
    var panel = document.createElement('div');
    panel.id = 'cbmi-history';
    panel.className = 'cbmi-history';
    var bar = document.querySelector('.ld-submitbar');
    if (bar && bar.parentNode) bar.parentNode.insertBefore(panel, bar.nextSibling);
    else document.body.appendChild(panel);
    refreshHistory();
  }

  // refresh history when a run finishes (poll() calls this via the global)
  window.cbmiOnRunDone = refreshHistory;

  function boot() {
    // Load page shows only the 4 hardware subsystems we actually run
    // (Drone / Base Station / Control Point / Check Point).
    var KEEP = { 'Drone': 1, 'Base Station': 1, 'Control Point': 1, 'Check Point': 1 };
    if (Array.isArray(window.LOAD_SYSTEMS)) {
      window.LOAD_SYSTEMS = window.LOAD_SYSTEMS.filter(function (g) { return KEEP[g.sys]; });
      if (typeof window.renderLoad === 'function') window.renderLoad();
    }
    installCapture();
    overrideSubmit();
    injectAppBadge();
    injectHistoryPanel();
  }

  if (document.readyState === 'complete' || document.readyState === 'interactive') setTimeout(boot, 150);
  else window.addEventListener('load', function () { setTimeout(boot, 150); });
})();
