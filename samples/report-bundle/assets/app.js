/*
 * Minimal stand-in for the report bundle runtime described in the TestimoX
 * reporting platform plan (section 6.2). It is intentionally ES5 and
 * dependency-free so it runs under file://, inside a UHV srcdoc iframe and
 * inside a UHV blob: iframe without a build step.
 *
 * - HfxData.register(id, chunk, payload) is called by data/<dataset>/<chunk>.js
 *   sidecar scripts (script tags, because fetch is blocked under file://).
 * - HfxData.load(id, chunk, callback) injects a sidecar script at runtime.
 * - The page renders a link at runtime so hosts can prove they intercept
 *   anchors that did not exist when the HTML was parsed.
 * - A proposed theme handshake listener applies host theme tokens sent by UHV
 *   via postMessage (not implemented in UHV yet; see
 *   docs/Report-Bundle-Compatibility-Spike.md).
 */
(function (global) {
  'use strict';

  var registry = {};
  var waiters = {};

  function key(id, chunk) {
    return String(id) + '/' + String(chunk);
  }

  var HfxData = {
    register: function (id, chunk, payload) {
      var entryKey = key(id, chunk);
      registry[entryKey] = payload;
      var pending = waiters[entryKey] || [];
      delete waiters[entryKey];
      for (var index = 0; index < pending.length; index += 1) {
        pending[index](null, payload);
      }
    },
    get: function (id, chunk) {
      return registry[key(id, chunk)];
    },
    resolveChunkUrl: function (id, chunk) {
      // Resolved against document.baseURI, which is UHV's injected <base href>
      // when hosted in SharePoint and the file location under file://.
      return new URL('data/' + id + '/' + chunk + '.js', document.baseURI).toString();
    },
    load: function (id, chunk, callback) {
      var entryKey = key(id, chunk);
      if (registry[entryKey]) {
        callback(null, registry[entryKey]);
        return;
      }
      (waiters[entryKey] = waiters[entryKey] || []).push(callback);
      if (waiters[entryKey].length > 1) {
        return;
      }
      var script = document.createElement('script');
      script.src = HfxData.resolveChunkUrl(id, chunk);
      script.async = true;
      script.setAttribute('data-hfx-sidecar', entryKey);
      var fail = function (reason) {
        var pending = waiters[entryKey] || [];
        delete waiters[entryKey];
        for (var index = 0; index < pending.length; index += 1) {
          pending[index](new Error(reason + ': ' + script.src));
        }
      };
      script.onerror = function () {
        fail('Sidecar failed to load');
      };
      script.onload = function () {
        // A sidecar that loaded but did not register this id/chunk would
        // otherwise leave its callers waiting forever.
        if (!registry[entryKey]) {
          fail('Sidecar did not register ' + entryKey);
        }
      };
      document.head.appendChild(script);
    }
  };

  function text(value) {
    return document.createTextNode(String(value));
  }

  function renderSummary(payload) {
    var target = document.getElementById('probe-summary');
    if (!target || !payload) {
      return;
    }
    var failing = 0;
    for (var index = 0; index < payload.rows.length; index += 1) {
      if (payload.rows[index][2] !== 'Healthy') {
        failing += 1;
      }
    }
    target.textContent = payload.rows.length + ' probes, ' + failing + ' need attention.';
  }

  function renderTable(payload) {
    var body = document.querySelector('#probe-table tbody');
    if (!body || !payload) {
      return;
    }
    for (var rowIndex = 0; rowIndex < payload.rows.length; rowIndex += 1) {
      var row = document.createElement('tr');
      for (var cellIndex = 0; cellIndex < payload.rows[rowIndex].length; cellIndex += 1) {
        var cell = document.createElement('td');
        cell.appendChild(text(payload.rows[rowIndex][cellIndex]));
        row.appendChild(cell);
      }
      body.appendChild(row);
    }
  }

  function renderRuntimeLinks() {
    var list = document.getElementById('runtime-link-list');
    if (!list) {
      return;
    }
    var item = document.createElement('li');
    var anchor = document.createElement('a');
    anchor.id = 'runtime-link-page2';
    anchor.setAttribute('href', 'page2.html#section');
    anchor.appendChild(text('Probe details (created at runtime)'));
    item.appendChild(anchor);
    list.appendChild(item);
  }

  var themeTokenPattern = /^#(?:[0-9a-f]{3,4}|[0-9a-f]{6}|[0-9a-f]{8})$/i;

  function applyHostTheme(message) {
    var root = document.documentElement;
    root.setAttribute('data-report-theme', message.isDark === true ? 'dark' : 'light');
    var tokens = message.tokens || {};
    var map = {
      accent: '--report-accent',
      surface: '--report-surface',
      text: '--report-text'
    };
    for (var name in map) {
      if (!Object.prototype.hasOwnProperty.call(map, name)) {
        continue;
      }
      if (typeof tokens[name] === 'string' && themeTokenPattern.test(tokens[name])) {
        root.style.setProperty(map[name], tokens[name]);
      } else {
        // Missing or invalid tokens fall back to the stylesheet value instead
        // of keeping a value from an earlier message.
        root.style.removeProperty(map[name]);
      }
    }
  }

  // Proposed contract: { type: 'uhv-host-theme', version: 1, isDark, tokens }.
  global.addEventListener('message', function (event) {
    var data = event.data;
    if (event.source !== global.parent || !data || data.type !== 'uhv-host-theme' || data.version !== 1) {
      return;
    }
    applyHostTheme(data);
  });

  document.addEventListener('DOMContentLoaded', function () {
    var page = document.body.getAttribute('data-report-page');
    if (page === 'index') {
      renderSummary(HfxData.get('probes', 0));
      renderRuntimeLinks();
    }
    var lazyDataset = document.body.getAttribute('data-lazy-dataset');
    if (lazyDataset) {
      HfxData.load(lazyDataset, document.body.getAttribute('data-lazy-chunk') || '0', function (error, payload) {
        if (!error) {
          renderTable(payload);
        }
      });
    }
  });

  global.HfxData = HfxData;
})(window);
