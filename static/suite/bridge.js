/* Shared bridge for the five embedded analysis modules.
 *  1. window.__LS  -> the shell's shared store when framed (so backups cover every module),
 *                     plain localStorage when a module is opened on its own.
 *  2. window.claude -> a browser-side stand-in for the claude.ai "downloads" and "permissions"
 *                     capabilities the modules were written against. save() is an ordinary
 *                     <a download>, so the modules work on any host, not only inside claude.ai.
 */
(function () {
  'use strict';
  var S = null;
  try { if (window.parent !== window) S = window.parent.__SUITE || null; } catch (e) { /* cross-origin parent */ }
  var m = {};
  var fb = {
    getItem: function (k) { return Object.prototype.hasOwnProperty.call(m, k) ? m[k] : null; },
    setItem: function (k, v) { m[k] = String(v); },
    removeItem: function (k) { delete m[k]; }
  };
  var ls = null;
  try { ls = window.localStorage; } catch (e) { /* blocked */ }
  window.__LS = S ? S.storage : (ls || fb);

  if (window.claude && typeof window.claude.use === 'function') return;

  function save(opts) {
    return new Promise(function (resolve, reject) {
      try {
        var d = opts && opts.data;
        var blob = d instanceof Blob ? d : new Blob([d == null ? '' : d], { type: 'application/octet-stream' });
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = (opts && opts.filename) || 'download';
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        setTimeout(function () { a.remove(); URL.revokeObjectURL(url); }, 4000);
        resolve({ filename: a.download, size: blob.size });
      } catch (err) {
        var e2 = new Error((err && err.message) || 'download failed');
        e2.code = 'unavailable';
        reject(e2);
      }
    });
  }
  var caps = {
    downloads: { save: save },
    permissions: {
      state: function () { return Promise.resolve('granted'); },
      manage: function () { return Promise.resolve(); }
    }
  };
  window.claude = {
    shim: true,
    use: function (name) {
      if (caps[name]) return Promise.resolve(caps[name]);
      var e = new Error('capability not available: ' + name);
      e.code = 'unavailable';
      return Promise.reject(e);
    }
  };
})();
