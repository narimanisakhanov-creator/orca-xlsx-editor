/* Boot watchdog for the editor frame.
 *
 * index.html starts this classic script before the module bundle, so a frame whose hashed assets
 * are missing — a plugin installed from a source checkout without `npm run build` — says so instead
 * of spinning on the loader forever. src/main.js sets window.__xlsxReady once it has mounted.
 *
 * The overlay text is compared with what this script saw on load: when the app already replaced the
 * loader with a real error, that message is the better one and the watchdog stays quiet. */
;(function () {
  'use strict'

  var DELAY_MS = 8000
  var boot = document.getElementById('boot')
  var loading = boot ? boot.textContent : null

  window.setTimeout(function () {
    if (window.__xlsxReady) return
    if (!boot || boot.textContent !== loading) return
    boot.hidden = false
    boot.textContent = 'Editor assets are missing. Run `npm run build` in the plugin folder.'
    console.error('[xlsx] the editor did not start within ' + DELAY_MS + ' ms — run `npm run build` in the plugin folder')
  }, DELAY_MS)
})()
