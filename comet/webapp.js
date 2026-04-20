/**
 * webapp.js — GAS Web App entry point for COMET.
 * VERSION: 0.2.3
 *
 * Serves the single-page HTML shell and handles the include() helper
 * used by index.html to inject stylesheet and javascript partials.
 */

/**
 * Entry point for all HTTP GET requests to the deployed web app.
 * Routes based on the `view` query parameter — currently only the
 * main SPA is served; future phases will add the employee exception form.
 *
 * @param {GoogleAppsScript.Events.DoGet} event
 * @returns {GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(event) {
  const template = HtmlService.createTemplateFromFile('index');
  template.appTitle = COMET_APP_TITLE;
  return template.evaluate()
    .setTitle(COMET_APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Injects the content of another HTML file into the current template.
 * Used in index.html as <?!= include('stylesheet') ?> and
 * <?!= include('javascript') ?>.
 *
 * @param {string} filename — The HTML filename to include (without .html extension).
 * @returns {string} Raw HTML content of the included file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
