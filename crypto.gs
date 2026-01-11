/**
 * Poparnassus Sheets Lib - Crypto helpers
 */

function fetchJson(url) {
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: { "Accept": "application/json" }
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`HTTP ${code}: ${text}`);
  }
  return JSON.parse(text);
}
