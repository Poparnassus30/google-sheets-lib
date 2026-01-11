/**
 * Simple stockage côté projet (PropertiesService)
 */

function kvSet(key, value) {
  PropertiesService.getScriptProperties().setProperty(key, String(value));
  return true;
}

function kvGet(key, defaultValue) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return (v === null || v === undefined) ? defaultValue : v;
}
