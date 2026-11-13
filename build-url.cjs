/* global URL, module */

function resolveProductionUrl(value) {
  if (!value) {
    throw new Error("ADDIN_PROD_URL is required for production builds.");
  }

  const url = new URL(value);
  if (url.protocol !== "https:") {
    throw new Error("ADDIN_PROD_URL must use HTTPS.");
  }

  if (!url.pathname.endsWith("/")) {
    url.pathname += "/";
  }

  return url.href;
}

function transformManifest(content, dev, productionUrl) {
  if (dev) {
    return content;
  }

  const baseUrl = resolveProductionUrl(productionUrl);
  const manifest = JSON.parse(content);
  manifest.extensions[0].contentRuntimes[0].code.page = new URL("content.html", baseUrl).href;
  manifest.validDomains = [new URL(baseUrl).hostname];
  return JSON.stringify(manifest, null, 2);
}

module.exports = { resolveProductionUrl, transformManifest };
