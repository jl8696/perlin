const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");

test("content page contains only the full-frame animation host", () => {
  const html = read("content.html");

  assert.match(html, /officeapis\.public\.onecdn\.static\.microsoft\/1\/office\.js/);
  assert.match(html, /href="content\.css"/);
  assert.match(html, /id="particle-background"/);
  assert.doesNotMatch(html, /Get slide details|Welcome|pnptelemetry/);
  assert.doesNotMatch(html, /cdn\.jsdelivr\.net\/npm\/p5/);
});

test("particle runtime bundles p5 and waits for the Office host", () => {
  const source = read("content.js");

  assert.match(source, /import p5 from "p5"/);
  assert.match(source, /Office\.onReady/);
  assert.match(source, /Office\.onReady\(\)\.then\(startAnimation\)\.catch/);
  assert.match(source, /new p5\(/);
  assert.match(source, /windowResized/);
  assert.doesNotMatch(source, /setTimeout\(startAnimation/);
});

test("animation frame and canvas fill the content add-in", () => {
  const css = read("content.css");

  assert.match(css, /html[\s\S]*body[\s\S]*#particle-background/);
  assert.match(css, /width:\s*100%/);
  assert.match(css, /height:\s*100%/);
  assert.match(css, /overflow:\s*hidden/);
  assert.match(css, /canvas[\s\S]*position:\s*fixed/);
});

test("particles render at 30 percent through a live glass filter", () => {
  const html = read("content.html");
  const css = read("content.css");

  assert.match(html, /id="powerpoint-glass"/);
  assert.match(html, /feTurbulence/);
  assert.match(html, /feDisplacementMap/);
  assert.match(css, /canvas[\s\S]*opacity:\s*0\.3/);
  assert.match(css, /canvas[\s\S]*filter:\s*url\("#powerpoint-glass"\)/);
  assert.doesNotMatch(html, /glass-overlay/);
  assert.doesNotMatch(css, /glass-overlay\.png/);
  assert.equal(fs.existsSync(path.join(root, "assets", "glass-overlay.png")), false);
});

test("webpack builds the JavaScript runtime and the manifest names the add-in", () => {
  const webpack = read("webpack.config.js");
  const manifest = JSON.parse(read("manifest.json"));

  assert.match(webpack, /content:\s*\[?"\.\/content\.js"\]?/);
  assert.equal(manifest.name.short, "Particle Background");
  assert.equal(manifest.name.full, "Live Particle Background for PowerPoint");
  assert.match(manifest.description.short, /live p5\.js particle background/i);
});

test("lint targets the root-level source files used by this sample", () => {
  const packageJson = JSON.parse(read("package.json"));

  assert.match(packageJson.scripts.lint, /--files/);
  assert.match(packageJson.scripts.lint, /content\.js/);
  assert.match(packageJson.scripts.lint, /webpack\.config\.js/);
});

test("production builds never fall back to Microsoft's sample host", () => {
  const webpack = read("webpack.config.js");

  assert.doesNotMatch(webpack, /officedev\.github\.io/);
  assert.match(webpack, /ADDIN_PROD_URL/);
  assert.doesNotMatch(webpack, /ADDIN_PROD_URL \|\| urlDev/);
});

test("production URLs are required, HTTPS, and normalized", () => {
  let buildUrl;
  assert.doesNotThrow(() => {
    buildUrl = require("../build-url.cjs");
  });

  assert.throws(() => buildUrl.resolveProductionUrl(), /ADDIN_PROD_URL is required/);
  assert.throws(
    () => buildUrl.resolveProductionUrl("http://example.com/addin"),
    /must use HTTPS/
  );
  assert.equal(
    buildUrl.resolveProductionUrl("https://example.com/addin"),
    "https://example.com/addin/"
  );
});

test("production manifest transformation updates the page and authorized domain", () => {
  const { transformManifest } = require("../build-url.cjs");
  const source = JSON.stringify({
    validDomains: ["localhost"],
    extensions: [{ contentRuntimes: [{ code: { page: "https://localhost:3000/content.html" } }] }],
  });
  const transformed = JSON.parse(
    transformManifest(source, false, "https://example.com/addin/")
  );

  assert.equal(
    transformed.extensions[0].contentRuntimes[0].code.page,
    "https://example.com/addin/content.html"
  );
  assert.deepEqual(transformed.validDomains, ["example.com"]);
});

test("manifest uses least privilege and package-ready icon sizes", () => {
  const manifest = JSON.parse(read("manifest.json"));

  assert.deepEqual(manifest.authorization.permissions, {});
  assert.equal(manifest.icons.color, "assets/icon-192.png");
  assert.ok(fs.existsSync(path.join(root, "assets", "icon-192.png")));
});

test("fresh-start and app-package scripts cover setup and deployment", () => {
  const packageJson = JSON.parse(read("package.json"));
  const startScript = read("start-local.ps1");
  const packageScript = read("package-app.ps1");

  assert.match(startScript, /node_modules/);
  assert.match(startScript, /npm install/);
  assert.match(packageJson.scripts["package:addin"], /package-app\.ps1/);
  assert.match(packageScript, /particle-background-app-package\.zip/);
  assert.equal(packageJson.engines.node, ">=22.15.0");
});
