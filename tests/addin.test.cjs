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
  assert.doesNotMatch(html, /dat\.GUI|img-path|credits|class="logo"|Controls|Random|Save/);
});

test("particle runtime bundles p5 and waits for the Office host", () => {
  const source = read("content.js");

  assert.match(source, /import p5 from "p5"/);
  assert.match(source, /Office\.onReady/);
  assert.match(source, /Office\.onReady\(\)\.then\(startAnimation\)\.catch/);
  assert.match(source, /new p5\(/);
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

test("particle options retain the original values and naming", () => {
  const source = read("content.js");

  assert.match(source, /Background:\s*"#0a0a0a"/);
  assert.match(source, /Color1:\s*"#ffffff"/);
  assert.match(source, /Color2:\s*"#0799f2"/);
  assert.match(source, /Color3:\s*"#45217c"/);
  assert.match(source, /Length:\s*10/);
  assert.match(source, /Nums:\s*400/);
  assert.match(source, /Size:\s*2/);
  assert.match(source, /noiseScale:\s*800/);
  assert.match(source, /ColorMode:\s*"Normal"/);
});

test("setup and draw retain the original particle flow", () => {
  const source = read("content.js");

  assert.match(source, /backgroundColor = sketch\.color\(options\.Background\)/);
  assert.match(source, /sketch\.createCanvas\(size\.width, size\.height\)/);
  assert.match(source, /sketch\.background\(options\.Background\)/);
  assert.match(source, /for \(let index = 0; index < 2500; index \+= 1\)/);
  assert.match(source, /maxLife = options\.Length/);
  assert.match(source, /for \(let index = 1; index < options\.Nums; index \+= 1\)/);
  assert.match(source, /sketch\.map\(index, 0, options\.Nums, 5, 1\)/);
  assert.match(source, /particles\[index\]\.move\(iterations\)/);
  assert.match(source, /particles\[index\]\.checkEdge\(\)/);
  assert.match(source, /particles\[index\]\.life \* 5/);
  assert.match(source, /maxLife - particles\[index\]\.life/);
  assert.match(source, /sketch\.fill\([\s\S]*alpha \* fadeRatio/);
  assert.match(source, /particles\[index\]\.display\(radius\)/);
  assert.doesNotMatch(source, /WEBGL|createGraphics|createShader|fragmentShaderSource|trailLayer/);
});

test("particles retain the original movement and point rendering", () => {
  const source = read("content.js");

  assert.match(source, /this\.vel = sketch\.createVector\(0, 0\)/);
  assert.match(source, /this\.pos = sketch\.createVector\([\s\S]*sketch\.random\(-50, sketch\.width \+ 50\)/);
  assert.match(source, /this\.life = sketch\.random\(0, maxLife\)/);
  assert.match(source, /this\.life -= 0\.01666/);
  assert.match(source, /sketch\.noise\(this\.pos\.x \/ options\.noiseScale, this\.pos\.y \/ options\.noiseScale\)/);
  assert.match(source, /sketch\.TWO_PI\s*\*\s*options\.noiseScale/);
  assert.match(source, /this\.vel\.x = sketch\.cos\(angle\)/);
  assert.match(source, /this\.vel\.y = sketch\.sin\(angle\)/);
  assert.match(source, /this\.vel\.mult\(0\.2\)/);
  assert.match(source, /this\.pos\.add\(this\.vel\)/);
  assert.match(source, /sketch\.ellipse\(this\.pos\.x, this\.pos\.y, radius, radius\)/);
});

test("host events never restart the particle field", () => {
  const source = read("content.js");

  assert.doesNotMatch(source, /touchStarted|mousePressed|setup\(\);|respawn\(true\)/);
  assert.doesNotMatch(source, /windowResized/);
});

test("webpack builds the JavaScript runtime and the manifest names the add-in", () => {
  const webpack = read("webpack.config.js");
  const manifest = JSON.parse(read("manifest.json"));

  assert.match(webpack, /content:\s*\[?"\.\/content\.js"\]?/);
  assert.match(webpack, /hash:\s*!dev/);
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
