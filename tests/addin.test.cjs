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

test("the complete accumulated frame receives one full-image GPU glass pass", () => {
  const html = read("content.html");
  const css = read("content.css");
  const source = read("content.js");

  assert.doesNotMatch(html, /id="powerpoint-glass"|<fe[A-Z]/);
  assert.doesNotMatch(css, /filter:\s*url/);
  assert.match(css, /canvas[\s\S]*opacity:\s*0\.65/);
  assert.match(source, /createCanvas\(size\.width, size\.height, sketch\.WEBGL\)/);
  assert.match(source, /createGraphics\(size\.width, size\.height\)/);
  assert.match(source, /createShader\(vertexShaderSource, fragmentShaderSource\)/);
  assert.match(source, /glassShader\.setUniform\("uTrailTexture", trailLayer\)/);
  assert.match(source, /uniform sampler2D uTrailTexture/);
  assert.match(source, /vec3 sampleGlassCell/);
  assert.match(source, /vec2 glassCell = floor\(gl_FragCoord\.xy \/ 5\.0\)/);
  assert.match(source, /float refractedEnergy/);
  assert.match(source, /float facetLight/);
  assert.match(source, /vec3 localContrast/);
  assert.match(source, /vec3 sharpenedFrame/);
  assert.match(source, /vec2 pixelJitter/);
  assert.match(source, /flowWarp \* 1\.6 \+ cellJitter \* 0\.9 \+ pixelJitter \* 1\.0/);
  assert.match(source, /colorBlend \* 0\.38/);
  assert.doesNotMatch(source, /neighborhoodAverage|gatherGlassNeighborhood/);
  assert.doesNotMatch(source, /trailMask\s*\*/);
  assert.doesNotMatch(source, /vec3\(silverEdge \+ glassChip\)/);
  assert.doesNotMatch(source, /particleOpacity/);
  assert.match(source, /255\s*\*\s*fadeRatio/);
  assert.match(source, /speed:\s*0\.2/);
  assert.match(source, /this\.life -= 0\.01666/);
  assert.match(source, /this\.velocity\.mult\(options\.speed\)/);
  assert.doesNotMatch(source, /frameScale/);
  assert.match(source, /noiseScale:\s*800/);
  assert.doesNotMatch(html, /glass-overlay/);
  assert.doesNotMatch(css, /glass-overlay\.png/);
  assert.equal(fs.existsSync(path.join(root, "assets", "glass-overlay.png")), false);
});

test("trail history clears quickly with a time-based half-life", () => {
  const source = read("content.js");

  assert.match(source, /particleCount:\s*400/);
  assert.match(source, /trailHalfLifeSeconds:\s*1\.5/);
  assert.match(source, /const deltaSeconds = sketch\.deltaTime \/ 1000/);
  assert.match(
    source,
    /const fadeAmount = 1 - Math\.pow\(0\.5, deltaSeconds \/ options\.trailHalfLifeSeconds\)/
  );
  assert.match(source, /trailLayer\.fill\(10, 10, 10, 255 \* fadeAmount\)/);
  assert.match(source, /trailLayer\.rect\(0, 0, trailLayer\.width, trailLayer\.height\)/);
  assert.match(source, /255 \* fadeRatio/);
  assert.match(source, /vec3 toneMapped/);
});

test("particles retain the original point-based p5 rendering", () => {
  const source = read("content.js");

  assert.match(source, /colors:\s*\[\s*\[255, 255, 255\]/);
  assert.match(source, /this\.iterations = Math\.ceil/);
  assert.match(source, /layer\.ellipse\(this\.position\.x, this\.position\.y, options\.particleSize/);
  assert.doesNotMatch(source, /layer\.line\(/);
  assert.doesNotMatch(source, /const particleColor = sketch\.color/);
});

test("host resize preserves particle state and the accumulated frame", () => {
  const source = read("content.js");

  assert.doesNotMatch(source, /Math\.min\(sketch\.deltaTime, 50\)/);
  assert.match(source, /if \(size\.width === sketch\.width && size\.height === sketch\.height\)/);
  assert.match(source, /const previousFrame = trailLayer\.get\(\)/);
  assert.match(source, /sketch\.resizeCanvas\(size\.width, size\.height, true\)/);
  assert.match(source, /trailLayer\.resizeCanvas\(size\.width, size\.height\)/);
  assert.match(source, /trailLayer\.image\(previousFrame, 0, 0, size\.width, size\.height\)/);
  assert.match(source, /particle\.scalePosition\(scaleX, scaleY\)/);
  assert.doesNotMatch(source, /particles\.forEach\(\(particle\) => particle\.respawn\(true\)\)/);
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
