import p5 from "p5";

/* global Office, console, document */

const options = {
  background: "#0a0a0a",
  colors: [
    [255, 255, 255],
    [7, 153, 242],
    [69, 33, 124],
  ],
  lifeSeconds: 10,
  particleCount: 400,
  particleSize: 2,
  noiseScale: 800,
  speed: 0.2,
  trailHalfLifeSeconds: 3,
};

const vertexShaderSource = `
  precision highp float;

  attribute vec3 aPosition;
  attribute vec2 aTexCoord;
  uniform mat4 uModelViewMatrix;
  uniform mat4 uProjectionMatrix;
  varying vec2 vTexCoord;

  void main() {
    vTexCoord = aTexCoord;
    gl_Position = uProjectionMatrix * uModelViewMatrix * vec4(aPosition, 1.0);
  }
`;

const fragmentShaderSource = `
  precision highp float;

  uniform sampler2D uTrailTexture;
  uniform vec2 uResolution;
  varying vec2 vTexCoord;

  float hash21(vec2 point) {
    point = fract(point * vec2(123.34, 456.21));
    point += dot(point, point + 45.32);
    return fract(point.x * point.y);
  }

  float valueNoise(vec2 point) {
    vec2 cell = floor(point);
    vec2 fraction = fract(point);
    fraction = fraction * fraction * (3.0 - 2.0 * fraction);

    float bottomLeft = hash21(cell);
    float bottomRight = hash21(cell + vec2(1.0, 0.0));
    float topLeft = hash21(cell + vec2(0.0, 1.0));
    float topRight = hash21(cell + vec2(1.0, 1.0));

    return mix(
      mix(bottomLeft, bottomRight, fraction.x),
      mix(topLeft, topRight, fraction.x),
      fraction.y
    );
  }

  float luminance(vec3 color) {
    return dot(color, vec3(0.2126, 0.7152, 0.0722));
  }

  vec2 clampUv(vec2 uv, vec2 pixel) {
    return clamp(uv, pixel * 8.0, vec2(1.0) - pixel * 8.0);
  }

  vec3 gatherGlassNeighborhood(vec2 uv, vec2 pixel) {
    vec2 nearOffset = pixel * 2.0;
    vec2 wideOffset = pixel * 6.0;
    vec3 center = texture2D(uTrailTexture, clampUv(uv, pixel)).rgb;
    vec3 horizontal =
      texture2D(uTrailTexture, clampUv(uv + vec2(nearOffset.x, 0.0), pixel)).rgb +
      texture2D(uTrailTexture, clampUv(uv - vec2(nearOffset.x, 0.0), pixel)).rgb;
    vec3 vertical =
      texture2D(uTrailTexture, clampUv(uv + vec2(0.0, nearOffset.y), pixel)).rgb +
      texture2D(uTrailTexture, clampUv(uv - vec2(0.0, nearOffset.y), pixel)).rgb;
    vec3 diagonal =
      texture2D(uTrailTexture, clampUv(uv + wideOffset, pixel)).rgb +
      texture2D(uTrailTexture, clampUv(uv - wideOffset, pixel)).rgb +
      texture2D(uTrailTexture, clampUv(uv + vec2(wideOffset.x, -wideOffset.y), pixel)).rgb +
      texture2D(uTrailTexture, clampUv(uv + vec2(-wideOffset.x, wideOffset.y), pixel)).rgb;
    vec3 neighborhoodAverage = (center * 4.0 + horizontal + vertical + diagonal * 0.35) / 9.4;
    return mix(center, neighborhoodAverage, 0.52);
  }

  void main() {
    vec2 uv = vec2(vTexCoord.x, 1.0 - vTexCoord.y);
    vec2 pixel = 1.0 / uResolution;

    vec2 glassCell = floor(gl_FragCoord.xy / 5.0);
    vec2 cellJitter = vec2(
      hash21(glassCell + vec2(11.0, 3.0)),
      hash21(glassCell + vec2(5.0, 19.0))
    ) - 0.5;
    vec2 flowWarp = vec2(
      valueNoise(uv * vec2(10.0, 6.0) + vec2(3.1, 7.7)),
      valueNoise(uv * vec2(8.0, 7.0) + vec2(19.4, 2.3))
    ) - 0.5;
    vec2 refractedUv = clampUv(uv + (flowWarp * 3.0 + cellJitter * 1.5) * pixel, pixel);

    vec3 sourceCenter = texture2D(uTrailTexture, refractedUv).rgb;
    vec3 refractedFrame = gatherGlassNeighborhood(refractedUv, pixel);
    float refractedEnergy = max(luminance(refractedFrame) - 0.035, 0.0);

    vec2 gradientOffset = pixel * 2.0;
    float leftEnergy = luminance(texture2D(uTrailTexture, clampUv(refractedUv - vec2(gradientOffset.x, 0.0), pixel)).rgb);
    float rightEnergy = luminance(texture2D(uTrailTexture, clampUv(refractedUv + vec2(gradientOffset.x, 0.0), pixel)).rgb);
    float downEnergy = luminance(texture2D(uTrailTexture, clampUv(refractedUv - vec2(0.0, gradientOffset.y), pixel)).rgb);
    float upEnergy = luminance(texture2D(uTrailTexture, clampUv(refractedUv + vec2(0.0, gradientOffset.y), pixel)).rgb);
    vec2 imageGradient = vec2(rightEnergy - leftEnergy, upEnergy - downEnergy);
    float edgeStrength = clamp(length(imageGradient) * 2.8, 0.0, 1.0);

    vec3 surfaceNormal = normalize(vec3(-imageGradient * 4.0, 0.9));
    vec3 lightDirection = normalize(vec3(-0.55, 0.7, 0.8));
    float facetLight = 0.5 + 0.5 * dot(surfaceNormal, lightDirection);
    float cellGrain = hash21(glassCell + floor(flowWarp * 11.0));
    float fineGrain = hash21(floor(gl_FragCoord.xy / 2.0) + vec2(29.0, 13.0));
    float grainGain = mix(0.72, 1.22, fineGrain * 0.7 + cellGrain * 0.3);
    float imageCoverage = smoothstep(0.012, 0.16, refractedEnergy + edgeStrength * 0.12);

    float strongestChannel = max(refractedFrame.r, max(refractedFrame.g, refractedFrame.b));
    vec3 sourceHue = strongestChannel > 0.001 ? refractedFrame / strongestChannel : vec3(0.0);
    vec3 glassHue = mix(vec3(0.5, 0.7, 0.98), sourceHue, 0.78);

    vec3 localContrast = sourceCenter - refractedFrame;
    vec3 glassDetail = abs(localContrast) * (0.32 + edgeStrength * 0.28);
    float directionalEmboss = max(dot(imageGradient, normalize(vec2(-0.7, 0.7))), 0.0);
    vec3 rawColor =
      refractedFrame * grainGain +
      glassDetail +
      glassHue * refractedEnergy * facetLight * 0.08 +
      vec3(imageCoverage * directionalEmboss * 0.12);
    vec3 toneMapped = clamp(rawColor, 0.0, 1.0);
    float mappedLuma = luminance(toneMapped);
    toneMapped = mix(vec3(mappedLuma), toneMapped, 0.92);
    vec3 finalColor = pow(clamp(toneMapped, 0.0, 1.0), vec3(1.02));
    gl_FragColor = vec4(finalColor, 1.0);
  }
`;

function createParticleSketch(sketch) {
  const particles = [];
  let trailLayer;
  let glassShader;

  class Particle {
    constructor(index) {
      this.velocity = sketch.createVector(0, 0);
      this.position = sketch.createVector(0, 0);
      this.color = options.colors[index % options.colors.length];
      this.iterations = Math.ceil(sketch.map(index, 0, options.particleCount, 5, 1));
      this.life = 0;
      this.respawn(true);
    }

    move() {
      this.life -= 0.01666;
      if (this.life < 0) {
        this.respawn();
      }

      let iterations = this.iterations;
      while (iterations > 0) {
        const angle =
          sketch.noise(this.position.x / options.noiseScale, this.position.y / options.noiseScale) *
          sketch.TWO_PI *
          options.noiseScale;

        this.velocity.set(sketch.cos(angle), sketch.sin(angle));
        this.velocity.mult(options.speed);
        this.position.add(this.velocity);
        iterations -= 1;
      }
    }

    checkEdge() {
      if (
        this.position.x > sketch.width ||
        this.position.x < 0 ||
        this.position.y > sketch.height ||
        this.position.y < 0
      ) {
        this.respawn();
      }
    }

    respawn(randomizeLife = false) {
      this.position.set(
        sketch.random(-50, sketch.width + 50),
        sketch.random(-50, sketch.height + 50)
      );
      this.life = randomizeLife ? sketch.random(0, options.lifeSeconds) : options.lifeSeconds;
    }

    scalePosition(scaleX, scaleY) {
      this.position.x *= scaleX;
      this.position.y *= scaleY;
    }

    display(layer) {
      layer.ellipse(this.position.x, this.position.y, options.particleSize, options.particleSize);
    }
  }

  function frameSize() {
    const host = document.getElementById("particle-background");
    return {
      host,
      width: Math.max(host.clientWidth, 1),
      height: Math.max(host.clientHeight, 1),
    };
  }

  function configureTrailLayer() {
    trailLayer.pixelDensity(1);
    trailLayer.noStroke();
  }

  sketch.setup = () => {
    const size = frameSize();
    const canvas = sketch.createCanvas(size.width, size.height, sketch.WEBGL);
    canvas.parent(size.host);
    sketch.pixelDensity(1);
    sketch.noStroke();
    sketch.rectMode(sketch.CENTER);

    trailLayer = sketch.createGraphics(size.width, size.height);
    configureTrailLayer();
    trailLayer.background(options.background);
    glassShader = sketch.createShader(vertexShaderSource, fragmentShaderSource);

    for (let index = 0; index < options.particleCount; index += 1) {
      particles.push(new Particle(index));
    }
  };

  sketch.draw = () => {
    const deltaSeconds = sketch.deltaTime / 1000;
    const fadeAmount = 1 - Math.pow(0.5, deltaSeconds / options.trailHalfLifeSeconds);

    trailLayer.push();
    trailLayer.noStroke();
    trailLayer.fill(10, 10, 10, 255 * fadeAmount);
    trailLayer.rect(0, 0, trailLayer.width, trailLayer.height);
    trailLayer.pop();

    for (let index = 0; index < particles.length; index += 1) {
      const particle = particles[index];

      particle.move();
      particle.checkEdge();

      let fadeRatio = sketch.min((particle.life * 5) / options.lifeSeconds, 1);
      fadeRatio = sketch.min(
        ((options.lifeSeconds - particle.life) * 5) / options.lifeSeconds,
        fadeRatio
      );

      trailLayer.fill(particle.color[0], particle.color[1], particle.color[2], 255 * fadeRatio);
      particle.display(trailLayer);
    }

    sketch.shader(glassShader);
    glassShader.setUniform("uTrailTexture", trailLayer);
    glassShader.setUniform("uResolution", [sketch.width, sketch.height]);
    sketch.rect(0, 0, sketch.width, sketch.height);
  };

  sketch.windowResized = () => {
    const size = frameSize();
    if (size.width === sketch.width && size.height === sketch.height) {
      return;
    }

    const previousFrame = trailLayer.get();
    const scaleX = size.width / sketch.width;
    const scaleY = size.height / sketch.height;

    particles.forEach((particle) => particle.scalePosition(scaleX, scaleY));
    sketch.resizeCanvas(size.width, size.height, true);
    trailLayer.resizeCanvas(size.width, size.height);
    trailLayer.background(options.background);
    trailLayer.image(previousFrame, 0, 0, size.width, size.height);
    configureTrailLayer();
  };
}

let animation;

function startAnimation() {
  if (!animation) {
    animation = new p5(createParticleSketch);
  }
}

function handleOfficeReadyFailure(error) {
  console.warn(
    "Office host initialization failed; starting the visual without Office APIs.",
    error
  );
  startAnimation();
}

if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady().then(startAnimation).catch(handleOfficeReadyFailure);
} else {
  startAnimation();
}
