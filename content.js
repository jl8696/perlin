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
  particleCount: 500,
  particleSize: 2,
  noiseScale: 800,
  speed: 0.34,
  trailHalfLifeSeconds: 18,
  trailOpacity: 0.24,
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

  void main() {
    vec2 uv = vec2(vTexCoord.x, 1.0 - vTexCoord.y);
    vec2 pixel = 1.0 / uResolution;

    float warpX = valueNoise(uv * vec2(8.0, 4.5) + vec2(3.1, 7.7));
    float warpY = valueNoise(uv * vec2(7.0, 5.5) + vec2(19.4, 2.3));
    vec2 refraction = (vec2(warpX, warpY) - 0.5) * pixel * 4.0;
    vec2 sampleUv = clamp(uv + refraction, pixel, vec2(1.0) - pixel);

    vec3 center = texture2D(uTrailTexture, sampleUv).rgb;
    float centerLuma = luminance(center);
    float north = luminance(texture2D(uTrailTexture, sampleUv + vec2(0.0, pixel.y)).rgb);
    float south = luminance(texture2D(uTrailTexture, sampleUv - vec2(0.0, pixel.y)).rgb);
    float east = luminance(texture2D(uTrailTexture, sampleUv + vec2(pixel.x, 0.0)).rgb);
    float west = luminance(texture2D(uTrailTexture, sampleUv - vec2(pixel.x, 0.0)).rgb);
    float edgeStrength = clamp(
      (abs(centerLuma - north) + abs(centerLuma - south) +
       abs(centerLuma - east) + abs(centerLuma - west)) * 3.4,
      0.0,
      1.0
    );

    vec2 grainCell = floor(gl_FragCoord.xy);
    float fineGrain = hash21(grainCell);
    float clusteredGrain = hash21(floor(grainCell * 0.43) + vec2(31.0, 17.0));
    float stipple = smoothstep(
      0.54,
      0.94,
      fineGrain * 0.7 + clusteredGrain * 0.3 + edgeStrength * 0.25
    );

    float trailMask = smoothstep(0.055, 0.18, centerLuma + edgeStrength * 0.7);
    float sparkleSignal = clusteredGrain * 0.65 + fineGrain * 0.35 + edgeStrength * 0.08;
    float sparkle = trailMask * smoothstep(0.88, 0.98, sparkleSignal);
    float strongestChannel = max(center.r, max(center.g, center.b));
    float weakestChannel = min(center.r, min(center.g, center.b));
    float centerChroma = strongestChannel - weakestChannel;
    vec3 trailHue = strongestChannel > 0.001 ? center / strongestChannel : vec3(0.0);
    vec3 glassHue = mix(
      vec3(0.46, 0.68, 0.95),
      trailHue,
      smoothstep(0.06, 0.28, centerChroma)
    );

    vec3 dimmedTrail = mix(center, glassHue * centerLuma, 0.55) * 0.16;
    float coloredFacet = trailMask * (0.04 + stipple * 0.24 + edgeStrength * 0.14);
    float silverFacet = trailMask * edgeStrength * (0.08 + stipple * 0.32);
    vec3 glassColor = glassHue * coloredFacet + vec3(silverFacet);

    vec3 rawColor = dimmedTrail + glassColor;
    vec3 toneMapped = rawColor / (vec3(0.68) + rawColor) * 1.28;
    float mappedLuma = luminance(toneMapped);
    toneMapped = mix(vec3(mappedLuma), toneMapped, 1.8) + vec3(sparkle * 0.85);
    vec3 finalColor = pow(clamp(toneMapped, 0.0, 1.0), vec3(1.08));
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
      this.previousPosition = sketch.createVector(0, 0);
      this.color = options.colors[index % options.colors.length];
      this.iterations = Math.ceil(sketch.map(index, 0, options.particleCount, 5, 1));
      this.life = 0;
      this.respawn(true);
    }

    move(frameScale) {
      this.previousPosition.set(this.position);
      this.life -= Math.min(sketch.deltaTime, 50) / 1000;
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
        this.velocity.mult(options.speed * frameScale);
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
      this.previousPosition.set(this.position);
      this.life = randomizeLife ? sketch.random(0, options.lifeSeconds) : options.lifeSeconds;
    }

    scalePosition(scaleX, scaleY) {
      this.position.x *= scaleX;
      this.position.y *= scaleY;
      this.previousPosition.x *= scaleX;
      this.previousPosition.y *= scaleY;
    }

    display(layer) {
      layer.line(
        this.previousPosition.x,
        this.previousPosition.y,
        this.position.x,
        this.position.y
      );
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
    trailLayer.strokeCap(sketch.ROUND);
    trailLayer.strokeWeight(options.particleSize);
    trailLayer.noFill();
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
    const deltaSeconds = Math.min(sketch.deltaTime, 50) / 1000;
    const frameScale = sketch.min(sketch.deltaTime / 16.6667, 1.5);
    const fadeAmount = 1 - Math.pow(0.5, deltaSeconds / options.trailHalfLifeSeconds);

    trailLayer.push();
    trailLayer.noStroke();
    trailLayer.fill(10, 10, 10, 255 * fadeAmount);
    trailLayer.rect(0, 0, trailLayer.width, trailLayer.height);
    trailLayer.pop();

    for (let index = 0; index < particles.length; index += 1) {
      const particle = particles[index];

      particle.move(frameScale);
      particle.checkEdge();

      let fadeRatio = sketch.min((particle.life * 5) / options.lifeSeconds, 1);
      fadeRatio = sketch.min(
        ((options.lifeSeconds - particle.life) * 5) / options.lifeSeconds,
        fadeRatio
      );

      trailLayer.stroke(
        particle.color[0],
        particle.color[1],
        particle.color[2],
        255 * options.trailOpacity * fadeRatio
      );
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
