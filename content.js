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
  trailHalfLifeSeconds: 1.5,
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

  vec3 sampleGlassCell(vec2 uv, vec2 pixel) {
    return texture2D(uTrailTexture, clampUv(uv, pixel)).rgb;
  }

  void main() {
    vec2 uv = vec2(vTexCoord.x, 1.0 - vTexCoord.y);
    vec2 pixel = 1.0 / uResolution;

    vec2 glassCell = floor(gl_FragCoord.xy / 5.0);
    vec2 cellJitter = vec2(
      hash21(glassCell + vec2(11.0, 3.0)),
      hash21(glassCell + vec2(5.0, 19.0))
    ) - 0.5;
    vec2 microCell = floor(gl_FragCoord.xy / 2.0);
    vec2 pixelJitter = vec2(
      hash21(microCell + vec2(37.0, 7.0)),
      hash21(microCell + vec2(13.0, 41.0))
    ) - 0.5;
    vec2 flowWarp = vec2(
      valueNoise(uv * vec2(10.0, 6.0) + vec2(3.1, 7.7)),
      valueNoise(uv * vec2(8.0, 7.0) + vec2(19.4, 2.3))
    ) - 0.5;
    vec2 refractedUv = clampUv(
      uv + (flowWarp * 1.6 + cellJitter * 0.9 + pixelJitter * 1.0) * pixel,
      pixel
    );
    vec2 secondaryUv = clampUv(
      uv + (-flowWarp.yx * 1.3 - cellJitter * 0.7 - pixelJitter * 0.8) * pixel,
      pixel
    );
    vec3 primarySample = sampleGlassCell(refractedUv, pixel);
    vec3 secondarySample = sampleGlassCell(secondaryUv, pixel);
    float colorBlend = smoothstep(0.18, 0.82, valueNoise(uv * vec2(48.0, 27.0) + 9.0));
    vec3 refractedFrame = mix(primarySample, secondarySample, colorBlend * 0.38);
    float refractedEnergy = max(luminance(refractedFrame) - 0.035, 0.0);

    vec2 gradientOffset = pixel * 1.5;
    vec3 leftSample = sampleGlassCell(refractedUv - vec2(gradientOffset.x, 0.0), pixel);
    vec3 rightSample = sampleGlassCell(refractedUv + vec2(gradientOffset.x, 0.0), pixel);
    vec3 downSample = sampleGlassCell(refractedUv - vec2(0.0, gradientOffset.y), pixel);
    vec3 upSample = sampleGlassCell(refractedUv + vec2(0.0, gradientOffset.y), pixel);
    vec3 localAverage = (leftSample + rightSample + downSample + upSample) * 0.25;
    vec3 localContrast = refractedFrame - localAverage;
    vec3 sharpenedFrame = clamp(refractedFrame + localContrast * 1.15, 0.0, 1.0);
    float leftEnergy = luminance(leftSample);
    float rightEnergy = luminance(rightSample);
    float downEnergy = luminance(downSample);
    float upEnergy = luminance(upSample);
    vec2 imageGradient = vec2(rightEnergy - leftEnergy, upEnergy - downEnergy);
    float edgeStrength = clamp(length(imageGradient) * 2.8, 0.0, 1.0);

    vec3 surfaceNormal = normalize(vec3(-imageGradient * 4.0, 0.9));
    vec3 lightDirection = normalize(vec3(-0.55, 0.7, 0.8));
    float facetLight = 0.5 + 0.5 * dot(surfaceNormal, lightDirection);
    float cellGrain = hash21(glassCell + floor(flowWarp * 11.0));
    float fineGrain = hash21(floor(gl_FragCoord.xy / 1.5) + vec2(29.0, 13.0));
    float grainGain = mix(0.74, 1.24, fineGrain * 0.72 + cellGrain * 0.28);
    float imageCoverage = smoothstep(0.012, 0.16, refractedEnergy + edgeStrength * 0.12);

    float strongestChannel = max(sharpenedFrame.r, max(sharpenedFrame.g, sharpenedFrame.b));
    vec3 sourceHue = strongestChannel > 0.001 ? sharpenedFrame / strongestChannel : vec3(0.0);
    vec3 glassHue = mix(vec3(0.5, 0.7, 0.98), sourceHue, 0.78);

    vec3 posterizedFrame = floor(sharpenedFrame * 7.0 + 0.5) / 7.0;
    vec3 crystallineColor = mix(sharpenedFrame, posterizedFrame, 0.42);
    vec3 glassDetail = abs(localContrast) * (0.44 + edgeStrength * 0.34);
    float directionalEmboss = max(dot(imageGradient, normalize(vec2(-0.7, 0.7))), 0.0);
    vec3 rawColor =
      crystallineColor * grainGain +
      glassDetail +
      glassHue * imageCoverage * edgeStrength * facetLight * 0.16 +
      vec3(imageCoverage * directionalEmboss * 0.16);
    vec3 toneMapped = clamp(rawColor, 0.0, 1.0);
    float mappedLuma = luminance(toneMapped);
    toneMapped = mix(vec3(mappedLuma), toneMapped, 1.08);
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
      this.life = 0;
      this.respawn(true);
    }

    move(iterations) {
      this.life -= 0.01666;
      if (this.life < 0) {
        this.respawn();
      }

      while (iterations > 0) {
        const angle =
          sketch.noise(this.position.x / options.noiseScale, this.position.y / options.noiseScale) *
          sketch.TWO_PI *
          options.noiseScale;

        this.velocity.set(sketch.cos(angle), sketch.sin(angle));
        this.velocity.mult(0.2);
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

    for (let index = 1; index < particles.length; index += 1) {
      const particle = particles[index];
      const iterations = sketch.map(index, 0, options.particleCount, 5, 1);

      particle.move(iterations);
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
