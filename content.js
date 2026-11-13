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
  speed: 0.34,
};

function createParticleSketch(sketch) {
  const particles = [];
  let backgroundColor;

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

    display() {
      sketch.line(
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

  sketch.setup = () => {
    const size = frameSize();
    const canvas = sketch.createCanvas(size.width, size.height);
    canvas.parent(size.host);
    sketch.pixelDensity(1);
    sketch.strokeCap(sketch.ROUND);
    sketch.strokeWeight(options.particleSize);
    sketch.noFill();

    backgroundColor = sketch.color(options.background);
    sketch.background(backgroundColor);

    for (let index = 0; index < options.particleCount; index += 1) {
      particles.push(new Particle(index));
    }
  };

  sketch.draw = () => {
    const frameScale = sketch.min(sketch.deltaTime / 16.6667, 1.5);

    for (let index = 0; index < particles.length; index += 1) {
      const particle = particles[index];

      particle.move(frameScale);
      particle.checkEdge();

      let fadeRatio = sketch.min((particle.life * 5) / options.lifeSeconds, 1);
      fadeRatio = sketch.min(
        ((options.lifeSeconds - particle.life) * 5) / options.lifeSeconds,
        fadeRatio
      );

      sketch.stroke(particle.color[0], particle.color[1], particle.color[2], 255 * fadeRatio);
      particle.display();
    }
  };

  sketch.windowResized = () => {
    const size = frameSize();
    if (size.width === sketch.width && size.height === sketch.height) {
      return;
    }

    const previousFrame = sketch.get();
    const scaleX = size.width / sketch.width;
    const scaleY = size.height / sketch.height;

    particles.forEach((particle) => particle.scalePosition(scaleX, scaleY));
    sketch.resizeCanvas(size.width, size.height, true);
    sketch.background(backgroundColor);
    sketch.image(previousFrame, 0, 0, size.width, size.height);
    sketch.strokeCap(sketch.ROUND);
    sketch.strokeWeight(options.particleSize);
    sketch.noFill();
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
