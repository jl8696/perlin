import p5 from "p5";

/* global Office, console, document */

const options = {
  background: "#0a0a0a",
  colors: ["#ffffff", "#0799f2", "#45217c"],
  lifeSeconds: 10,
  particleCount: 400,
  particleSize: 2,
  noiseScale: 800,
};

function createParticleSketch(sketch) {
  const particles = [];
  let backgroundColor;

  class Particle {
    constructor() {
      this.velocity = sketch.createVector(0, 0);
      this.position = sketch.createVector(0, 0);
      this.life = 0;
      this.respawn(true);
    }

    move(iterations) {
      this.life -= sketch.deltaTime / 1000;
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

    display(radius) {
      sketch.ellipse(this.position.x, this.position.y, radius, radius);
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
    sketch.noStroke();

    backgroundColor = sketch.color(options.background);
    sketch.background(backgroundColor);

    for (let index = 0; index < options.particleCount; index += 1) {
      particles.push(new Particle());
    }
  };

  sketch.draw = () => {
    for (let index = 0; index < particles.length; index += 1) {
      const particle = particles[index];
      const iterations = sketch.map(index, 0, particles.length, 5, 1);

      particle.move(iterations);
      particle.checkEdge();

      let fadeRatio = sketch.min((particle.life * 5) / options.lifeSeconds, 1);
      fadeRatio = sketch.min(
        ((options.lifeSeconds - particle.life) * 5) / options.lifeSeconds,
        fadeRatio
      );

      const particleColor = sketch.color(options.colors[index % options.colors.length]);
      sketch.fill(
        sketch.red(particleColor),
        sketch.green(particleColor),
        sketch.blue(particleColor),
        255 * fadeRatio
      );
      particle.display(options.particleSize);
    }
  };

  sketch.windowResized = () => {
    const size = frameSize();
    sketch.resizeCanvas(size.width, size.height);
    backgroundColor = sketch.color(options.background);
    sketch.background(backgroundColor);
    particles.forEach((particle) => particle.respawn(true));
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
