import p5 from "p5";

/* global Office, console, document */

const options = {
  Background: "#0a0a0a",
  Color1: "#ffffff",
  Color2: "#0799f2",
  Color3: "#45217c",
  Length: 10,
  Nums: 400,
  Size: 2,
  noiseScale: 800,
  ColorMode: "Normal",
};

function createParticleSketch(sketch) {
  const particles = [];
  let maxLife;
  let backgroundColor;

  function frameSize() {
    const host = document.getElementById("particle-background");
    return {
      host,
      width: Math.max(host.clientWidth, 1),
      height: Math.max(host.clientHeight, 1),
    };
  }

  function Particle() {
    this.vel = sketch.createVector(0, 0);
    this.pos = sketch.createVector(
      sketch.random(-50, sketch.width + 50),
      sketch.random(-50, sketch.height + 50)
    );
    this.life = sketch.random(0, maxLife);

    this.move = function (iterations) {
      if ((this.life -= 0.01666) < 0) {
        this.respawn();
      }

      while (iterations > 0) {
        const angle =
          sketch.noise(this.pos.x / options.noiseScale, this.pos.y / options.noiseScale) *
          sketch.TWO_PI *
          options.noiseScale;
        this.vel.x = sketch.cos(angle);
        this.vel.y = sketch.sin(angle);
        this.vel.mult(0.2);
        this.pos.add(this.vel);
        iterations -= 1;
      }
    };

    this.checkEdge = function () {
      if (
        this.pos.x > sketch.width ||
        this.pos.x < 0 ||
        this.pos.y > sketch.height ||
        this.pos.y < 0
      ) {
        this.respawn();
      }
    };

    this.respawn = function () {
      this.pos.x = sketch.random(-50, sketch.width + 50);
      this.pos.y = sketch.random(-50, sketch.height + 50);
      this.life = maxLife;
    };

    this.display = function (radius) {
      sketch.ellipse(this.pos.x, this.pos.y, radius, radius);
    };
  }

  sketch.setup = () => {
    const size = frameSize();
    backgroundColor = sketch.color(options.Background);
    const canvas = sketch.createCanvas(size.width, size.height);
    canvas.parent(size.host);
    sketch.background(options.Background);

    for (let index = 0; index < 2500; index += 1) {
      particles[index] = new Particle();
    }
  };

  sketch.draw = () => {
    sketch.noStroke();
    sketch.smooth();

    maxLife = options.Length;
    for (let index = 1; index < options.Nums; index += 1) {
      const iterations = sketch.map(index, 0, options.Nums, 5, 1);
      const radius = options.Size;

      particles[index].move(iterations);
      particles[index].checkEdge();

      const alpha = 255;
      let particleColor;
      let fadeRatio = sketch.min((particles[index].life * 5) / maxLife, 1);
      fadeRatio = sketch.min(((maxLife - particles[index].life) * 5) / maxLife, fadeRatio);
      const lifeRatioGrayscale = sketch.min(
        255,
        (255 * particles[index].life) / maxLife + sketch.red(backgroundColor)
      );
      void lifeRatioGrayscale;

      if (options.ColorMode === "Normal") {
        if (index % 3 === 0) particleColor = options.Color1;
        if (index % 3 === 1) particleColor = options.Color2;
        if (index % 3 === 2) particleColor = options.Color3;
      }

      if (options.ColorMode === "Linera Gradient") {
        const percent1 = sketch.norm(particles[index].pos.x, 0, sketch.width / 2);
        const percent2 = sketch.norm(particles[index].pos.x, sketch.width / 2, sketch.width);
        const from = sketch.color(options.Color1);
        const middle = sketch.color(options.Color2);
        const to = sketch.color(options.Color3);
        const between1 = sketch.lerpColor(from, middle, percent1);
        const between2 = sketch.lerpColor(middle, to, percent2);
        if (particles[index].pos.x > 0 && particles[index].pos.x < sketch.width / 2) {
          particleColor = between1;
        } else {
          particleColor = between2;
        }
      }

      if (options.ColorMode === "Radial Gradient") {
        const distance = sketch.dist(
          particles[index].pos.x,
          particles[index].pos.y,
          sketch.width / 2,
          sketch.height / 2
        );
        const percent1 = sketch.norm(distance, 0, 400);
        const percent2 = sketch.norm(distance, 400, sketch.width / 2);
        const from = sketch.color(options.Color1);
        const middle = sketch.color(options.Color2);
        const to = sketch.color(options.Color3);
        const between1 = sketch.lerpColor(from, middle, percent1);
        const between2 = sketch.lerpColor(middle, to, percent2);
        if (distance < 400) {
          particleColor = between1;
        } else {
          particleColor = between2;
        }
      }

      if (options.ColorMode === "Splice") {
        if (
          particles[index].pos.x >= sketch.width / 3 &&
          particles[index].pos.x <= (sketch.width / 3) * 2
        ) {
          if (index % 3 === 0) particleColor = options.Color1;
          if (index % 3 === 1) particleColor = options.Color2;
          if (index % 3 === 2) particleColor = options.Color3;
        } else if (particles[index].pos.x < sketch.width / 3) {
          if (index % 3 === 0) particleColor = 20;
          if (index % 3 === 1) particleColor = 100;
          if (index % 3 === 2) particleColor = 220;
        } else if (particles[index].pos.x > (sketch.width / 3) * 2) {
          if (index % 3 === 0) {
            particleColor = sketch.color(
              255 - sketch.red(options.Color1),
              255 - sketch.green(options.Color1),
              255 - sketch.blue(options.Color1)
            );
          }
          if (index % 3 === 1) {
            particleColor = sketch.color(
              255 - sketch.red(options.Color2),
              255 - sketch.green(options.Color2),
              255 - sketch.blue(options.Color2)
            );
          }
          if (index % 3 === 2) {
            particleColor = sketch.color(
              255 - sketch.red(options.Color3),
              255 - sketch.green(options.Color3),
              255 - sketch.blue(options.Color3)
            );
          }
        }
      }

      sketch.fill(
        sketch.red(particleColor),
        sketch.green(particleColor),
        sketch.blue(particleColor),
        alpha * fadeRatio
      );
      particles[index].display(radius);
    }
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
