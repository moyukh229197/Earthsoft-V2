"use strict";

function createBackgroundBeams() {
  const container = document.createElement('div');
  container.className = 'background-beams-container';

  // Custom styling that mimics the Animated Grid Pattern demo mask
  container.style.position = 'absolute';
  container.style.left = '0';
  container.style.right = '0';
  container.style.top = '-30%';
  container.style.bottom = '-30%';
  container.style.height = '160%'; // adjusted for better layout
  container.style.pointerEvents = 'none';
  container.style.zIndex = '0';
  container.style.transform = 'skewY(12deg)';

  // We apply the demo mask radial gradient
  container.style.maskImage = "radial-gradient(ellipse at center, white 10%, transparent 60%)";
  container.style.webkitMaskImage = "radial-gradient(ellipse at center, white 10%, transparent 70%)";

  const width = 45;
  const height = 45;
  const numSquares = 40;
  const maxOpacity = 0.15;
  const duration = 2.5;
  const repeatDelay = 0.5;
  const x = -1;
  const y = -1;
  const strokeDasharray = 0;

  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("aria-hidden", "true");
  svg.style.position = "absolute";
  svg.style.inset = "0";
  svg.style.height = "100%";
  svg.style.width = "100%";
  // Subtle grey stroke for grid lines
  svg.style.fill = "rgba(156, 163, 175, 0.4)";
  svg.style.stroke = "rgba(156, 163, 175, 0.4)";

  const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");
  const pattern = document.createElementNS("http://www.w3.org/2000/svg", "pattern");
  const patternId = "grid-pattern-id";
  pattern.setAttribute("id", patternId);
  pattern.setAttribute("width", width.toString());
  pattern.setAttribute("height", height.toString());
  pattern.setAttribute("patternUnits", "userSpaceOnUse");
  pattern.setAttribute("x", x.toString());
  pattern.setAttribute("y", y.toString());

  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", `M.5 ${height}V.5H${width}`);
  path.setAttribute("fill", "none");
  path.setAttribute("stroke-dasharray", strokeDasharray.toString());
  pattern.appendChild(path);
  defs.appendChild(pattern);
  svg.appendChild(defs);

  const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  rect.setAttribute("width", "100%");
  rect.setAttribute("height", "100%");
  rect.setAttribute("fill", `url(#${patternId})`);
  svg.appendChild(rect);

  const innerSvg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  innerSvg.setAttribute("x", x.toString());
  innerSvg.setAttribute("y", y.toString());
  innerSvg.style.overflow = "visible";

  let squares = [];
  let w = window.innerWidth;
  let h = window.innerHeight * 2; // Approximating due to skew and height multiplier

  function getPos() {
    return [
      Math.floor((Math.random() * w) / width),
      Math.floor((Math.random() * h) / height),
    ];
  }

  // Create the animated squares
  for (let i = 0; i < numSquares; i++) {
    const sq = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    sq.setAttribute("width", (width - 1).toString());
    sq.setAttribute("height", (height - 1).toString());

    const pos = getPos();
    sq.setAttribute("x", (pos[0] * width + 1).toString());
    sq.setAttribute("y", (pos[1] * height + 1).toString());
    sq.setAttribute("fill", "currentColor");
    // In dark themes it will be white, in light maybe dark depending on body color
    // We force white here for the cool "tech" effect against the dark background
    sq.style.color = "rgba(255, 255, 255, 1)";
    sq.setAttribute("stroke-width", "0");
    sq.style.opacity = "0";

    sq.style.transition = `opacity ${duration}s ease-in-out`;

    innerSvg.appendChild(sq);
    squares.push({ el: sq, index: i });
  }

  svg.appendChild(innerSvg);
  container.appendChild(svg);

  function animateSquares() {
    squares.forEach((sq) => {
      const initialDelay = sq.index * 0.1 * 1000; // staggered start
      setTimeout(() => {
        const loop = () => {
          // Fade in
          sq.el.style.opacity = maxOpacity.toString();

          setTimeout(() => {
            // Fade out
            sq.el.style.opacity = "0";

            setTimeout(() => {
              // Re-position when hidden
              const pos = getPos();
              sq.el.setAttribute("x", (pos[0] * width + 1).toString());
              sq.el.setAttribute("y", (pos[1] * height + 1).toString());

              // Repeat
              loop();
            }, duration * 1000 + repeatDelay * 1000);

          }, duration * 1000); // stay visible/animating for `duration`
        };
        loop();
      }, initialDelay);
    });
  }

  requestAnimationFrame(() => {
    animateSquares();
  });

  window.addEventListener('resize', () => {
    w = window.innerWidth;
    h = window.innerHeight * 2;
  });

  return container;
}

document.addEventListener("DOMContentLoaded", () => {
  const existing = document.querySelector('.background-beams-container');
  if (existing) {
    existing.innerHTML = ''; // Clear if it existed
  }
  const beams = createBackgroundBeams();
  // We prepend to the body
  document.body.prepend(beams);
});
