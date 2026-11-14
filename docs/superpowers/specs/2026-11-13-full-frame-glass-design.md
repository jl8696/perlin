# Full-Frame Glass Shader Design

## Intent

Make the accumulated particle animation read as one image processed by PowerPoint's Glass artistic effect, rather than as individually frosted particle trails.

## Goals

- Restore the supplied original particle settings: 400 particles, 0.2 movement step, 2px point rendering, and the original Perlin iteration mapping.
- Apply 65% opacity once to the complete rendered canvas.
- Use a 3-second trail half-life so old paths disappear substantially faster.
- Run one GPU post-process over the complete offscreen trail frame.
- Use multi-pixel refraction, neighborhood sampling, faceting, and edge lighting so the effect remains visible when PowerPoint scales the add-in.
- Keep black regions dark while allowing bright trail regions to refract and spread into nearby pixels.
- Maintain approximately 60 FPS at 1268 by 714 pixels.

## Non-goals

- Reproduce Microsoft's proprietary filter implementation byte-for-byte.
- Apply a separate filter to each particle or trail segment.
- Change the Perlin movement calculations.

## Approach

Retain the existing offscreen `trailLayer` as the source image. Replace the trail-masked fragment treatment with a full-frame shader that computes a coarse glass field, displaces the complete source texture, samples a multi-pixel neighborhood, and derives facets and silver highlights from the refracted image. The particle mask may gate final visibility from sampled image energy, but it must not gate refraction, neighborhood aggregation, or glass-field generation.

## Constraints

- Windows 11 and PowerShell.
- p5.js and WebGL remain the rendering stack.
- Production bundles remain cache-busted for PowerPoint.
- The public GitHub Pages deployment must be verified after the change.
