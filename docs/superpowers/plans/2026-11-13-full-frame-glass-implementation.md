# Full-Frame Glass Shader Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace particle-masked frosting with one full-frame glass post-process over the accumulated trail image.

**Architecture:** `trailLayer` remains the animated source texture. The WebGL fragment shader builds a coarse full-frame distortion field, refracts the source, aggregates neighboring samples, then shades the refracted frame with faceting and edge highlights.

**Tech Stack:** JavaScript, p5.js, WebGL/GLSL, Node test runner, Webpack

**Spec:** `docs/superpowers/specs/2026-11-13-full-frame-glass-design.md`

## Global Constraints

- Restore the original 400-particle, 0.2-step, 2px point-rendering behavior.
- Apply 65% opacity to the complete canvas and use a 3-second trail half-life.
- Do not gate refraction, neighborhood aggregation, or glass-field generation with a particle mask.
- Maintain approximately 60 FPS at 1268 by 714 pixels.
- Preserve cache-busted production bundle URLs.

## Review Focus

- PowerPoint downscaling must not erase the glass facets.
- Empty black regions must remain predominantly dark after refraction.
- Trail history must continue fading instead of accumulating indefinitely.
- Resize must preserve the accumulated frame and particle state.
- The shader must compile without WebGL runtime errors.

---

### Task 1: Full-Frame Glass Pass

**Files:**
- Modify: `tests/addin.test.cjs`
- Modify: `content.js`

**Interfaces:**
- Consumes: `trailLayer` as a complete-frame sampler2D texture.
- Produces: one shaded full-frame WebGL rectangle through `glassShader`.

- [ ] **Step 1: Write the failing regression test**

Require multi-pixel neighborhood sampling and a coarse glass field, and reject multiplying refraction or facet generation by `trailMask`.

- [ ] **Step 2: Run `npm test` and verify the new assertion fails**

- [ ] **Step 3: Replace the masked shader with full-frame displacement, neighborhood aggregation, and facet lighting**

Leave particle movement and trail-fade code unchanged.

- [ ] **Step 4: Run `npm test` and verify the complete suite passes**

### Task 2: Visual, Performance, and Deployment Verification

**Files:**
- Modify: generated `dist/*`

**Interfaces:**
- Consumes: production Webpack build.
- Produces: cache-busted GitHub Pages deployment.

- [ ] **Step 1: Render the animation at 1268 by 714 and inspect it against the supplied references**

- [ ] **Step 2: Measure animation frame timing and confirm no browser errors**

- [ ] **Step 3: Run tests, lint, diff checks, and the production build**

- [ ] **Step 4: Commit, push, wait for GitHub Pages, and verify the deployed bundle contains the full-frame shader**
