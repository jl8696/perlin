# PowerPoint Particle Add-in Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build and launch a PowerPoint content add-in that displays the supplied p5.js particle animation as a live slide background.

**Architecture:** Start from Microsoft's PowerPoint content add-in sample and keep its manifest, webpack development server, and sideload scripts. Replace the sample UI with a full-frame p5.js canvas and bundle p5.js as an npm dependency so the runtime is self-contained apart from Office.js.

**Tech Stack:** HTML, CSS, JavaScript, p5.js, Office.js, Webpack, Microsoft 365 unified manifest

**Spec:** `docs/superpowers/specs/2026-11-13-powerpoint-particle-addin-design.md`

## Global Constraints

- Use PowerShell on Windows 11.
- Use HTTPS for the add-in source location.
- Preserve the original particle colors and behavior.
- Do not add visible controls or sample content.

## Review Focus

- The canvas fills the content frame without scrollbars or margins.
- Resizing the add-in resizes and repopulates the canvas.
- The project builds without depending on the public p5.js CDN.
- The unified manifest validates and targets PowerPoint.
- Sideload failure reports any Microsoft 365 policy restriction clearly.

---

### Task 1: Scaffold the add-in

**Files:**
- Create: `powerpoint-particle-addin/` from the official sample
- Create: `powerpoint-particle-addin/docs/superpowers/specs/2026-11-13-powerpoint-particle-addin-design.md`
- Create: `powerpoint-particle-addin/docs/superpowers/plans/2026-11-13-powerpoint-particle-addin.md`

**Interfaces:**
- Consumes: Microsoft's `powerpoint-content-hello-world` sample
- Produces: installable npm project with manifest and webpack configuration

- [ ] Clone the Office add-in samples repository with a sparse checkout of the PowerPoint content sample.
- [ ] Copy the sample into the dedicated `powerpoint-particle-addin` project folder.
- [ ] Install dependencies and add p5.js as a local runtime dependency.

### Task 2: Replace the sample UI with the particle background

**Files:**
- Modify: `powerpoint-particle-addin/content.html`
- Modify: `powerpoint-particle-addin/content.css`
- Modify: `powerpoint-particle-addin/content.js`
- Modify: `powerpoint-particle-addin/manifest.json`
- Modify: `powerpoint-particle-addin/package.json`

**Interfaces:**
- Consumes: local p5.js package and Office.js host initialization
- Produces: full-frame animated content page at `https://localhost:3000/content.html`

- [ ] Replace the sample markup with an empty accessible animation host.
- [ ] Add full-frame canvas styling with no visible UI or overflow.
- [ ] Move the particle implementation into `content.js` and initialize after Office is ready.
- [ ] Rename the package and add-in display names for the particle background.

### Task 3: Verify and launch

**Files:**
- Verify: `powerpoint-particle-addin/dist/*`

**Interfaces:**
- Consumes: completed add-in project
- Produces: validated build, browser rendering evidence, and PowerPoint sideload result

- [ ] Run lint, manifest validation, and production build.
- [ ] Start the HTTPS development server and verify `content.html` responds.
- [ ] Inspect the rendered animation at a 16:9 viewport.
- [ ] Confirm the canvas contains non-background pixels and fills the viewport.
- [ ] Attempt the standard PowerPoint sideload and report any policy or sign-in block.

