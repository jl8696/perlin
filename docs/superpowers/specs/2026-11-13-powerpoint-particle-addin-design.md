# PowerPoint Particle Add-in Design

## Intent

Create a reusable PowerPoint content add-in that renders the supplied p5.js particle animation as a live, full-frame slide background. The local project must be easy to start for development and ready to switch to an approved HTTPS host for distribution.

## Goals

- Base the project on Microsoft's PowerPoint content add-in sample.
- Render only the particle animation, with no controls, credits, or sample UI.
- Fill the add-in frame at every size and react to resizing.
- Support local HTTPS development through the sample's existing npm workflow.
- Keep the manifest and project understandable for later Microsoft 365 deployment.

## Non-goals

- Create or centrally deploy the production Microsoft 365 app package.
- Choose the final corporate HTTPS host.
- Add slide text or other PowerPoint content.

## Approach

Use Microsoft's sample scaffold instead of building an Office add-in from scratch. Replace its sample content page and scripts with the supplied p5.js animation, retain Office.js initialization, and bundle p5.js locally so the animation does not depend on a public CDN at presentation time. Validate with the sample's manifest validator, production build, a browser rendering check, and a PowerPoint sideload attempt.

## Constraints

- Windows 11 and PowerShell.
- Desktop Microsoft 365 PowerPoint.
- The add-in uses HTTPS, including localhost development.
- Organization policy may block sideloading or require administrator approval.

