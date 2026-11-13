# PowerPoint Particle Background

A Microsoft PowerPoint content add-in that renders a continuous p5.js particle animation inside a slide. The animation has no visible controls and resizes with the add-in frame.

## Start in PowerPoint

From PowerShell in this folder:

```powershell
.\start-local.ps1
```

The script starts the HTTPS development server and sideloads the add-in into desktop PowerPoint. If the content does not appear automatically, open **Home > Add-ins** and choose **Particle Background**.

On a fresh checkout, the script installs the npm dependencies before starting PowerPoint. Node.js 22.15 or newer is required.

To stop the server and unregister the development add-in:

```powershell
.\stop-local.ps1
```

## Development commands

```powershell
$env:NODE_OPTIONS = "--use-system-ca"
npm install
npm test
npm run lint
npm run validate
npm run build
```

`NODE_OPTIONS=--use-system-ca` lets Node.js use the trusted Windows certificate store on managed corporate networks.

## Share with other people

Localhost is only for development. To distribute the animation:

1. Host the generated web files from `dist` on an approved HTTPS site.
2. Set `ADDIN_PROD_URL` to the final HTTPS directory, including the trailing slash.
3. Run `npm run package:addin`.
4. Ask the Microsoft 365 administrator to deploy `dist/particle-background-app-package.zip` through Integrated apps.
5. Share the PowerPoint presentation after the add-in is available to its recipients.

The presentation stores a reference to the add-in. It does not embed the HTML and JavaScript inside the `.pptx` file.

Example production build:

```powershell
$env:ADDIN_PROD_URL = "https://intranet.example.com/powerpoint-particle-addin/"
npm run package:addin
```

Before production deployment, replace the placeholder developer website, privacy, and terms URLs in `manifest.json` with approved corporate pages.

## Publish on GitHub Pages

This repository publishes the production build to GitHub Pages whenever `main` is updated. The workflow builds with this production URL:

```text
https://jl8696.github.io/perlin/
```

To enable the first deployment:

1. Open the repository on GitHub.
2. Go to **Settings > Pages**.
3. Under **Build and deployment**, choose **GitHub Actions** as the source.
4. Open the **Actions** tab and run **Deploy GitHub Pages** if a deployment is not already running.
5. After the workflow succeeds, open `https://jl8696.github.io/perlin/content.html`.

Every later push to `main` republishes the site automatically.

To create the PowerPoint add-in package against the published site from PowerShell:

```powershell
$env:ADDIN_PROD_URL = "https://jl8696.github.io/perlin/"
npm run package:addin
```

The generated package is `dist/particle-background-app-package.zip`. GitHub Pages hosts the web app, but the add-in still needs to be sideloaded for development or deployed by a Microsoft 365 administrator before other people can insert it in PowerPoint.

## Source

This project is based on Microsoft's [PowerPoint content add-in Hello World sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-content-hello-world).
