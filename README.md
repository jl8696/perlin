# PowerPoint Particle Background

A Microsoft PowerPoint content add-in that renders a continuous p5.js particle animation inside a slide. The animation has no visible controls and resizes with the add-in frame.

## Start in PowerPoint

From PowerShell in this folder:

```powershell
.\start-local.ps1
```

The script starts the HTTPS development server and sideloads the add-in into desktop PowerPoint. If the content does not appear automatically, open **Home > Add-ins** and choose **Particle Background**.

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
3. Run `npm run build`.
4. Ask the Microsoft 365 administrator to deploy `dist/manifest.json` through Integrated apps.
5. Share the PowerPoint presentation after the add-in is available to its recipients.

The presentation stores a reference to the add-in. It does not embed the HTML and JavaScript inside the `.pptx` file.

Example production build:

```powershell
$env:ADDIN_PROD_URL = "https://intranet.example.com/powerpoint-particle-addin/"
npm run build
```

## Source

This project is based on Microsoft's [PowerPoint content add-in Hello World sample](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/powerpoint-content-hello-world).
