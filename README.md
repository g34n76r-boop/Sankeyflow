# SankeyFlow

An Excel add-in that draws Sankey charts from your data.

## Installing (Windows desktop Excel)

1. Download [SankeyFlow-Setup.zip](https://g34n76r-boop.github.io/SankeyFlow/SankeyFlow-Setup.zip) and unzip it.
2. Close Excel, then double-click **Install SankeyFlow.cmd**. If Windows shows "Windows protected your PC", choose **More info → Run anyway**.
3. Open Excel. The **SankeyFlow** button is on the Home tab.

Updates arrive automatically. To remove it, run **Uninstall SankeyFlow.cmd**.

On Mac or Excel on the web, use **Add-ins → More Add-ins → My Add-ins → Upload My Add-in** with the [published manifest](https://g34n76r-boop.github.io/SankeyFlow/manifest.xml).

## Development

```
npm install
npm start        # dev server + Excel with the add-in sideloaded from localhost
npm run build    # production build in dist/
npm run typecheck
npm run lint
```

Pushes to `main` deploy to GitHub Pages (`.github/workflows/pages.yml`).
