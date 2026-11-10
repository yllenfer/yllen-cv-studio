#  📃CV Studio

A local React workspace for opening and editing CVs with Apryse WebViewer 12.2.1. The interface uses a classic 90s desktop publishing look, with clear **Open your CV** and **Download edited PDF** actions. Some of the documentation was hand-written by moi so I can avoid falling in the so called *"cognitive surrender"*

Feel free to fork it or contribute if you think something's missing. The live demo uses my personal demo key.

## Run locally

Requires Node.js 20 or newer.

```sh
npm install
npm run dev
```

Open the local URL printed by Vite (normally `http://localhost:5173`). Choose a PDF CV for the most reliable editing experience. Use **Edit text**, then double-click the name or paragraph you want to change. Click **Done editing** when finished; this switches back to the hand/pan tool if Esc does not exit the mode. Use **Download edited PDF** to save a copy; the original file stays unchanged.

DOC/DOCX files are handed to Apryse WebViewer's Office Editor. Editing and exporting Office documents depends on the active Apryse license and SDK configuration. The app exports a PDF copy; it does not provide an independent Microsoft Word editor or promise DOCX round-trip export. For scanned PDFs or text stored as an image/path, direct text editing may require OCR or editing the original DOCX instead.

## License key

The app reads `VITE_DEMO_KEY` from the environment. Keep keys in a local `.env` file, which is ignored by Git. Do not commit any `.env` file.

This is a browser app: a `VITE_` environment variable is bundled into the JavaScript sent to each visitor. Never set a private/shared key in Netlify for a public site. Apryse confirms that a client-side WebViewer key cannot be hidden from visitors. Without a key, WebViewer may run in its limited trial mode; continued production use requires a license appropriate for the deployed site.

## Deploy to Netlify

The repository includes `netlify.toml` with the Vite build command (`npm run build`), publish folder (`dist`), and Node 20. After pushing the repository to GitHub, import it from Netlify. Netlify installs dependencies and runs the postinstall script, which copies WebViewer assets into the build. The large generated `public/lib` directory is intentionally ignored by Git.

## Checks

```sh
npm run lint
npm run build
```

## Live Project
[https://apryser.netlify.app/](PDF editor live page)