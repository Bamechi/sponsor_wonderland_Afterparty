# Down the Rabbit Hole — sponsor site

The official after party of BEYOND Wonderland at AWKN Ranch. Single-page, mobile-first partner invitation.

- `index.html` — the whole site (HTML, CSS and JS in one file)
- `og-image.jpg` — the preview card that appears when the link is texted or posted
- `favicon.ico`, `icon-*.png`, `site.webmanifest` — icons
- `vercel.json` — caching rules
- `Code.gs` — the Google Apps Script that logs partner submissions to the tracker sheet

## Deploy (GitHub → Vercel)
1. Upload every file in this folder to the repo root: https://github.com/Bamechi/sponsor_wonderland_Afterparty
2. vercel.com → Add New → Project → Import this repo.
3. Framework Preset: **Other**. Build command: empty. Output directory: empty. Root: `./`
4. Deploy. The live URL is `https://<project-name>.vercel.app`.
5. If the project name is NOT `sponsor-wonderland-afterparty`, open `index.html` and replace that domain in the three `og:image` / `og:url` / `twitter:image` tags, then commit again.

## Form tracking
Submissions post to the Apps Script web app and land in the Sponsors sheet.
Endpoint lives in one place in `index.html`: the `const ENDPOINT` line near the bottom.
