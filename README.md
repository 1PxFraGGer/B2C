# University Workspace — React frontend

This package contains only the React/Vite frontend. It has no backend credentials,
password hashes, session secrets, Apps Script files, or demo data.

## Push to GitHub

1. Extract this ZIP into a folder.
2. Create your GitHub repository.
3. Upload the extracted contents so package.json, index.html, src/, and vercel.json
   are at the repository root. Do not upload the ZIP itself as the project.
4. Keep .gitignore. Do not upload node_modules or dist.

## Deploy on Vercel

1. Import the GitHub repository into Vercel.
2. Select the Vite framework and the repository root as Root Directory.
3. Build Command: npm run build
4. Output Directory: dist
5. Deploy. No environment variables are required for this frontend-only stage.
6. Send the resulting HTTPS Vercel link back in this chat.

## What works at this stage

The application builds and displays the responsive login screen. Dashboard, team
workspaces, content forms, per-item image editors, and reports are included in the
source. They require authenticated backend data before they can be used. There is
no frontend password validation or fabricated assignment data.

The message "Backend connection pending" is expected until the backend is
connected. Login and uploads are not live in this frontend-only deployment.

## Connect the backend later

The frontend calls /api/login, /api/session, /api/logout, /api/data,
/api/saveItem, /api/saveContent, /api/image, and /api/reports on the same origin.
We will add the already prepared Vercel API and Google Apps Script integration
after you provide the deployed frontend URL. The API routes must be added to
vercel.json before the SPA fallback when integrating that backend.

## Run locally

Use Node.js 22.12 or later:

```sh
npm install
npm run dev
```

```sh
npm run build
npm run preview
```
