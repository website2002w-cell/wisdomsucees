Deployment Options
------------------

Local Docker (recommended):

1. Build and run with Docker Compose:

```bash
docker compose build
docker compose up -d
```

2. Open http://localhost:3000

Build image directly:

```bash
docker build -t wisdom-school:latest .
docker run -p 3000:3000 --env PORT=3000 wisdom-school:latest
```

CI / GitHub Actions:

- The workflow at `.github/workflows/ci-deploy.yml` builds the app and uploads the `dist/` artifact on pushes to `main`.
- If you add `DOCKERHUB_USERNAME` and `DOCKERHUB_TOKEN` secrets, the workflow will also push a Docker image.

Deploy to Render / Fly / Railway:

- Use a Docker deployment: point the service to this repo and the platform will build the Docker image using the `Dockerfile`.

Deploy to Netlify
-----------------

This repo also supports deploying the built static site to Netlify (Vite builds to `dist/`). Netlify will run `npm run build` and publish the `dist` folder.

1. In Netlify, click "New site from Git" → connect your Git provider and authorize the repo `website2002w-cell/wisdomsucees`.
2. Set build command: `npm run build`
3. Set publish directory: `dist`
4. (Optional) Add environment variables in Site settings → Build & deploy → Environment:
	- `NODE_ENV` = `production`
5. The included `public/_redirects` and `netlify.toml` ensure SPA routing works and Netlify publishes the `dist` folder.
6. Deploy the site and open the generated Netlify URL. If the build previously failed with "Failed to resolve /src/main.tsx", ensure the repo contains the updated `index.html` that references `./src/main.tsx` (this repo has that fix).

 - Use a Docker deployment: point the service to this repo and the platform will build the Docker image using the `Dockerfile`.

GitHub Actions automatic Fly deploy
---------------------------------

You can deploy automatically from GitHub using the workflow `.github/workflows/deploy-to-fly.yml` included in this repo.

1. Create a Fly app (via Fly UI) and note the app name.
2. In the GitHub repo settings → Secrets → Actions, add two secrets:
	- `FLY_API_TOKEN` — create via `flyctl auth token` (or from Fly dashboard) and paste the token.
	- `FLY_APP_NAME` — the Fly app name you created.
3. Push to `main` — the workflow will build and run `flyctl deploy` to your Fly app.

If deploy fails, open Actions → Deploy to Fly to view logs and paste errors here for debugging.

Notes:
- Ensure sensitive environment variables (API keys, DB URIs) are set as platform secrets. Do not commit them.
- For automatic migrations or production DB, update `server.ts` to read database connection info from `process.env`.
