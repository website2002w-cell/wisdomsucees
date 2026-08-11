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

Notes:
- Ensure sensitive environment variables (API keys, DB URIs) are set as platform secrets. Do not commit them.
- For automatic migrations or production DB, update `server.ts` to read database connection info from `process.env`.
