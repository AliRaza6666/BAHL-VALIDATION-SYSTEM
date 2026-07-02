# Deploy Story

This project is deployed as a single Vercel app from the repository root.

## Repository Layout

- Frontend: [frontend/](frontend)
- Backend: [backend/api/index.py](backend/api/index.py)
- Root Vercel config: [vercel.json](vercel.json)

## Deployment Model

The app uses one Vercel project, not separate frontend and backend hosting.

- The frontend is built with Vite from `frontend/package.json`
- The backend is a Python serverless function exposed through `backend/api/index.py`
- The root `vercel.json` routes `/api/*` to the backend
- All other routes go to the frontend

## Important Environment Variable

Set:

- `VITE_API_URL=/api`

This makes the frontend call the same deployed domain for backend requests.

## What Was Fixed

1. The Vercel config version was corrected to `2`
2. The backend import-time directory creation was removed so the serverless function can initialize on Vercel
3. The live `/api/health` endpoint now returns `{"status":"ok"}`
4. The frontend now uses `/api` for upload and validation requests

## Current Vercel Config

The root `vercel.json` is:

```json
{
  "version": 2,
  "builds": [
    {
      "src": "frontend/package.json",
      "use": "@vercel/static-build",
      "config": { "distDir": "dist" }
    },
    {
      "src": "backend/api/index.py",
      "use": "@vercel/python"
    }
  ],
  "routes": [
    {
      "src": "/api/(.*)",
      "dest": "/backend/api/index.py"
    },
    {
      "src": "/(.*)",
      "dest": "/frontend/$1"
    }
  ]
}
```

## Deployment Steps

1. Push changes to `main`
2. Let Vercel auto-deploy, or run `vercel --prod`
3. Open the production site
4. Test `/api/health`
5. Upload a small Excel file and validate it

## Notes

- The app may still fail on very large Excel uploads because Vercel serverless functions have limits
- Small files should work reliably
- The `/api/health` endpoint is the best quick check that the backend is alive
