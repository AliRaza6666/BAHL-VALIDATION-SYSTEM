from backend.app import app

# Vercel serverless entrypoint for backend API
# The Flask app is already defined and configured in backend/app.py

# Expose `app` as the callable required by Vercel Python runtime

__all__ = ["app"]
