# BAHL Validation System — Run Instructions

## Backend

1. Create and activate a Python virtual environment:
   - PowerShell:
     ```powershell
     python -m venv .venv
     .venv\Scripts\Activate.ps1
     ```
   - Command Prompt:
     ```cmd
     python -m venv .venv
     .venv\Scripts\activate.bat
     ```

2. Install backend dependencies:
   ```powershell
   pip install -r backend/requirements.txt
   ```

3. Start the backend from the project root:
   ```powershell
   python app.py
   ```

4. Backend URL:
   ```
   http://127.0.0.1:5000
   ```

## Frontend

1. Change to the frontend directory:
   ```powershell
   cd frontend
   ```

2. Install frontend dependencies:
   ```powershell
   npm install
   ```

3. Start the Vite development server:
   ```powershell
   $env:VITE_API_URL='http://127.0.0.1:5000'
   npm run dev -- --host 127.0.0.1 --port 5173
   ```

4. Frontend URL:
   ```
   http://127.0.0.1:5173
   ```

> Note: The frontend uses `frontend/.env.local` to set `VITE_API_URL` in development. Ensure it is set to `http://127.0.0.1:5000` and restart the frontend server after any change.

## Notes

- The backend main module is `backend/app.py`.
- The frontend configuration is in `frontend/package.json`.
- A VS Code task already exists for running the frontend dev server as `Run frontend dev server`.
