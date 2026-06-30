# AURIC API — Render Deployment

This folder contains everything needed to deploy the AURIC Navigation Bar API to Render.com for free.

## Files

| File | Purpose |
|---|---|
| `app.py` | The Flask application |
| `requirements.txt` | Python dependencies (Flask, gunicorn, openpyxl) |
| `render.yaml` | Render infrastructure config (auto-detected) |
| `Procfile` | Process definition (backup config) |
| `runtime.txt` | Python version lock |
| `Navigation_Bar_Final.xlsx` | Source data |

## Quick deploy steps

1. Push this folder to a new public GitHub repo (e.g., `auric-api`)
2. Sign up at https://render.com (free, GitHub login works)
3. **New +** → **Web Service** → connect your GitHub repo
4. Render auto-detects `render.yaml`. Click **Apply** / **Create Web Service**
5. Wait 2–5 minutes for build
6. Your API is live at: `https://auric-api-XXXX.onrender.com`

See `DEPLOYMENT_GUIDE.docx` for detailed step-by-step instructions with screenshots placeholders.

## Test locally first (optional)

```bash
pip install -r requirements.txt
python app.py
# Open http://localhost:5000
```

## After deployment

Test the live URL:
```bash
curl https://YOUR-APP.onrender.com/api/health
curl https://YOUR-APP.onrender.com/api/plot
```

## Free-tier notes

Render's free tier spins down the service after 15 minutes of inactivity. The first request after a spin-down takes ~30 seconds to wake up. Subsequent requests are instant. For always-on, upgrade to the $7/month Starter plan.
