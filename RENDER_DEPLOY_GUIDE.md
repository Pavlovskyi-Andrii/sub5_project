# Render Deployment Guide

## Problem Fixed

The deployment was failing with `gunicorn: command not found` because Render was auto-detecting the app instead of using the correct configuration.

## Solution - Two Deployment Methods

### Method 1: Deploy with render.yaml (Blueprint) - **RECOMMENDED**

1. **Go to Render Dashboard**: https://dashboard.render.com/
2. **Click "New +" → "Blueprint"**
3. **Connect your repository**
4. **Render will automatically detect `render.yaml`**
5. **Set environment variables** in the Render dashboard:
   - `GARMIN_EMAIL` - Your Garmin email
   - `GARMIN_PASSWORD` - Your Garmin password
   - `FLASK_SECRET_KEY` - Auto-generated (or set your own)
   - `SERVICE_ACCOUNT_JSON` - Your Google service account JSON (if using Google Sheets sync)
   - `GOOGLE_SHEET_URL` - Your Google Sheet URL (if using Google Sheets sync)

6. **Click "Apply"** - Render will use the configuration from `render.yaml`

### Method 2: Manual Web Service Deployment

If you've already created a Web Service manually:

1. **Go to your service settings**
2. **Update the following settings:**
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT`
   - **Python Version**: `3.12.0` (set in Environment → Add Environment Variable: `PYTHON_VERSION=3.12.0`)

3. **Set environment variables:**
   - `PORT` = `10000` (Render usually sets this automatically)
   - `GARMIN_EMAIL` = your Garmin email
   - `GARMIN_PASSWORD` = your Garmin password
   - `FLASK_SECRET_KEY` = random secret key (use a password generator)
   - Optional: `SERVICE_ACCOUNT_JSON`, `GOOGLE_SHEET_URL`, `DAYS_TO_SYNC`

4. **Save Changes** and trigger a manual deploy

## What Changed

### Fixed Files:

1. **render.yaml**:
   - Updated start command to: `gunicorn app:app --bind 0.0.0.0:$PORT`
   - Removed `preDeployCommand` that might cause issues
   - Added PORT environment variable

2. **app.py**:
   - Added automatic database initialization when the app starts
   - No need for separate database setup step

3. **Procfile** (already correct):
   - Command: `web: python -m gunicorn app:app --bind 0.0.0.0:$PORT`

## Verification

After deployment succeeds, you should see:
- ✅ Build successful
- ✅ Deploy successful
- ✅ Service is live at `https://your-service-name.onrender.com`

Visit your URL to see the dashboard!

## Common Issues

### Issue: "gunicorn: command not found"
**Solution**: Make sure the Build Command includes `pip install -r requirements.txt`

### Issue: "No module named 'app'"
**Solution**: Make sure Start Command is `gunicorn app:app` (not `gunicorn your_application.wsgi`)

### Issue: Database errors
**Solution**: The app now auto-initializes the database. If issues persist, check file write permissions.

### Issue: "Failed to connect to Garmin"
**Solution**: Verify `GARMIN_EMAIL` and `GARMIN_PASSWORD` environment variables are set correctly.

## Next Steps

1. **Commit these changes**:
   ```bash
   git add render.yaml app.py RENDER_DEPLOY_GUIDE.md
   git commit -m "Fix Render deployment configuration"
   git push
   ```

2. **Deploy using Method 1 (Blueprint)** for best results

3. **Test your deployment** by visiting the service URL

4. **Sync your Garmin data** by clicking the sync button in the web UI

## Support

If you continue to have issues:
- Check Render logs: Dashboard → Your Service → Logs
- Verify all environment variables are set
- Ensure `requirements.txt` includes all dependencies
