# Initiation BOM Generator — Deploy Guide (free hosting on Render)

A web page where the team uploads an ERP quotation CSV and downloads the
Initiation BOM Excel. Master data is built in.

## What you get
A public URL like `https://initiation-bom.onrender.com` that anyone can open.
(Free tier sleeps after 15 min idle; first visit after sleeping takes ~30s to wake.)

## One-time setup (about 10 minutes)

### Step 1 — Put these files in a GitHub repo
1. Create a free account at https://github.com if you don't have one.
2. Click "New repository", name it `initiation-bom`, keep it Private, Create.
3. Upload ALL these files (drag-and-drop on GitHub's "Add file → Upload files"):
   - app.py
   - initiation_bom.py
   - master_data.xlsx
   - bom_template.xlsx
   - requirements.txt
   - Procfile
   - render.yaml
   - .gitignore
4. Commit.

### Step 2 — Deploy on Render
1. Create a free account at https://render.com (sign in with GitHub — easiest).
2. Click "New +" → "Web Service".
3. Connect your `initiation-bom` GitHub repo.
4. Render auto-detects the settings from render.yaml. If asked, confirm:
   - Runtime: Python
   - Build command: pip install -r requirements.txt
   - Start command: gunicorn app:app
   - Plan: Free
5. Click "Create Web Service". Wait ~2-3 min for the first build.
6. When it says "Live", your URL appears at the top, e.g.
   https://initiation-bom.onrender.com  — share this with the team.

## Updating the master data later
When you add new items to the master:
1. On GitHub, open the repo → click master_data.xlsx → delete it.
2. Upload the new master_data.xlsx (same filename).
3. Render auto-redeploys in ~2 min. Done.

## Using a custom domain (erp.decorpot.com) — optional, later
Render supports custom domains even on paid plans. In the service's
"Settings → Custom Domain", add erp.decorpot.com, then add the CNAME record
Render gives you into Decorpot's DNS (whoever manages decorpot.com does this).
Free tier allows custom domains too, but the sleep behaviour remains; a paid
plan ($7/mo) keeps it always-on.

## Run locally to try first (optional)
    pip install -r requirements.txt
    python3 app.py
Then open http://localhost:8000
