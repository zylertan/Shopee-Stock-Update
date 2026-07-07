# Deploy to Streamlit Community Cloud (free, shareable URL)

You'll create a GitHub repo, upload 5 files, and connect it to Streamlit Cloud.
No coding needed — just uploads and clicks. ~10 minutes.

## Files to upload to GitHub (in this folder)
```
app.py
shopee_core.py
requirements.txt
README.md
.gitignore
.streamlit/config.toml
```
Do NOT upload your Masterlist / Shopee / registry Excel files — those are uploaded
inside the app each time you use it (they're processed in memory, never stored).

## Step 1 — Create a GitHub account & repository
1. Sign up at https://github.com (free).
2. Click the **+** (top-right) → **New repository**.
3. Name it e.g. `shopee-stock-update`. Set **Private** (recommended). Click **Create repository**.

## Step 2 — Upload the files
1. On the new repo page click **uploading an existing file** (or **Add file → Upload files**).
2. Drag in `app.py`, `shopee_core.py`, `requirements.txt`, `README.md`, `.gitignore`.
3. For the `.streamlit/config.toml`: click **Add file → Create new file**, type
   `.streamlit/config.toml` as the name (the `/` makes the folder), paste the file's
   contents, then commit. (Or upload the whole folder if your browser allows.)
4. Click **Commit changes**.

## Step 3 — Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io and **Sign in with GitHub** (authorise access).
2. Click **Create app** → **Deploy a public app from a repo** (private repos work too).
3. Fill in:
   - **Repository:** `your-username/shopee-stock-update`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Deploy**. First build takes ~2–3 minutes.
5. You get a URL like `https://shopee-stock-update.streamlit.app` — bookmark & share it.

## Step 4 — Use it
Open the URL → upload the 4 Excel files → **Generate** → download the two updated
Shopee files → upload them in Shopee Seller Centre → **Mass Update → Stock**.

## Updating the app later
Edit a file in GitHub (pencil icon) and commit — Streamlit Cloud auto-redeploys.

## Access control
- Private GitHub repo keeps the code private.
- In Streamlit Cloud → app **Settings → Sharing**, you can restrict who can view the app
  (invite specific Google emails) so only your team can use it.
