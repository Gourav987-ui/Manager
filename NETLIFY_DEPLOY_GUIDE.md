# Step-by-Step Netlify Deploy Guide

## Part 1: Fix the Build (Local)

### Step 1: Verify netlify.toml
Open `netlify.toml` in your project. It should contain:

```toml
[build]
  command = "echo 'No build required'"
  publish = "public"
```

This tells Netlify to run a simple command (instead of Hugo) and publish the `public` folder.

---

## Part 2: Push to Git

### Step 2: Stage the config file
```bash
git add netlify.toml
```

### Step 3: Commit
```bash
git commit -m "Fix Netlify build: remove Hugo command"
```

### Step 4: Push to your remote
```bash
git push origin main
```
*(Replace `main` with your branch name if different, e.g. `master`)*

---

## Part 3: Update Netlify Dashboard

### Step 5: Open Netlify
1. Go to [https://app.netlify.com](https://app.netlify.com)
2. Log in
3. Open the site that’s failing

### Step 6: Go to Build settings
1. Click **Site configuration** (or **Site settings**)
2. In the left sidebar, open **Build & deploy** → **Build settings**

### Step 7: Fix the build command
1. Click **Edit settings** → **Build settings**
2. Find **Build command**
3. Either:
   - **Clear it** (leave blank), so Netlify uses `netlify.toml`, or
   - Set it to: `echo 'No build required'`
4. Click **Save**

### Step 8: Trigger a new deploy
1. Go to **Deploys**
2. Click **Trigger deploy** → **Deploy site**
3. Wait for the build to finish

---

## Part 4: After Deploy

### What works
- Your site will load from the `public` folder
- HTML, CSS, and JavaScript will be served

### What does not work (important)
Test Sheet Manager is a **Node.js + Express** app. Netlify only hosts static files, not a Node server. This means:

- Login will not work
- Listing sheets will fail
- Upload, download, delete will fail

### For full functionality
Host the backend on a Node.js-capable platform, for example:

1. **Railway** – [railway.app](https://railway.app)
2. **Render** – [render.com](https://render.com)
3. **Fly.io** – [fly.io](https://fly.io)

I can create deployment config for one of these if you want to run the full app in the cloud.
