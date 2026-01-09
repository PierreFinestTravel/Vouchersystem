# Free Deployment Guide - Finest Travel Voucher System

This guide covers **FREE** deployment options for your voucher system.

---

## 🏆 Recommended: Render.com (FREE)

**Why Render?**
- ✅ Completely free tier
- ✅ Docker support (needed for LibreOffice)
- ✅ Auto-deploy from GitHub
- ✅ Free SSL certificates
- ✅ Easy setup (5 minutes)

**Limitations:**
- App sleeps after 15 minutes of inactivity (wakes up on request, ~30s delay)
- 750 free hours/month (enough for a single service)

---

## 🚀 Quick Deploy to Render (5 minutes)

### Step 1: Push to GitHub

Make sure your code is on GitHub (already done ✅):
```
https://github.com/PierreFinestTravel/Vouchersystem.git
```

### Step 2: Create Render Account

1. Go to [render.com](https://render.com)
2. Click **"Get Started for Free"**
3. Sign up with your **GitHub account** (easiest)

### Step 3: Deploy from GitHub

1. In Render Dashboard, click **"New +"** → **"Web Service"**
2. Connect your GitHub account if not already connected
3. Find and select **`PierreFinestTravel/Vouchersystem`**
4. Configure settings:

| Setting | Value |
|---------|-------|
| **Name** | `finest-travel-vouchers` |
| **Region** | Frankfurt (or closest to you) |
| **Branch** | `main` |
| **Runtime** | `Docker` |
| **Plan** | `Free` |

5. Click **"Create Web Service"**

### Step 4: Wait for Build

- First build takes **10-15 minutes** (LibreOffice installation)
- Subsequent deploys are faster (~3-5 minutes)
- You'll see build logs in real-time

### Step 5: Access Your App

Once deployed, your app will be available at:
```
https://finest-travel-vouchers.onrender.com
```

---

## 📋 Alternative Free Options

### Option 2: Railway.app

**Pros:** $5 free credit/month, faster builds, no sleep
**Cons:** Credit runs out with heavy usage

```bash
# Install Railway CLI
npm install -g @railway/cli

# Login
railway login

# Deploy
railway init
railway up
```

### Option 3: Fly.io

**Pros:** Global edge deployment, generous free tier
**Cons:** More complex setup, requires CLI

```bash
# Install Fly CLI
# Windows: Download from https://fly.io/install/windows

# Login and deploy
fly auth login
fly launch
fly deploy
```

### Option 4: Google Cloud Run

**Pros:** 2M free requests/month, scales to zero
**Cons:** Requires Google Cloud account setup

---

## 🔧 Environment Variables

Set these in your hosting platform:

| Variable | Value | Required |
|----------|-------|----------|
| `PORT` | Auto-set by platform | No |
| `VOUCHER_TEMPLATE_PATH` | `/app/templates/_Voucher blank.docx` | No |

---

## 🐛 Troubleshooting

### App won't start
- Check build logs for errors
- Ensure `Dockerfile` is in root directory
- Verify all files are committed to GitHub

### PDF conversion failing
- LibreOffice needs memory - free tier should be sufficient
- Check logs: Render Dashboard → Your Service → Logs

### Slow first request
- Normal on free tier - app sleeps after 15 min inactivity
- First request after sleep takes ~30 seconds

### Build timeout
- Render allows 30 min builds (sufficient for LibreOffice)
- If still timing out, check Dockerfile for errors

---

## 📊 Free Tier Comparison

| Platform | Free Limits | Docker | Sleep | Best For |
|----------|-------------|--------|-------|----------|
| **Render** | 750 hrs/mo | ✅ | After 15 min | Best balance |
| **Railway** | $5 credit/mo | ✅ | No | Active usage |
| **Fly.io** | 3 shared VMs | ✅ | Configurable | Global apps |
| **Cloud Run** | 2M req/mo | ✅ | Yes | High traffic |

---

## 🔄 Auto-Deploy Setup

Render automatically deploys when you push to `main`:

```bash
# Make changes
git add .
git commit -m "Update feature"
git push

# Render detects push and deploys automatically!
```

---

## 🌐 Custom Domain (Optional)

1. In Render Dashboard → Your Service → Settings
2. Scroll to "Custom Domains"
3. Add your domain (e.g., `vouchers.finesttravel.africa`)
4. Update DNS records as instructed
5. Free SSL certificate is auto-generated

---

## 💡 Tips for Free Tier

1. **Keep app warm**: Use a free uptime monitor like [UptimeRobot](https://uptimerobot.com) to ping your app every 10 minutes

2. **Optimize Docker image**: The provided Dockerfile uses `libreoffice-writer-nogui` for smaller size

3. **Monitor usage**: Check Render dashboard for bandwidth/compute usage

---

## 📞 Support

- **Render Docs**: https://render.com/docs
- **Render Community**: https://community.render.com
- **GitHub Issues**: Report bugs in your repository

---

*Last updated: January 2026*

