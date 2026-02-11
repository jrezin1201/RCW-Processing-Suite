# Deployment Guide for RCW Processing Tools

## Quick Start - Deploy to Railway

### Prerequisites
1. Create a GitHub account if you don't have one
2. Push your code to GitHub
3. Sign up for Railway.app (free tier available)

### Step 1: Prepare for Production

#### Add Production Files

**Procfile** (no extension, in project root):
```
web: uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

**runtime.txt** (specify Python version):
```
python-3.11.0
```

**.env.example** (for environment variables):
```
# Optional: External storage
AWS_S3_BUCKET=
AWS_ACCESS_KEY_ID=
AWS_SECRET_ACCESS_KEY=
```

### Step 2: Update Code for Production

#### Fix CORS (app/main.py)
```python
# Change from:
allow_origins=["*"]  # Development

# To:
allow_origins=[
    "https://your-domain.railway.app",
    "https://your-custom-domain.com"
]  # Production
```

#### Handle File Storage
Currently files are stored locally, which won't persist on most platforms. Options:

**Option A: Use Cloud Storage (Recommended)**
```python
# app/services/storage.py
import boto3
from io import BytesIO

def upload_to_s3(file_bytes: bytes, filename: str) -> str:
    s3 = boto3.client('s3')
    s3.upload_fileobj(
        BytesIO(file_bytes),
        os.environ['AWS_S3_BUCKET'],
        filename
    )
    return f"https://{os.environ['AWS_S3_BUCKET']}.s3.amazonaws.com/{filename}"
```

**Option B: Use Temporary Storage**
Most platforms provide `/tmp` directory for temporary files (files deleted on restart)

### Step 3: Deploy to Railway

1. **Connect GitHub to Railway:**
   ```bash
   # In your project directory
   git add .
   git commit -m "Add deployment configuration"
   git push origin main
   ```

2. **Create Railway Project:**
   - Go to [railway.app](https://railway.app)
   - Click "New Project"
   - Select "Deploy from GitHub repo"
   - Choose your repository
   - Railway will auto-deploy!

3. **Configure Environment (if needed):**
   - Click on your service
   - Go to "Variables" tab
   - Add any environment variables

4. **Get Your URL:**
   - Railway provides: `https://your-app.railway.app`
   - Share this with your users!

## Alternative Deployment Options

### Render.com
- Similar to Railway
- Add `render.yaml`:
```yaml
services:
  - type: web
    name: rcw-processing
    env: python
    buildCommand: "pip install -r requirements.txt"
    startCommand: "uvicorn app.main:app --host 0.0.0.0 --port $PORT"
    envVars:
      - key: PYTHON_VERSION
        value: 3.11.0
```

### Heroku
- Requires credit card (even for free tier)
- Add `Procfile` and `runtime.txt`
- Deploy with Heroku CLI:
```bash
heroku create your-app-name
git push heroku main
```

### DigitalOcean App Platform
- $5/month minimum
- Great for production apps
- Built-in database and storage options

### Google Cloud Run
- Pay per request (can be very cheap)
- Requires `Dockerfile`:
```dockerfile
FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]
```

## Production Checklist

### Essential Changes:
- [ ] CORS configured for production domains
- [ ] File storage solution (S3, CloudFlare R2, etc.)
- [ ] Environment variables for sensitive data
- [ ] Remove debug mode/reload
- [ ] Add proper logging
- [ ] SSL/HTTPS (usually automatic on platforms)

### Optional but Recommended:
- [ ] Error tracking (Sentry free tier)
- [ ] Custom domain name
- [ ] Rate limiting
- [ ] Authentication (if needed)

## Cost Estimates

### Free Options:
- **Railway**: Free tier (limited hours/month)
- **Render**: Free tier (spins down after inactivity)
- **Vercel**: Free tier (serverless, need adjustments)

### Paid Options:
- **Railway**: ~$5-10/month for always-on
- **Render**: $7/month for always-on
- **DigitalOcean**: $5/month minimum
- **AWS/GCP/Azure**: Pay per use (can be <$5/month for light use)

## Quick Deploy Commands

### For Railway (Recommended):
```bash
# 1. Install Railway CLI
npm install -g @railway/cli

# 2. Login
railway login

# 3. Initialize project
railway init

# 4. Deploy
railway up

# 5. Get your URL
railway open
```

### For Render:
```bash
# Just push to GitHub and connect via web UI
git push origin main
# Then go to render.com and connect your repo
```

## Troubleshooting

### Common Issues:

1. **"Module not found" errors**
   - Make sure all imports are in `requirements.txt`
   - Run `pip freeze > requirements.txt` locally

2. **File uploads fail**
   - Check file size limits (usually 100MB default)
   - Implement cloud storage for large files

3. **App crashes after deploy**
   - Check logs: `railway logs` or platform dashboard
   - Usually missing environment variables or dependencies

4. **"Port binding" errors**
   - Make sure to use `$PORT` environment variable
   - Don't hardcode port 8001

## Next Steps

1. **Start with Railway free tier** - It's the easiest
2. **Test with a few users** - Make sure it works
3. **Add monitoring** - Watch for errors
4. **Scale as needed** - Upgrade when you have more users

## Questions?

The app is production-ready for small-medium use cases. For enterprise deployment, consider:
- Load balancing
- Database for job tracking
- CDN for static assets
- Backup strategies
- CI/CD pipeline