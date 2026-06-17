# GitHub Setup Instructions

## Step 1: Create a New Repository on GitHub

1. Go to https://github.com/new
2. **Repository name**: `xml-generator` (or your preferred name)
3. **Description**: "Generate sample XML from Excel mappings with auto-detected columns"
4. Choose **Public** (so GitHub Pages works) or **Private** (if you prefer)
5. Do NOT initialize with README (you already have one)
6. Click **Create repository**

## Step 2: Push Your Code

Open PowerShell in your workspace folder and run:

```powershell
git config user.name "Your Name"
git config user.email "your.email@example.com"
git add .
git commit -m "Initial commit: Python XML generator with browser UI"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/xml-generator.git
git push -u origin main
```

Replace `YOUR_USERNAME` with your actual GitHub username.

## Step 3: Enable GitHub Pages

1. Go to your repository on GitHub
2. Click **Settings** (gear icon)
3. On the left sidebar, click **Pages**
4. Under "Build and deployment":
   - **Source**: Select "Deploy from a branch"
   - **Branch**: Select `main`
   - **Folder**: Select `/ (root)`
5. Click **Save**

GitHub will automatically deploy your `docs/` folder to GitHub Pages.

## Step 4: Access Your Site

After a few seconds, you'll see a message at the top of the Pages settings with your site URL:

```
Your site is live at: https://YOUR_USERNAME.github.io/xml-generator/
```

Users can then access:
- **Browser UI**: `https://YOUR_USERNAME.github.io/xml-generator/docs/python_xml_generator_ui.html`
- **Main docs**: `https://YOUR_USERNAME.github.io/xml-generator/docs/index.html`

## Done!

Your project is now on GitHub and publicly accessible via GitHub Pages.

Users can:
- Use the browser UI without any installation
- Clone the repo and run the Python script locally
- See your README and documentation on GitHub
