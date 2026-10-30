# Desktop Packaging Approach for Question Categorizer

This directory contains the complete desktop packaging architecture, build specifications, CI/CD automated cloud workflows, locked branding assets, and user documentation for the **Question Categorizer Desktop App**.

---

## 📁 Directory Contents

| File / Folder | Purpose |
| :--- | :--- |
| **`icon.ico` / `icon.icns` / `icon.png`** | Official locked application icons for Windows, macOS, and high-res display. |
| **[`run_desktop.py`](run_desktop.py)** | Standalone desktop launcher. Detects open ports, runs Streamlit silently in the background, protects against console clicks, and automatically opens the user's default web browser. |
| **[`desktop.spec`](desktop.spec)** | PyInstaller specification blueprint with embedded custom branding. Compiles into `QuestionCategorizer.exe` (Windows) and `QuestionCategorizer.app` (macOS). |
| **[`.github/workflows/build-desktop.yml`](.github/workflows/build-desktop.yml)** | Automated GitHub Actions cloud CI/CD workflow. Builds 3 native targets: Windows x64, macOS Apple Silicon, and macOS Intel. |
| **[`build_local.sh`](build_local.sh)** | Shell script for building the macOS desktop application locally on a Mac. |
| **[`build_local.bat`](build_local.bat)** | Batch script for building the Windows desktop application locally on a Windows PC. |
| **[`HOW_TO_BUILD_AND_DEPLOY.md`](HOW_TO_BUILD_AND_DEPLOY.md)** | Step-by-step administrator guide for triggering cloud builds, publishing releases, and managing versions. |
| **[`USER_GUIDE.md`](USER_GUIDE.md)** | Non-technical end-user manual for colleagues explaining download, extraction, security prompts, and app usage. |

---

## 🚀 Quick Start: How to Enable Cloud Builds on GitHub

1. Commit and push the packaging files from your git repository (`Streamlit approach`):
   ```bash
   cd "Streamlit approach"
   git add run_desktop.py desktop.spec icon.ico icon.icns icon.png assets/ .github/workflows/build-desktop.yml .gitignore
   git commit -m "Lock custom Question Categorizer icon and enable automated cloud builds"
   git push origin main
   ```
2. Go to GitHub in your browser $\rightarrow$ **Actions** tab $\rightarrow$ select **"Build Standalone Desktop Applications"** $\rightarrow$ click **"Run workflow"**.
3. Download the ready-to-run `.zip` files under **Artifacts** when complete!

For full details, see [`HOW_TO_BUILD_AND_DEPLOY.md`](HOW_TO_BUILD_AND_DEPLOY.md).
