# Complete Guide: Desktop Packaging & Cloud CI/CD Deployment

This document explains the end-to-end architecture, implementation, and build process for the **Question Categorizer Standalone Desktop Application** (Windows & macOS).

---

## 1. Architectural Overview

```
                                      [ GitHub Repository ]
                                                │
                          ┌─────────────────────┴─────────────────────┐
                          │                                           │
                          ▼                                           ▼
            [ Streamlit Community Cloud ]              [ GitHub Actions Cloud CI/CD ]
            (Live Web App on Streamlit Cloud)         (On-Demand Desktop Builders)
                          │                                           │
                          │                     ┌─────────────────────┼─────────────────────┐
                          │                     ▼                     ▼                     ▼
                          │             [ Windows Runner ]    [ macOS ARM64 ]       [ macOS Intel ]
                          │                     │                     │                     │
                          │             QuestionCategorizer-  QuestionCategorizer-  QuestionCategorizer-
                          │             Windows.zip           macOS-AppleSilicon.zip macOS-Intel.zip
                          │                     │                     │                     │
                          ▼                     └─────────────────────┼─────────────────────┘
                  [ Cloud Users ]                                     ▼
                (Access via browser)                          [ Desktop Users ]
                                                      (Download zip, run locally)
```

### Key Highlights
1. **Zero Impact on Cloud Streamlit:** `app.py`, `pipeline.py`, and `requirements.txt` remain untouched.
2. **Custom Locked Branding:** Bundles the official high-tech celestial crest icon (`icon.ico` for Windows, `icon.icns` for macOS).
3. **Triple-Target Native Architecture:**
   - Windows 64-bit (`.exe`)
   - Apple Silicon macOS (`.app` for M1/M2/M3/M4)
   - Intel Chipset macOS (`.app` for legacy Intel Macs)
4. **Self-Contained Executable:** End users **do not** need Python, Git, or a terminal.

---

## 2. Directory Structure & File Roles

```
Question paper categorization/
├── Streamlit approach/                    # Active Git repository
│   ├── app.py                            # Streamlit user interface
│   ├── pipeline.py                       # Categorization & AI pipeline
│   ├── requirements.txt                  # Python dependencies
│   ├── icon.ico                          # Windows application icon
│   ├── icon.icns                         # macOS application icon
│   ├── icon.png                          # High-resolution PNG logo
│   ├── run_desktop.py                    # Background launcher & browser opener
│   ├── desktop.spec                      # PyInstaller packaging blueprint
│   └── .github/
│       └── workflows/
│           └── build-desktop.yml         # GitHub Actions automated cloud build
│
└── Desktop packaging approach/           # Dedicated reference & documentation folder
    ├── icon.ico
    ├── icon.icns
    ├── icon.png
    ├── run_desktop.py
    ├── desktop.spec
    ├── build_local.sh                    # Helper script for local Mac builds
    ├── build_local.bat                   # Helper script for local Windows builds
    ├── .github/workflows/build-desktop.yml
    ├── HOW_TO_BUILD_AND_DEPLOY.md        # This guide
    └── USER_GUIDE.md                     # Colleague instruction manual
```

---

## 3. How to Trigger the Cloud Build on GitHub (3 Steps)

### Step 1: Commit and Push to GitHub
```bash
cd "/Users/Deepu/Everything AI/Projects/Question paper categorization/Streamlit approach"
git add run_desktop.py desktop.spec icon.ico icon.icns icon.png assets/ .github/workflows/build-desktop.yml .gitignore
git commit -m "Lock custom Question Categorizer icon and enable automated cloud builds"
git push origin main
```

### Step 2: Trigger the Build in Your Browser
1. Go to: `https://github.com/our-silhouette/question-paper-categorizer`
2. Click the **Actions** tab at the top.
3. In the left sidebar, click **"Build Standalone Desktop Applications"**.
4. Click **"Run workflow"** $\rightarrow$ click the green **"Run workflow"** button.

### Step 3: Download Finished Apps
When the run finishes (~5 minutes), scroll to the **Artifacts** section at the bottom:
* 📥 **`QuestionCategorizer-Windows`**
* 📥 **`QuestionCategorizer-macOS-AppleSilicon`** (for M1/M2/M3/M4)
* 📥 **`QuestionCategorizer-macOS-Intel`** (for older Intel Macs)

Users simply unzip and double-click to launch!
