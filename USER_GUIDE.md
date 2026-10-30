# User Guide: Question Categorizer (Desktop App)

Welcome to the **Question Categorizer Desktop App**! This guide explains how to install, launch, and use the application on **Windows** and **macOS** computers.

---

## 1. System Requirements

* **Operating System:** Windows 10/11 (64-bit) OR macOS 11+ (Apple Silicon or Intel)
* **Web Browser:** Google Chrome, Microsoft Edge, Safari, or Firefox
* **Internet Connection:** Active connection (to communicate with Gemini AI)
* **Google Gemini API Key:** Free key from [Google AI Studio](https://aistudio.google.com/app/apikey)
* **Prerequisites:** **NONE!** You do *not* need Python, Git, or a terminal.

---

## 2. Windows Installation & Launch

1. **Download:** Download `QuestionCategorizer-Windows.zip`.
2. **Extract:** Right-click the `.zip` file and select **"Extract All..."**, then choose a folder (e.g. `Desktop` or `Documents`).
3. **Launch:** Open the extracted folder and double-click:
   ```
   QuestionCategorizer.exe
   ```
4. **First-Time Security Prompt (SmartScreen):**
   * Windows may display a blue dialog: *"Windows protected your PC"*.
   * Click **"More info"**, then click **"Run anyway"**. *(Only needed once).*
5. **App Starts:** A small console window appears, and within a few seconds **your default web browser will automatically open** showing the Question Categorizer!

---

## 3. macOS Installation & Launch

1. **Download:**
   * For M1/M2/M3/M4 Macs: Download `QuestionCategorizer-macOS-AppleSilicon.zip`.
   * For older Intel Macs: Download `QuestionCategorizer-macOS-Intel.zip`.
2. **Extract:** Double-click the `.zip` file to extract `QuestionCategorizer.app`.
3. **Move:** Drag `QuestionCategorizer.app` into your **Applications** folder.
4. **First-Time Security Prompt (Gatekeeper):**
   * **Right-click (or Control-click)** on `QuestionCategorizer.app` and choose **"Open"**.
   * When prompted *"macOS cannot verify the developer"*, click **"Open"**. *(Only needed once).*
5. **App Starts:** Safari or Chrome will automatically open with the application running!

---

## 4. How to Use the Application

### Step 1: Enter your Gemini API Key
* In the left sidebar under **⚙️ Configuration**, paste your Gemini API Key.
* *(Get a free key in 10 seconds at: https://aistudio.google.com/app/apikey)*

### Step 2: Upload Documents
* **Question File:** Upload your question document in **PDF** or **Word (.docx)** format.
* **Master Knowledge Components (KCs):** Upload your syllabus/curriculum hierarchy in **Excel (.xlsx)** or **Word (.docx)** format.

### Step 3: Run Categorization
* Click the blue **"🚀 Start Categorization Pipeline"** button.

### Step 4: Download Results
* Once complete, click the blue **"📥 Download Categorization Report (.xlsx)"** button to save your analyzed spreadsheet.

---

## 5. How to Exit

* Close your browser tab.
* In the background console window, press `Ctrl + C` or close the window to shut down the server.
