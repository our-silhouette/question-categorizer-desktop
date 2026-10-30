import os
import re
import json
import uuid
import base64
import streamlit as st
import pandas as pd
import io
import streamlit.components.v1 as components

def render_base64_download_button(file_bytes, filename, button_text, mime_type, center=True):
    """
    Encodes file bytes into a base64 data URI and renders a pure client-side HTML5 download button.
    This bypasses server round-trips, prevents Streamlit script rerun desynchronization,
    and works reliably even if the ephemeral cloud server hibernates or restarts.
    """
    if isinstance(file_bytes, io.BytesIO):
        file_bytes = file_bytes.getvalue()
    b64_data = base64.b64encode(file_bytes).decode('utf-8')
    
    align_style = "text-align: center;" if center else ""
    display_style = "display: inline-flex; justify-content: center; align-items: center; min-width: 320px; max-width: 600px;" if center else "display: flex; justify-content: center; align-items: center; width: 100%;"
    
    html = f"""
    <div style="{align_style} margin-top: 8px; margin-bottom: 8px;">
        <a href="data:{mime_type};base64,{b64_data}" 
           download="{filename}" 
           class="custom-download-btn"
           style="
               {display_style}
               text-align: center;
               background-color: #2563EB;
               color: #FFFFFF !important;
               -webkit-text-fill-color: #FFFFFF !important;
               border: 1px solid #1D4ED8;
               border-radius: 8px;
               font-weight: 600;
               font-size: 1rem;
               padding: 0.65rem 1.5rem;
               text-decoration: none;
               transition: all 0.2s ease-in-out;
               box-sizing: border-box;
               box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2);
           "
        >
           <span class="custom-download-text" style="color: #FFFFFF !important; -webkit-text-fill-color: #FFFFFF !important; font-weight: 600; text-decoration: none;">{button_text}</span>
        </a>
    </div>
    """
    return html

# Import pipeline orchestration engine with dynamic reload to prevent stale module caching
import importlib
import pipeline
importlib.reload(pipeline)
from pipeline import run_categorization_pipeline, format_excel_outputs

# Page Setup
st.set_page_config(
    page_title="Question Paper Categorizer",
    page_icon="📚",
    layout="wide"
)

# Custom Styling (Explicit Vibrant Blue Button & Error Card Rules)
st.markdown("""
<style>
    .main .block-container {
        padding-top: 1.5rem;
        padding-bottom: 3rem;
        max-width: 1200px;
    }
    div.stButton > button {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        border: 1px solid #1D4ED8 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.5rem 1rem !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button:hover {
        background-color: #1D4ED8 !important;
        color: #FFFFFF !important;
        border-color: #1E40AF !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2) !important;
    }
    div.stButton > button:active {
        background-color: #1E40AF !important;
        color: #FFFFFF !important;
    }
    /* Vibrant Bright White Download Button Matching Start Pipeline Button */
    .stMarkdown a.custom-download-btn,
    .stMarkdown a.custom-download-btn:link,
    .stMarkdown a.custom-download-btn:visited,
    .stMarkdown a.custom-download-btn:hover,
    .stMarkdown a.custom-download-btn:active,
    .stMarkdown a.custom-download-btn:focus,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:link,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:visited,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:hover,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:active,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:focus {
        background-color: #2563EB !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        border: 1px solid #1D4ED8 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
        padding: 0.65rem 1.5rem !important;
        text-decoration: none !important;
        cursor: pointer !important;
        box-shadow: 0 2px 4px rgba(37, 99, 235, 0.2) !important;
        transition: all 0.2s ease-in-out !important;
        opacity: 1 !important;
    }
    .stMarkdown a.custom-download-btn:hover,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:hover {
        background-color: #1D4ED8 !important;
        color: #FFFFFF !important;
        border-color: #1E40AF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2) !important;
    }
    .stMarkdown a.custom-download-btn:active,
    div[data-testid="stMarkdownContainer"] a.custom-download-btn:active {
        background-color: #1E40AF !important;
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
    }
    .custom-download-btn span,
    .custom-download-text {
        color: #FFFFFF !important;
        -webkit-text-fill-color: #FFFFFF !important;
        font-weight: 600 !important;
        text-decoration: none !important;
    }
    .metric-card {
        background: #F8FAFC;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-val {
        font-size: 1.8rem;
        font-weight: 700;
        color: #1E293B;
    }
    .metric-lbl {
        font-size: 0.85rem;
        font-weight: 600;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
</style>
""", unsafe_allow_html=True)

# Main Title & Subtitle
st.title("📚 Question Paper Categorizer")
st.markdown("Automated Knowledge Component (KC) mapping, difficulty analysis, and outlier isolation powered by Gemini AI.")

# Session State Initialization
if "saved_api_key" not in st.session_state:
    st.session_state.saved_api_key = ""

# Sidebar Configuration
with st.sidebar:
    st.header("⚙️ Configuration")
    
    # API Key Input (Masked, In-Memory Session Persistence, Supports Multi-Key Pool)
    api_key_input = st.text_input(
        "Gemini API Key(s):",
        value=st.session_state.saved_api_key,
        type="password",
        help="Paste your Google AI Studio API Key. Enter multiple keys separated by commas for instant zero-wait rate-limit pooling!"
    )
    
    if api_key_input:
        st.session_state.saved_api_key = api_key_input.strip()
        
    if st.session_state.saved_api_key:
        keys_count = len([k for k in st.session_state.saved_api_key.replace(";", ",").split(",") if k.strip()])
        if keys_count > 1:
            st.success(f"⚡ **Active Key Pool:** `{keys_count}` Free Keys loaded for instant 0s rotation!")
        if st.button("🗑️ Clear Key(s) from Memory", help="Clears the current API key(s) from this session"):
            st.session_state.saved_api_key = ""
            st.rerun()
            
    st.markdown("""
    💡 **Get 100% Free Gemini API Keys:**  
    Create free keys in 10 seconds at:  
    👉 [Google AI Studio API Keys](https://aistudio.google.com/app/apikey)  
    *(Tip: Creating 2 or 3 keys in separate projects enables instant zero-wait execution!)*
    """, unsafe_allow_html=True)
    
    st.divider()
    
    st.info("🤖 **Engine:** Stable Production Flash Suite (`gemini-3.8-flash`, `gemini-flash-latest`, `gemini-flash-lite-latest`)")
    
    try:
        import subprocess
        commit_hash = subprocess.check_output(["git", "rev-parse", "--short", "HEAD"]).decode("utf-8").strip()
    except Exception:
        commit_hash = "main"
    st.caption(f"📌 **Deployed Version:** `{commit_hash}` *(Branch: `main`)*")

# Session state key for uploader reset
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

import json

CHECKPOINT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scratch", "categorizer_checkpoint.json")

def load_disk_checkpoint():
    try:
        if os.path.exists(CHECKPOINT_PATH):
            with open(CHECKPOINT_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        print(f"Could not load disk checkpoint: {e}", flush=True)
    return None

def save_disk_checkpoint(data):
    try:
        os.makedirs(os.path.dirname(CHECKPOINT_PATH), exist_ok=True)
        with open(CHECKPOINT_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception as e:
        print(f"Could not save disk checkpoint: {e}", flush=True)

def clear_disk_checkpoint():
    try:
        if os.path.exists(CHECKPOINT_PATH):
            os.remove(CHECKPOINT_PATH)
    except Exception as e:
        pass

# Session State for Results & Checkpoint Persistence
if "results" not in st.session_state:
    st.session_state.results = None
if "checkpoint" not in st.session_state:
    st.session_state.checkpoint = load_disk_checkpoint()
if "partial_deliverables" not in st.session_state:
    st.session_state.partial_deliverables = None
if "current_run_id" not in st.session_state:
    st.session_state.current_run_id = None

# Auto-hydrate partial deliverables dynamically from latest disk checkpoint
if st.session_state.checkpoint and not st.session_state.results:
    chk = st.session_state.checkpoint
    try:
        p_master, p_metrics = format_excel_outputs(
            chk["all_conforming"], chk["all_non_conforming"], chk["master_kcs"]
        )
        st.session_state.partial_deliverables = {
            "master_bytes": p_master,
            "metrics": p_metrics,
            "completed_chunks": chk["completed_chunks"],
            "total_chunks": chk["total_chunks"],
            "last_page": chk["last_processed_page"],
            "pdf_base_name": os.path.splitext(chk.get("pdf_name", "Document.pdf"))[0]
        }
    except Exception as ex:
        print(f"Could not auto-generate partial deliverables from disk checkpoint: {ex}", flush=True)

# Main Form Area
col1, col2 = st.columns(2)

with col1:
    uploaded_pdf = st.file_uploader(
        "📄 Upload Question Paper (PDF or Word DOCX):",
        type=["pdf", "docx", "doc"],
        key=f"pdf_uploader_{st.session_state.uploader_key}",
        help="Upload the PDF or Word document containing questions to classify."
    )

with col2:
    uploaded_skeleton = st.file_uploader(
        "📊 Upload KC List (Excel, CSV, or Word DOCX):",
        type=["xlsx", "xls", "docx", "doc", "csv"],
        key=f"kc_uploader_{st.session_state.uploader_key}",
        help="Upload the master KC hierarchy (Excel .xlsx/.xls, CSV, or Word .docx). Note: NEET KC documents must have the word 'NEET' in the file name."
    )
    st.caption("ℹ️ **NEET Syllabus Note:** NEET KC documents must have the word **NEET** in the file name.")
    if uploaded_skeleton and re.search(r'\bNEET\b', getattr(uploaded_skeleton, 'name', ''), re.IGNORECASE):
        st.info("🎯 **NEET KC Mode Active:** Competency will be calibrated purely on cognitive depth (zero historical exam benchmarking).")

st.divider()

# Checkpoint Banner & Resume Controls
resume_clicked = False
discard_checkpoint = False

if st.session_state.partial_deliverables and not st.session_state.results:
    p_deliv = st.session_state.partial_deliverables
    p_metrics = p_deliv["metrics"]
    st.warning(f"""
    ### ⏯️ Interrupted Run Checkpoint Detected
    - **Progress Salvaged:** Chunk **{p_deliv['completed_chunks']} of {p_deliv['total_chunks']}** (Pages up to **{p_deliv['last_page']}**)
    - **Questions Preserved:** **{p_metrics['total_scanned']} total** ({p_metrics['conforming']} conforming mapped, {p_metrics['non_conforming']} outliers)
    - **Zero Work Lost:** You can download the partial deliverables below or resume from Page {p_deliv['last_page'] + 1}.
    """)
    
    res_btn_col1, res_btn_col2 = st.columns(2)
    with res_btn_col1:
        resume_clicked = st.button(f"▶️ Resume Execution from Page {p_deliv['last_page'] + 1} (Chunk {p_deliv['completed_chunks'] + 1} of {p_deliv['total_chunks']})", use_container_width=True)
    with res_btn_col2:
        if st.button("🔄 Discard Checkpoint & Start Fresh", use_container_width=True):
            st.session_state.checkpoint = None
            st.session_state.partial_deliverables = None
            clear_disk_checkpoint()
            st.rerun()
            
    # Partial Deliverable Download (Client-Side In-Memory Base64 Streaming)
    p_pdf_base = p_deliv.get("pdf_base_name", "Document")
    st.markdown(
        render_base64_download_button(
            file_bytes=p_deliv["master_bytes"],
            filename=f"{p_pdf_base}_Q_Analysis_PARTIAL.xlsx",
            button_text=f"📄 Download {p_pdf_base}_Q_Analysis_PARTIAL.xlsx",
            mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            center=True
        ),
        unsafe_allow_html=True
    )
    st.divider()

# Process Action & Symmetrical Button Row
btn_col1, btn_col2 = st.columns(2)

with btn_col1:
    run_clicked = st.button("🚀 Start Categorization Pipeline", use_container_width=True)

with btn_col2:
    if st.session_state.results or st.session_state.partial_deliverables:
        if st.button("🧹 Clear Results & Reset App", use_container_width=True, help="Clears current output previews, resets file uploaders, and prepares the app for a new document"):
            st.session_state.results = None
            st.session_state.checkpoint = None
            st.session_state.partial_deliverables = None
            clear_disk_checkpoint()
            st.session_state.uploader_key += 1
            st.rerun()

if run_clicked or resume_clicked:
    current_key = st.session_state.saved_api_key or api_key_input
    if not current_key:
        st.error("⚠️ **Missing API Key:** Please enter your Google Gemini API Key in the sidebar before running.")
    elif not uploaded_pdf:
        st.error("📄 **Missing Document:** Please upload a Question Paper (PDF or Word DOCX).")
    elif not uploaded_skeleton:
        st.error("📊 **Missing KC List:** Please upload a Master KC List (Excel, CSV, or Word DOCX).")
    else:
        import uuid
        run_id = str(uuid.uuid4())
        st.session_state.current_run_id = run_id
        
        progress_bar = st.progress(0.0)
        status_text = st.empty()
        
        def update_progress(pct, msg):
            if pct is not None:
                progress_bar.progress(pct)
            if msg:
                status_text.info(msg)
            
        def on_chunk_finished(completed_chunks, total_chunks, all_conf, all_non_conf, last_page, master_kcs):
            chk_data = {
                "pdf_name": getattr(uploaded_pdf, 'name', 'Document.pdf'),
                "completed_chunks": completed_chunks,
                "total_chunks": total_chunks,
                "all_conforming": list(all_conf),
                "all_non_conforming": list(all_non_conf),
                "last_processed_page": last_page,
                "master_kcs": master_kcs
            }
            st.session_state.checkpoint = chk_data
            save_disk_checkpoint(chk_data)
            
        start_chunk_idx = 0
        existing_conf = []
        existing_non_conf = []
        
        if resume_clicked and st.session_state.checkpoint:
            chk = st.session_state.checkpoint
            start_chunk_idx = chk["completed_chunks"]
            existing_conf = chk["all_conforming"]
            existing_non_conf = chk["all_non_conforming"]
            update_progress(start_chunk_idx / chk["total_chunks"], f"Resuming execution from Chunk {start_chunk_idx + 1} of {chk['total_chunks']} (Page {chk['last_processed_page'] + 1})...")
        else:
            st.session_state.checkpoint = None
            st.session_state.partial_deliverables = None
            
        try:
            with st.spinner("Processing document and classifying questions..."):
                master_bytes, metrics = run_categorization_pipeline(
                    api_key=current_key.strip(),
                    pdf_file_source=uploaded_pdf,
                    skeleton_file_source=uploaded_skeleton,
                    progress_callback=update_progress,
                    run_id=run_id,
                    start_chunk_idx=start_chunk_idx,
                    existing_conforming=existing_conf,
                    existing_non_conforming=existing_non_conf,
                    chunk_completed_callback=on_chunk_finished
                )
                
            if master_bytes is None:
                status_text.warning("Run superseded by a newer execution.")
                st.stop()
                
            update_progress(1.0, "Categorization Complete!")
            pdf_base_name = os.path.splitext(uploaded_pdf.name)[0] if hasattr(uploaded_pdf, 'name') else "Master_Categorized"
            st.session_state.results = {
                "master_bytes": master_bytes,
                "metrics": metrics,
                "pdf_base_name": pdf_base_name
            }
            st.session_state.checkpoint = None
            st.session_state.partial_deliverables = None
            clear_disk_checkpoint()
            st.success("Successfully categorized question paper!")
            st.rerun()
        except Exception as e:
            err_type = type(e).__name__
            err_msg = str(e)
            err_msg_lower = err_msg.lower()
            
            import traceback
            traceback.print_exc()
            print(f"PIPELINE EXECUTION ERROR [{err_type}]: {err_msg}", flush=True)
            
            # Salvage partial deliverables if checkpoint has data
            if st.session_state.checkpoint and (st.session_state.checkpoint.get("all_conforming") or st.session_state.checkpoint.get("all_non_conforming")):
                chk = st.session_state.checkpoint
                try:
                    p_master, p_metrics = format_excel_outputs(
                        chk["all_conforming"], chk["all_non_conforming"], chk["master_kcs"]
                    )
                    st.session_state.partial_deliverables = {
                        "master_bytes": p_master,
                        "metrics": p_metrics,
                        "completed_chunks": chk["completed_chunks"],
                        "total_chunks": chk["total_chunks"],
                        "last_page": chk["last_processed_page"],
                        "pdf_base_name": os.path.splitext(chk["pdf_name"])[0]
                    }
                except Exception as ex:
                    print(f"Error generating partial salvage deliverables: {ex}", flush=True)
            
            # 1. Clean handling of Streamlit stop button
            if "StopException" in err_type or "ScriptControlException" in err_type:
                status_text.warning("⏹️ Execution stopped by user.")
                st.stop()
                
            # 2. Graceful 429 Quota / Rate Limit Handler
            elif "429" in err_msg or "resource_exhausted" in err_msg_lower or "quota" in err_msg_lower:
                status_text.empty()
                progress_bar.empty()
                st.error("""
                ### ⏳ Google Rate-Limit Cooldown Active (HTTP 429)
                Google API's per-minute request rate limit was momentarily reached during this intensive scan.

                **What to do:**
                1. Please wait **30 to 60 seconds** for Google's rolling minute window to reset.
                2. Click **Start Categorization Pipeline** or **Resume** above to continue without losing progress.
                3. Alternatively, generate a fresh project key at [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey) (**Create API Key $\rightarrow$ Create in new project**).
                """)
                
            # 3. Graceful 403/400 Invalid API Key Error Handler
            elif "api_key" in err_msg_lower or "invalid" in err_msg_lower or "403" in err_msg or "400" in err_msg:
                status_text.empty()
                progress_bar.empty()
                st.error("""
                ### 🔑 Invalid or Expired API Key
                The Gemini API key entered in the sidebar is invalid, expired, or typed incorrectly.

                **How to Resolve:**
                1. Verify your key at [aistudio.google.com/app/apikey](https://aistudio.google.com/app/apikey).
                2. Paste the valid key into the sidebar and re-run.
                """)
                
            # 4. Graceful Network / Timeout Handler
            elif "connection" in err_msg_lower or "timeout" in err_msg_lower or "dns" in err_msg_lower:
                status_text.empty()
                progress_bar.empty()
                st.error("""
                ### 📡 Network Connection Timeout
                Could not establish a connection to Google AI Studio servers.

                **How to Resolve:**
                1. Please check your internet connection.
                2. Click **Resume** above to retry the remaining pages.
                """)
                
            # 5. General Fallback with Technical Accordion
            else:
                status_text.empty()
                progress_bar.empty()
                st.error("### ⚠️ Pipeline Execution Error")
                st.warning(f"An unexpected issue occurred while processing: **{err_msg}**")
                with st.expander("🔍 View Technical Error Details"):
                    st.code(err_msg, language="bash")

# Display Results if available
if st.session_state.results:
    res = st.session_state.results
    metrics = res["metrics"]
    
    st.subheader("📊 Execution Summary")
    
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-val">{metrics["total_scanned"]}</div>
            <div class="metric-lbl">Total Scanned</div>
        </div>
        ''', unsafe_allow_html=True)
    with m_col2:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-val" style="color: #16A34A;">{metrics["conforming"]}</div>
            <div class="metric-lbl">Conforming Mapped</div>
        </div>
        ''', unsafe_allow_html=True)
    with m_col3:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-val" style="color: #DC2626;">{metrics["non_conforming"]}</div>
            <div class="metric-lbl">Outliers Isolated</div>
        </div>
        ''', unsafe_allow_html=True)
    with m_col4:
        st.markdown(f'''
        <div class="metric-card">
            <div class="metric-val" style="color: #9333EA;">{metrics["mapped_kcs"]}</div>
            <div class="metric-lbl">KCs Covered</div>
        </div>
        ''', unsafe_allow_html=True)

    st.divider()
    
    # Download Section (Single Unified Dual-Sheet Excel via Client-Side Base64 Streaming)
    st.subheader("📥 Download Deliverables")
    
    pdf_base = res.get("pdf_base_name", "Document")
    master_filename = f"{pdf_base}_Q_Analysis.xlsx" if pdf_base else "Question_Paper_Q_Analysis.xlsx"
    
    st.markdown(
        render_base64_download_button(
            file_bytes=res["master_bytes"],
            filename=master_filename,
            button_text=f"📄 Download {master_filename}",
            mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            center=True
        ),
        unsafe_allow_html=True
    )

    st.divider()

    # Previews
    st.subheader("🔍 Data Previews")
    tab1, tab2 = st.tabs(["📊 Categorized Questions (Tab 1)", "⚠️ Non-Conforming Questions (Tab 2)"])
    
    with tab1:
        try:
            df_master_preview = pd.read_excel(io.BytesIO(res["master_bytes"]), sheet_name="Categorized Questions")
            st.dataframe(df_master_preview, use_container_width=True)
        except Exception as e:
            st.warning(f"Could not render Categorized Questions preview: {e}")
        
    with tab2:
        try:
            df_outliers_preview = pd.read_excel(io.BytesIO(res["master_bytes"]), sheet_name="Non-Conforming Questions")
            if df_outliers_preview.empty:
                st.info("No non-conforming questions found. All questions conformed to standard formats.")
            else:
                st.dataframe(df_outliers_preview, use_container_width=True)
        except Exception as e:
            st.info("No non-conforming questions found or could not render preview.")

