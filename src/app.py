import streamlit as st
import datetime
import os
from generator import create_photo_report, analyze_docx_structure
from utils import load_image, crop_to_ratio, resize_with_padding
from streamlit_sortables import sort_items

# Page Config
st.set_page_config(
    page_title="現況照片清冊生成器",
    page_icon="📷",
    layout="wide"
)

# Custom CSS implementation
st.markdown("""
<style>
    /* Main Background & Fonts */
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* Custom Header Style */
    .main-header {
        background: linear-gradient(90deg, #0E4E8E 0%, #003366 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: white !important;
        margin: 0;
        font-size: 2.2rem;
        font-weight: 700;
    }
    .main-header p {
        color: #e0e0e0;
        margin-top: 0.5rem;
        font-size: 1rem;
    }

    /* Card Style */
    div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlock"] {
        /* Target generic containers if possible, but st.container(border=True) 
           adds specific classes. We rely on the border=True mostly. */
    }
    
    /* Button Styles */
    .stButton>button {
        width: 100%;
        border-radius: 6px;
        font-weight: 600;
    }
    
    /* Primary Action Button (Generate) */
    div[data-testid="stButton"] > button[kind="primary"] {
        background-color: #0E4E8E;
        border: none;
        padding: 0.75rem 0;
        font-size: 1.1rem;
        transition: all 0.3s ease;
    }
    div[data-testid="stButton"] > button[kind="primary"]:hover {
        background-color: #003366;
        box-shadow: 0 4px 12px rgba(14, 78, 142, 0.3);
    }

    /* Delete Button - Red Warning */
    button[key*="del_"] {
        background-color: transparent;
        color: #d9534f;
        border: 1px solid #d9534f;
    }
    button[key*="del_"]:hover {
        background-color: #d9534f;
        color: white;
    }

    /* Input Fields */
    .stTextInput > div > div > input {
        border-radius: 4px;
    }
</style>
""", unsafe_allow_html=True)

# Custom Header
st.markdown("""
<div class="main-header">
    <h1>📷 現況照片清冊生成器</h1>
    <p>快速生成標準化 Word 蒐證照片報表 • 支援自動排版與 EXIF 資料處理</p>
</div>
""", unsafe_allow_html=True)

# --- Sidebar ---
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=64) # Placeholder generic icon or use local if available
    st.title("設定面板")
    
    with st.expander("📖 使用教學 (User Guide)"):
        st.markdown("""
        **操作流程：**
        1. **設定**：填寫案由、地點等資訊。
        2. **上傳**：拖曳照片至右側區塊。
        3. **排序**：在側邊欄拖曳調整順序。
        4. **編輯**：輸入個別說明與時間。
        5. **輸出**：點擊底部按鈕生成報表。
        """)
    
    st.markdown("---")

    st.subheader("📋 報表全域資訊")
    # Global Fields
    subject = st.text_input("案由 (Project)", value="", placeholder="例：竊盜案現場勘查")
    report_header = st.text_input("頁首標題 (Header)", value="臺南市政府警察局新化分局蒐證照片")
    location = st.text_input("地點 (Location)", value="", placeholder="例：新化區中山路...")
    maker = st.text_input("製作人 (Maker)", value="")
    report_date = st.date_input("日期 (Date)", value=None)
    global_description = st.text_area("全域說明 (Global Desc.)", value="", help="未填寫個別說明的照片將自動套用此說明")
    
    st.markdown("---")
    
    # Layout & Template - Accordion
    with st.expander("📄 模板與排版設定 (Templates)", expanded=False):
        # Use Absolute Path for Assets to fix portable version issues
        # Locate 'assets' relative to 'src/app.py' -> '../assets'
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        assets_dir = os.path.join(project_root, "assets")

        if not os.path.exists(assets_dir):
            os.makedirs(assets_dir)
            
        templates = [f for f in os.listdir(assets_dir) if f.endswith(".docx") and not f.startswith("~$")]
        
        selected_template = None
        layout_style_code = "A4_Vertical" # Default
        
        if templates:
            template_name = st.selectbox("選擇 Word 模板", templates)
            selected_template = os.path.join(assets_dir, template_name)
            
            # Auto-detect Layout Mode based on filename
            if "左右" in template_name or "Side" in template_name:
                layout_style_code = "A4_SideBySide"
                st.caption("ℹ️ 模式：雙欄對照 (Side-by-Side)")
            else:
                layout_style_code = "A4_Vertical"
                st.caption("ℹ️ 模式：直式標準 (Vertical)")

            # Analyzer Button
            if st.checkbox("🔍 顯示模板結構分析 (Debug)"):
                if st.button("開始分析"):
                    structure = analyze_docx_structure(selected_template)
                    st.text_area("分析結果", structure, height=300)
        else:
            st.error(f"❌ 無模板 (請將 .docx 放至 {assets_dir})")

# --- Main Content ---

uploaded_files = st.file_uploader(
    "📤 上傳照片 (支援 JPG, PNG, HEIC)", 
    type=['jpg', 'jpeg', 'png', 'heic'], 
    accept_multiple_files=True,
    help="可一次選擇多張照片，支援拖放上傳"
)



photos_data = []

if uploaded_files:
    # Sidebar - Photo Sorting
    with st.sidebar:
        st.markdown("---")
        st.subheader("🔃 照片排序")
        st.caption("拖曳下方項目以調整順序")
        
        # Session State for File Management
        if 'managed_files' not in st.session_state:
            st.session_state.managed_files = {} # {filename: file_obj}
        if 'file_order' not in st.session_state:
            st.session_state.file_order = []    # [filename1, filename2...]
        if 'deleted_files' not in st.session_state:
            st.session_state.deleted_files = set() # {filename}
        if 'delete_history' not in st.session_state:
            st.session_state.delete_history = [] # List of filenames

        # Sync uploaded_files with Session State
        current_filenames = [f.name for f in uploaded_files]
        
        # Add new files
        for f in uploaded_files:
            # Only add if NOT in managed files AND NOT explicitly deleted
            if f.name not in st.session_state.managed_files:
                if f.name not in st.session_state.deleted_files:
                    st.session_state.managed_files[f.name] = f
                    st.session_state.file_order.append(f.name)
        
        # Remove deleted files (This handles if user explicitly removes from widget)
        keys_to_remove = []
        for fname in st.session_state.managed_files:
            if fname not in current_filenames:
                keys_to_remove.append(fname)
        
        for k in keys_to_remove:
            del st.session_state.managed_files[k]
            if k in st.session_state.file_order:
                st.session_state.file_order.remove(k)

        # Undo Delete Button
        if st.session_state.delete_history:
            last_entry = st.session_state.delete_history[-1]
            if isinstance(last_entry, tuple):
                last_idx, last_fname = last_entry
            else:
                last_idx, last_fname = -1, last_entry

            if st.button(f"↩️ 復原刪除 ({last_fname})"):
                # Undo Logic
                entry_to_restore = st.session_state.delete_history.pop()
                idx_to_restore, fname_to_restore = entry_to_restore if isinstance(entry_to_restore, tuple) else (-1, entry_to_restore)
                
                # 1. Remove from deleted_files
                if fname_to_restore in st.session_state.deleted_files:
                    st.session_state.deleted_files.remove(fname_to_restore)
                
                # 2. Find file object and restore
                file_obj = next((f for f in uploaded_files if f.name == fname_to_restore), None)
                if file_obj:
                    st.session_state.managed_files[fname_to_restore] = file_obj
                    if idx_to_restore != -1 and idx_to_restore <= len(st.session_state.file_order):
                        st.session_state.file_order.insert(idx_to_restore, fname_to_restore)
                    else:
                        st.session_state.file_order.append(fname_to_restore)
                st.rerun()

        # Sortable Component
        if st.session_state.file_order:
            display_order = [f"{i+1}. {fname}" for i, fname in enumerate(st.session_state.file_order)]
            sorted_display = sort_items(display_order, direction='vertical')
            
            new_order = []
            for item in sorted_display:
                parts = item.split('. ', 1)
                new_order.append(parts[1] if len(parts) == 2 else item)
            
            if new_order != st.session_state.file_order:
                st.session_state.file_order = new_order
                st.rerun()
        else:
            sorted_files = []
            
    # Main Loop - Photo Grid
    sorted_file_objs = []
    
    # Use current state order
    for fname in st.session_state.file_order:
        if fname in st.session_state.managed_files:
            sorted_file_objs.append(st.session_state.managed_files[fname])
            
    # Dynamic column layout
    st.info(f"📸 已載入 {len(sorted_file_objs)} 張照片")

    # Grid Layout - 3 Columns
    for i in range(0, len(sorted_file_objs), 3):
        files_batch = sorted_file_objs[i:i+3]
        cols = st.columns(3)
        
        for j, file in enumerate(files_batch):
            idx = i + j
            with cols[j]:
                with st.container(border=True):
                    # Toolbar Row
                    c_title, c_del = st.columns([5, 1])
                    with c_title:
                        st.markdown(f"**#{idx+1} {file.name}**")
                    with c_del:
                        if st.button("🗑️", key=f"del_{file.name}"):
                            # Delete Logic
                            current_idx = -1
                            if file.name in st.session_state.file_order:
                                current_idx = st.session_state.file_order.index(file.name)
                                st.session_state.file_order.remove(file.name)
                            
                            if file.name in st.session_state.managed_files:
                                del st.session_state.managed_files[file.name]
                                
                            st.session_state.deleted_files.add(file.name)
                            st.session_state.delete_history.append((current_idx, file.name))
                            st.rerun()

                    image = load_image(file)
                    if image:
                        # Image Preview
                        thumb = resize_with_padding(image, target_ratio=14.4/9.8)
                        st.image(thumb, use_container_width=True)
                        
                        unique_key = file.name
                        def_time = datetime.datetime.now().strftime("%H:%M")
                        
                        # Data Inputs
                        c1, c2 = st.columns(2)
                        with c1:
                            date_label = "日期" 
                            date_help = f"預設: {report_date}" if report_date else "未填寫將使用空白"
                            p_date = st.date_input(date_label, value=None, key=f"date_{unique_key}", help=date_help)
                        with c2:
                            p_time = st.text_input("時間", value=def_time, key=f"time_{unique_key}")
                            
                        p_location = st.text_input("📍 地點", value="", placeholder=f"同全域: {location}" if location else "", key=f"loc_{unique_key}")
                        p_desc = st.text_area("📝 說明", value="", placeholder=f"同全域: {global_description}" if global_description else "", key=f"desc_{unique_key}", height=80)
                        
                        photos_data.append({
                            'image': image,
                            'no': f"{idx+1:02d}", 
                            'date': str(p_date) if p_date else (str(report_date) if report_date else ""),
                            'time': p_time,
                            'location': p_location if p_location.strip() else location,
                            'desc': (p_desc if p_desc.strip() else global_description).strip(),
                            'filename': file.name
                        })

    st.markdown("---")
    
    # Generate Button Section
    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        if st.button("🚀 生成 Word 報表", type="primary", use_container_width=True):
            if not photos_data:
                st.warning("⚠️ 請先上傳並保留至少一張照片")
            else:
                with st.spinner("⏳ 正在生成報表，請稍候..."):
                    try:
                        context = {
                            'header_text': report_header,
                            '案由': subject,
                            '地點': location,
                            '製作人': maker,
                            '日期': str(report_date) if report_date else ""
                        }
                        
                        docx_file = create_photo_report(
                            context,
                            photos_data,
                            selected_template,
                            layout_style=layout_style_code
                        )
                        
                        file_name = f"{subject}_{report_date}.docx" if subject and report_date else "現況照片報表.docx"
                        st.success("✅ 報表生成成功！")
                        st.download_button(
                            label="📥 點此下載 Word 檔", 
                            data=docx_file, 
                            file_name=file_name, 
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"❌ 生成失敗: {e}")

# Debug Mode
with st.sidebar:
    st.divider()
    with st.expander("🔧 開發者偵錯模式"):
        st.write("Session State Data:")
        st.json(st.session_state)
        if st.button("🗑️ 清除所有狀態 (Reset)"):
            st.session_state.clear()
            st.rerun()
