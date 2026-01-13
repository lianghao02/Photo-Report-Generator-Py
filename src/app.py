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

# CSS Styling
st.markdown("""
<style>
    .stButton>button {
        width: 100%;
        background-color: #4A90E2;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📷 現況照片清冊生成器 (Python 版)")

with st.sidebar:
    st.header("📋 報表全域設定")
    # Global Fields
    subject = st.text_input("案由 (Project)", value="")
    report_header = st.text_input("頁首標題 (Header)", value="臺南市政府警察局新化分局蒐證照片")
    location = st.text_input("地點 (Location)", value="")
    maker = st.text_input("製作人 (Maker)", value="")
    report_date = st.date_input("日期 (Date)", value=None)
    global_description = st.text_area("全域說明 (Description)", value="")
    
    st.divider()
    # Layout & Template
    st.subheader("📄 模板與排版")
    
    assets_dir = "assets"
    if not os.path.exists(assets_dir):
        os.makedirs(assets_dir)
        
    templates = [f for f in os.listdir(assets_dir) if f.endswith(".docx") and not f.startswith("~$")]
    
    selected_template = None
    layout_style_code = "A4_Vertical" # Default
    
    if templates:
        template_name = st.selectbox("請選擇 Word 模板", templates)
        selected_template = os.path.join(assets_dir, template_name)
        
        # Auto-detect Layout Mode based on filename
        if "左右" in template_name or "Side" in template_name:
            layout_style_code = "A4_SideBySide"
            st.info("ℹ️ 已自動切換為：雙欄對照模式 (Side-by-Side)")
        else:
            layout_style_code = "A4_Vertical"
            st.info("ℹ️ 已自動切換為：直式標準模式 (Vertical)")

        # Analyzer Button
        if st.checkbox("🔍 顯示模板結構分析 (Debug)"):
            if st.button("開始分析"):
                structure = analyze_docx_structure(selected_template)
                st.text_area("分析結果", structure, height=300)
    else:
        st.info(f"無模板 (請將 .docx 放至 {assets_dir})")

# --- Main ---
uploaded_files = st.file_uploader(
    "請選擇或拖曳照片", 
    type=['jpg', 'jpeg', 'png', 'heic'], 
    accept_multiple_files=True
)

if 'photo_details' not in st.session_state:
    st.session_state.photo_details = {}

photos_data = []

if uploaded_files:
    with st.sidebar:
        st.divider()
        st.subheader("🔃 圖片排序 (Drag to Reorder)")
        
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
        # However, we now have a custom delete button. 
        # If user deletes via custom button, we keep it in deleted_files so uploader doesn't re-add it.
        # If user removes from uploader, we also remove from managed_files (done below).
        
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
            # Handle backward compatibility or tuple format
            if isinstance(last_entry, tuple):
                last_idx, last_fname = last_entry
            else:
                last_idx, last_fname = -1, last_entry

            if st.sidebar.button(f"↩️ 復原刪除 ({last_fname})", help="還原上一次刪除的照片"):
                # Undo Logic
                entry_to_restore = st.session_state.delete_history.pop()
                idx_to_restore, fname_to_restore = entry_to_restore if isinstance(entry_to_restore, tuple) else (-1, entry_to_restore)
                
                # 1. Remove from deleted_files
                if fname_to_restore in st.session_state.deleted_files:
                    st.session_state.deleted_files.remove(fname_to_restore)
                
                # 2. Find file object and restore to managed_files AND specific position in file_order
                # We need to find the file object from uploaded_files again
                file_obj = next((f for f in uploaded_files if f.name == fname_to_restore), None)
                if file_obj:
                    st.session_state.managed_files[fname_to_restore] = file_obj
                    
                    if idx_to_restore != -1 and idx_to_restore <= len(st.session_state.file_order):
                        st.session_state.file_order.insert(idx_to_restore, fname_to_restore)
                    else:
                        st.session_state.file_order.append(fname_to_restore) # Fallback to append
                
                st.rerun()

        # Sortable Component
        if st.session_state.file_order:
            # 1. Generate Display List with Auto-Numbering (Index always resets to 1..N)
            # Format: "1. filename.jpg", "2. filename.jpg"
            display_order = [f"{i+1}. {fname}" for i, fname in enumerate(st.session_state.file_order)]
            
            sorted_display = sort_items(display_order, direction='vertical')
            
            # 2. Parse back to clean filename list
            new_order = []
            for item in sorted_display:
                parts = item.split('. ', 1)
                if len(parts) == 2:
                    new_order.append(parts[1])
                else:
                    new_order.append(item) # Fallback
            
            # Update state if changed
            if new_order != st.session_state.file_order:
                st.session_state.file_order = new_order
                st.rerun() # Force rerun to re-generate numbers correctly
        else:
            sorted_files = []
            
    # Main Loop - Iterate over SORTED order
    # Note: We need to reconstruct the batch list from sorted_files
    sorted_file_objs = []
    
    # Use current state order (which is always clean filenames)
    for fname in st.session_state.file_order:
        if fname in st.session_state.managed_files:
            sorted_file_objs.append(st.session_state.managed_files[fname])
            
    # Grid Layout for photos - 3 Columns
    # Process in chunks of 3
    for i in range(0, len(sorted_file_objs), 3):
        files_batch = sorted_file_objs[i:i+3]
        cols = st.columns(3)
        
        for j, file in enumerate(files_batch):
            idx = i + j
            with cols[j]:
                with st.container(border=True): # Add border to frame
                    # Header Column for Title + Delete Button
                    h_col1, h_col2 = st.columns([4, 1])
                    with h_col1:
                        # Show Photo Index + Filename
                        st.subheader(f"照片 {idx+1} - {file.name}")
                    with h_col2:
                        if st.button("🗑️", key=f"del_{file.name}", help="移除此照片"):
                            # Delete Logic
                            # 1. Find current index to restore later
                            current_idx = -1
                            if file.name in st.session_state.file_order:
                                current_idx = st.session_state.file_order.index(file.name)
                                st.session_state.file_order.remove(file.name)
                            
                            if file.name in st.session_state.managed_files:
                                del st.session_state.managed_files[file.name]
                                
                            # Add to ignore list
                            st.session_state.deleted_files.add(file.name)
                            # Add to history (Save INDEX and FILENAME)
                            st.session_state.delete_history.append((current_idx, file.name))
                            st.rerun()

                    image = load_image(file)
                    if image:
                        # Display Image
                        # Use resize_with_padding for consistent height + full image content (Letterboxing)
                        thumb = resize_with_padding(image, target_ratio=14.4/9.8)
                        st.image(thumb, use_container_width=True)
                        
                        unique_key = file.name
                        
                        # Defaults
                        def_time = datetime.datetime.now().strftime("%H:%M")
                        
                        # Inputs - Vertical Stack
                        # Use unique_key (filename) for widget keys to persist data across reorders
                        
                        # Row 1: Date | Time
                        c1, c2 = st.columns(2)
                        with c1:
                            # Dynamic Label to hint default
                            date_label = f"日期 (預設: {report_date})" if report_date else "日期 (請選擇)"
                            p_date = st.date_input(date_label, value=None, key=f"date_{unique_key}")
                        with c2:
                            p_time = st.text_input("時間", value=def_time, key=f"time_{unique_key}")
                            
                        # Row 2: Location
                        # Default to global location, but allow override
                        p_location = st.text_input("地點", value="", placeholder=f"預設: {location}" if location else "請輸入地點", key=f"loc_{unique_key}")
                        
                        # Row 3: Description
                        p_desc = st.text_area("說明", value="", placeholder=f"預設: {global_description}" if global_description else "請輸入說明", key=f"desc_{unique_key}", height=100)
                        
                        photos_data.append({
                            'image': image,
                            'no': f"{idx+1:02d}", 
                            'date': str(p_date) if p_date else (str(report_date) if report_date else ""),
                            'time': p_time,
                            # Fallback Logic: If local input is empty, use global location
                            'location': p_location if p_location.strip() else location,
                            # Fallback Logic: If local input is empty, use global description
                            'desc': (p_desc if p_desc.strip() else global_description).strip(),
                            'filename': file.name
                        })

    st.divider()
    if st.button("🚀 生成 Word 報表", type="primary"):
        if not photos_data:
            st.warning("請先上傳照片")
        else:
            with st.spinner("生成中..."):
                try:
                    # Prepare Global Context
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
                    
                    file_name = f"{subject}_{report_date}.docx"
                    st.success("✅ 成功！")
                    st.download_button("📥 下載", docx_file, file_name, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                except Exception as e:
                    st.error(f"生成失敗: {e}")
