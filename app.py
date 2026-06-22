import streamlit as st
import google.generativeai as genai
import fitz          # PyMuPDF
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import time
import base64
import re

# ── 頁面設定 ──────────────────────────────────────────────
st.set_page_config(page_title="未來線中運量機電特別技術規範 - 智慧改寫平台", layout="wide")

for key, val in [
    ("running", False),
    ("result_text", ""),
    ("extracted_old_text", ""),
    ("cf620_total_pages", 0),
    ("cf620_toc", []),
    ("cf620_toc_raw", []),          # 邏輯頁碼（未套偏移）
    ("cf620_detected_offset", 0),   # Gemini 自動偵測的偏移
    ("pdf_bytes_cache", None),
    ("chapter_id", ""),
    ("next_chapter_id", ""),
    ("cf620_pdf_name", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = val

def set_running():
    st.session_state.running = True

# ══════════════════════════════════════════════════════════
# 工具函式
# ══════════════════════════════════════════════════════════

def detect_toc_pages(file_bytes, max_scan=20):
    """偵測前 max_scan 頁中目錄頁（章節編號密度高的頁面），回傳 page_no list（0-based）"""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    toc_pat = re.compile(r"^\d{1,2}(?:\.\d{1,3}){1,4}")
    toc_pages = []
    for page_no in range(min(len(doc), max_scan)):
        lines = doc[page_no].get_text("text").splitlines()
        hits = sum(1 for ln in lines if toc_pat.match(ln.strip()))
        if hits >= 5:
            toc_pages.append(page_no)
    return toc_pages, len(doc)


def render_page_to_png(file_bytes, page_no, dpi=150):
    """將 PDF 指定頁（0-based）render 成 PNG bytes"""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    page = doc[page_no]
    mat = fitz.Matrix(dpi / 72, dpi / 72)
    pix = page.get_pixmap(matrix=mat)
    return pix.tobytes("png")


def detect_page_offset(api_key, model_name, file_bytes, toc_page_nos, first_logical_page):
    """
    計算 PDF 物理頁碼與目錄所列邏輯頁碼的偏移量。
    作法：取目錄後第一頁 render 成圖，問 Gemini「這頁的頁碼是幾」，
    offset = (pdf_index + 1) - gemini_回答頁碼
    first_logical_page：目錄中最小的頁碼數字（作為預期值）
    """
    if not toc_page_nos:
        return 0
    first_content_pdf_idx = max(toc_page_nos) + 1   # TOC 最後一頁的下一頁
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    if first_content_pdf_idx >= len(doc):
        return 0

    # 嘗試往後幾頁找到有明確頁碼的頁面
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    for probe_idx in range(first_content_pdf_idx, min(first_content_pdf_idx + 5, len(doc))):
        png_bytes = render_page_to_png(file_bytes, probe_idx)
        b64 = base64.b64encode(png_bytes).decode("utf-8")
        prompt_parts = [
            {"inline_data": {"mime_type": "image/png", "data": b64}},
            {"text": (
                "這是一份技術文件的某一頁。"
                "請只回答這頁底部或頁首所顯示的頁碼數字（阿拉伯數字）。"
                "若看不到頁碼，請回答 UNKNOWN。"
                "只輸出一個數字或 UNKNOWN，不要其他文字。"
            )}
        ]
        try:
            resp = model.generate_content([{"parts": prompt_parts}])
            ans = resp.text.strip()
            if ans.isdigit():
                printed = int(ans)
                offset = (probe_idx + 1) - printed   # PDF 1-based - 印刷頁碼
                return offset
        except Exception:
            pass
    return 0


def build_toc_via_vision(api_key, model_name, file_bytes):
    """
    主流程：
      1. 偵測目錄頁
      2. 把每頁 render 成圖送 Gemini Vision → 解析邏輯頁碼目錄
      3. 偵測 PDF 物理頁碼偏移，轉換為絕對 PDF 頁碼
    回傳 (toc, total_pages)
      toc = [[level, "X.X.X 標題", absolute_pdf_page], ...]
    """
    toc_page_nos, total = detect_toc_pages(file_bytes)

    # fallback：找不到目錄頁就掃前 5 頁
    if not toc_page_nos:
        toc_page_nos = list(range(min(5, total)))

    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    raw_lines = []
    for page_no in toc_page_nos:
        png_bytes = render_page_to_png(file_bytes, page_no)
        b64 = base64.b64encode(png_bytes).decode("utf-8")
        prompt_parts = [
            {"inline_data": {"mime_type": "image/png", "data": b64}},
            {"text": (
                "這是一份技術規範文件的目錄頁圖片。"
                "請擷取所有目錄條目，每行輸出格式為：\n"
                "章節編號 章節標題 頁碼\n"
                "例如：2.1.12 空調系統與火災警報設備 84\n"
                "只輸出條目清單，不要輸出任何說明或前言。"
                "若圖片中沒有目錄，請輸出：NO_TOC"
            )}
        ]
        try:
            resp = model.generate_content([{"parts": prompt_parts}])
            raw_lines.extend(resp.text.strip().splitlines())
        except Exception:
            pass

    # 解析 Gemini 回傳的文字行
    toc_raw = []
    seen = set()
    pat = re.compile(
        r"^(\d{1,2}(?:\.\d{1,3}){1,5}|第\s*[\d一二三四五六七八九十百]+\s*[章節條篇])"
        r"\s+(.+?)\s+(\d{1,4})\s*$"
    )
    for ln in raw_lines:
        ln = ln.strip()
        if not ln or ln == "NO_TOC":
            continue
        ln_clean = re.sub(r"[.·‥…\s]{4,}", " ", ln).strip()
        m = pat.match(ln_clean)
        if not m:
            continue
        sec_id    = m.group(1).strip()
        title     = m.group(2).strip()
        logical_p = int(m.group(3))
        full_title = f"{sec_id} {title}"
        if full_title in seen:
            continue
        seen.add(full_title)
        dots  = sec_id.count(".")
        level = min(dots + 1, 3)
        toc_raw.append([level, full_title, logical_p])

    # ── 計算偏移，將邏輯頁碼轉為絕對 PDF 頁碼 ──────────────
    first_logical = toc_raw[0][2] if toc_raw else 1
    offset = detect_page_offset(api_key, model_name, file_bytes, toc_page_nos, first_logical)

    toc = [
        [level, title, max(1, logical_p + offset)]
        for level, title, logical_p in toc_raw
    ]

    return toc, toc_raw, offset, total


def extract_pages(file_bytes, page_start, page_end):
    """萃取指定頁碼範圍（1-based），保留表格結構，過濾頁首頁尾"""
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    total = len(doc)
    p0 = max(0, page_start - 1)
    p1 = min(total - 1, page_end - 1)
    all_text = ""

    # 頁首頁尾過濾：匹配 "CF620/ 97"、"DF115/ 3" 或純數字頁碼
    hf_pat = re.compile(r'^[A-Z]{2}\d{3}/\s*\d+\s*$|^\d{1,4}\s*$')

    for i in range(p0, p1 + 1):
        page = doc[i]
        all_text += f"\n\n--- 第 {i+1} 頁 ---\n"
        try:
            tables = page.find_tables()
            if tables.tables:
                # 取得所有表格的矩形範圍，避免文字與表格重複輸出
                table_rects = [fitz.Rect(tbl.bbox) for tbl in tables.tables]

                # 先輸出不在表格區域內的文字 blocks（同時過濾頁首頁尾）
                for b in page.get_text("blocks"):
                    if b[6] != 0:          # 略過圖片 block
                        continue
                    b_rect = fitz.Rect(b[:4])
                    if any(b_rect.intersects(tr) for tr in table_rects):
                        continue           # 屬於表格區域，跳過（由下方 markdown 輸出）
                    for ln in b[4].strip().splitlines():
                        if not hf_pat.match(ln.strip()):
                            all_text += ln + "\n"

                # 再輸出表格 Markdown（完整保留，包含最後一頁表格）
                for tbl in tables.tables:
                    df = tbl.to_pandas()
                    all_text += "\n[表格]\n" + df.to_markdown(index=False) + "\n"
            else:
                # 無表格：逐 block 輸出，過濾頁首頁尾
                for b in page.get_text("blocks"):
                    if b[6] != 0:
                        continue
                    for ln in b[4].strip().splitlines():
                        if not hf_pat.match(ln.strip()):
                            all_text += ln + "\n"
        except Exception:
            for b in page.get_text("blocks"):
                all_text += b[4].strip() + "\n"
    return all_text, total


def trim_to_chapter(text, chapter_id, next_chapter_id=None):
    """
    從萃取的多頁文字中，精確裁切出目標章節的內容：
    - 往前裁：找到 chapter_id 開頭的那一行，丟棄之前的內容
    - 往後裁：找到 next_chapter_id 開頭的那一行，丟棄之後的內容
    chapter_id 例如 "2.1.12"，允許後面接空格或中文
    """
    if not chapter_id:
        return text

    lines = text.splitlines()

    def make_pat(sec_id):
        # 逸出點號，允許編號後接空格或中文字
        escaped = re.escape(sec_id)
        return re.compile(r"^\s*" + escaped + r"(\s|\u4e00-\u9fff|$)")

    start_pat = make_pat(chapter_id)
    start_idx = None
    for i, ln in enumerate(lines):
        if start_pat.match(ln):
            start_idx = i
            break

    if start_idx is None:
        # 找不到精確起點就原文回傳
        return text

    end_idx = len(lines)
    if next_chapter_id:
        end_pat = make_pat(next_chapter_id)
        for i in range(start_idx + 1, len(lines)):
            if end_pat.match(lines[i]):
                end_idx = i
                break

    return "\n".join(lines[start_idx:end_idx]).strip()


def extract_pdf_with_tables(file_bytes):
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    all_text = ""
    for page in doc:
        try:
            tables = page.find_tables()
            page_text = page.get_text("text")
            if tables.tables:
                for tbl in tables.tables:
                    all_text += "\n[表格]\n" + tbl.to_pandas().to_markdown(index=False) + "\n"
            all_text += page_text + "\n"
        except Exception:
            all_text += page.get_text() + "\n"
    return all_text


def image_to_base64(file_bytes):
    return base64.b64encode(file_bytes).decode("utf-8")


def extract_text_from_image(api_key, model_name, img_bytes, mime_type):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    response = model.generate_content([{
        "parts": [
            {"inline_data": {"mime_type": mime_type, "data": image_to_base64(img_bytes)}},
            {"text": "請完整辨識圖片中所有文字與表格，表格以 Markdown 格式（| 欄 | 欄 |）輸出，保留所有數字、單位與編號。"}
        ]
    }])
    return response.text


def result_to_docx(result_text):
    """
    將 AI 改寫結果（=== 區塊格式）轉換為 Word .docx bytes。
    支援：章節標題、一般段落、Markdown 表格、▸/✕/💡 條列。
    """
    doc = docx.Document()

    # ── 全域字體 ────────────────────────────────────────────
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Arial'
    normal_style.font.size = Pt(11)

    # ── 標題樣式 ────────────────────────────────────────────
    for level, pt, bold in [(1, 14, True), (2, 12, True)]:
        h = doc.styles[f'Heading {level}']
        h.font.name = 'Arial'
        h.font.size = Pt(pt)
        h.font.bold = bold
        h.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)   # 深藍

    # ── 解析 === 區塊 ────────────────────────────────────────
    sections = result_text.split("===")
    for i in range(1, len(sections) - 1, 2):
        title   = sections[i].strip()
        content = sections[i + 1].strip() if i + 1 < len(sections) else ""

        doc.add_heading(title, level=1)

        lines = content.split("\n")
        j = 0
        while j < len(lines):
            line    = lines[j]
            stripped = line.strip()

            # ── Markdown 表格 ────────────────────────────────
            if stripped.startswith("|"):
                tbl_rows = []
                while j < len(lines) and (
                    lines[j].strip().startswith("|") or
                    (not lines[j].strip() and j + 1 < len(lines) and
                     lines[j + 1].strip().startswith("|"))
                ):
                    if lines[j].strip():
                        tbl_rows.append(lines[j].strip())
                    j += 1

                # 過濾分隔行（| --- | --- |）
                data_rows = [r for r in tbl_rows
                             if not re.match(r"^\|[-:\s|]+\|$", r)]

                if data_rows:
                    num_cols = len(data_rows[0].strip().strip("|").split("|"))
                    tbl = doc.add_table(rows=0, cols=num_cols)
                    tbl.style = 'Table Grid'

                    for row_idx, row_text in enumerate(data_rows):
                        cells = [c.strip() for c in
                                 row_text.strip().strip("|").split("|")]
                        row = tbl.add_row()
                        for col_idx, cell_text in enumerate(cells[:num_cols]):
                            cell = row.cells[col_idx]
                            cell.text = cell_text
                            if row_idx == 0:   # 表頭加粗
                                for para in cell.paragraphs:
                                    for run in para.runs:
                                        run.bold = True

                    doc.add_paragraph()   # 表格後空行
                continue   # j 已在內層迴圈推進

            # ── 條列符號行（▸ ✕ 💡）────────────────────────
            elif stripped and stripped[0] in ("▸", "✕", "💡"):
                p = doc.add_paragraph(stripped, style='Normal')
                p.paragraph_format.left_indent = Inches(0.3)
                p.paragraph_format.space_after  = Pt(4)

            # ── 小標（以數字+點+空白開頭，如「1. 」「A. 」）──
            elif re.match(r"^\d{1,2}[\.、]\s|^[A-Z]\.\s", stripped) and stripped:
                p = doc.add_paragraph(stripped, style='Normal')
                p.paragraph_format.left_indent = Inches(0.2)

            # ── 一般文字段落 ─────────────────────────────────
            elif stripped:
                doc.add_paragraph(stripped, style='Normal')

            # ── 空行（視為段落間距，不重複加）────────────────
            else:
                pass   # spacing_after 已由段落格式處理

            j += 1

        # 章節之間加分隔線（水平段落框線）
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after  = Pt(4)
        from docx.oxml.ns import qn
        from docx.oxml    import OxmlElement
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), 'BBBBBB')
        pBdr.append(bottom)
        pPr.append(pBdr)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════
# 標題
# ══════════════════════════════════════════════════════════
st.title("🚆 未來線中運量機電特別技術規範 - 智慧改寫平台")

# ══════════════════════════════════════════════════════════
# 左側欄
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.header("⚙️ 系統設定")
    api_key = st.text_input("🔑 輸入 Gemini API Key", type="password",
                             help="https://aistudio.google.com/app/apikey")
    selected_model = st.selectbox(
        "🤖 選擇模型",
        ["gemini-3.1-flash-lite", "gemini-3.5-flash"],
        index=0,
        help="3.1-flash-lite：速度快、省配額，適合一般改寫。\n3.5-flash：細節保留較完整，建議用於重要章節。"
    )

    st.markdown("---")
    st.header("📂 匯入精進文件")
    st.caption("上傳未來線最新測試規定（PDF / Word）")
    uploaded_files = st.file_uploader("選擇檔案（可多選）",
        accept_multiple_files=True, type=["pdf", "docx"], key="kb_uploader")

    st.markdown("---")
    st.header("✏️ 額外提示語（選填）")
    st.caption("可輸入改寫方向、特殊要求、精進文字（不上傳文件時直接貼上亦可），AI 改寫時將一併參考")
    user_hint = st.text_area(
        "提示語",
        height=300,
        placeholder="例如：\n・請特別注意電力系統改用 750V DC\n・保留所有測試步驟數值\n・或直接貼上精進文件全文（約 1000 字以內效果最佳）…",
        label_visibility="collapsed"
    )

kb_text = ""
if uploaded_files:
    for f in uploaded_files:
        try:
            if f.name.lower().endswith(".pdf"):
                kb_text += extract_pdf_with_tables(f.read())
            elif f.name.lower().endswith(".docx"):
                wd = docx.Document(io.BytesIO(f.read()))
                for para in wd.paragraphs:
                    kb_text += para.text + "\n"
        except Exception as e:
            st.sidebar.warning(f"⚠️ 無法讀取 {f.name}：{e}")
    st.sidebar.success(f"✅ 成功載入 {len(uploaded_files)} 份精進文件！")

# ══════════════════════════════════════════════════════════
# 主畫面
# ══════════════════════════════════════════════════════════
st.markdown("---")
col1, col2 = st.columns(2)

with col1:
    st.subheader("📝 ① 上傳 預修改之PTS")

    tab_pdf, tab_paste, tab_img = st.tabs([
        "📄 上傳 PDF（推薦）",
        "✍️ 直接貼上文字",
        "🖼️ 截圖辨識",
    ])

    # ── Tab 1：大型 PDF 選頁 ──────────────────────────────
    with tab_pdf:
        st.caption("💡 使用流程：上傳 PDF → 選擇章節 → 按「萃取內容」，即可自動擷取文字")

        cf620_pdf = st.file_uploader(
            "上傳 PTS 完整 PDF", type=["pdf"], key="cf620_big_pdf"
        )

        if cf620_pdf:
            # ── 只在「新上傳」時重新偵測，其餘 rerun 直接用快取 ──
            if cf620_pdf.name != st.session_state.cf620_pdf_name:
                pdf_bytes = cf620_pdf.read()
                st.session_state.pdf_bytes_cache = pdf_bytes
                st.session_state.cf620_pdf_name = cf620_pdf.name

                with st.spinner("📖 使用 Gemini Vision 辨識目錄中，請稍候…"):
                    if not api_key:
                        st.error("⚠️ 請先輸入 API Key 才能辨識目錄！")
                        toc, toc_raw, detected_offset, total_pages = [], [], 0, 1
                    else:
                        toc, toc_raw, detected_offset, total_pages = build_toc_via_vision(
                            api_key, selected_model, pdf_bytes)
                    st.session_state.cf620_total_pages  = total_pages
                    st.session_state.cf620_toc          = toc
                    st.session_state.cf620_toc_raw      = toc_raw
                    st.session_state.cf620_detected_offset = detected_offset
            else:
                # 同一份 PDF，直接從快取載入
                pdf_bytes       = st.session_state.pdf_bytes_cache
                toc             = st.session_state.cf620_toc
                toc_raw         = st.session_state.cf620_toc_raw
                detected_offset = st.session_state.cf620_detected_offset
                total_pages     = st.session_state.cf620_total_pages

            st.info(f"📖 已載入：共 **{total_pages}** 頁　｜　偵測到 **{len(toc)}** 個章節")

            # ── 頁碼校正：預設摺疊，進階用戶才需要展開 ────────
            if toc_raw:
                with st.expander("⚙️ 頁碼有偏差？展開校正（選用）", expanded=False):
                    st.caption("選一個你確定的章節，輸入它在 PDF 中實際的頁碼，程式會自動修正所有章節頁碼")
                    ref_options = [
                        f"第 {item[2]} 頁 | {item[1]}"
                        for item in toc
                    ]
                    col_ref1, col_ref2 = st.columns([2, 1])
                    with col_ref1:
                        selected_ref = st.selectbox(
                            "選擇參考章節",
                            ["— 不校正 —"] + ref_options,
                            key="ref_select",
                        )
                    with col_ref2:
                        if selected_ref != "— 不校正 —":
                            actual_page = st.number_input(
                                "實際頁碼",
                                min_value=1, max_value=total_pages,
                                value=int(selected_ref.split()[1]),
                                key="actual_page",
                            )
                        else:
                            actual_page = None

                    # 根據參考點重算偏移
                    if selected_ref != "— 不校正 —" and actual_page is not None:
                        ref_idx = ref_options.index(selected_ref)
                        correct_offset = actual_page - toc_raw[ref_idx][2]
                        st.success(f"✅ 校正完成！偏移 {correct_offset:+d} 頁")
                        toc = [
                            [lv, t, max(1, lp + correct_offset)]
                            for lv, t, lp in toc_raw
                        ]

            if toc:
                toc_options = [
                    f"第{item[2]}頁｜{'　' * (item[0]-1)}{item[1]}"
                    for item in toc
                ]

                # ── 預設：單一章節選擇（selectbox 內建搜尋）────
                selected_toc = st.selectbox(
                    f"選擇章節（共 {len(toc)} 筆，可直接輸入關鍵字搜尋）",
                    ["— 請選擇 —"] + toc_options,
                    key="toc_select",
                )

                # ── 跨章節模式：checkbox 預設隱藏 ─────────
                multi_mode = st.checkbox("需要跨多個章節？", key="multi_mode")

                if not multi_mode:
                    # ══ 單一章節 ══════════════════════════
                    if selected_toc != "— 請選擇 —":
                        chosen_idx = toc_options.index(selected_toc)
                        auto_start = toc[chosen_idx][2]
                        auto_end   = total_pages
                        _next_id   = ""
                        for j in range(chosen_idx + 1, len(toc)):
                            if toc[j][0] <= toc[chosen_idx][0]:
                                auto_end = toc[j][2] - 1
                                _next_id = toc[j][1].split()[0]
                                break
                        auto_end = max(auto_end, auto_start)
                        auto_end = min(auto_end, auto_start + 29)  # 最多 30 頁
                        st.session_state.chapter_id      = toc[chosen_idx][1].split()[0]
                        st.session_state.next_chapter_id = _next_id
                    else:
                        auto_start, auto_end = 1, min(10, total_pages)
                        st.session_state.chapter_id = ""
                        st.session_state.next_chapter_id = ""

                else:
                    # ══ 跨章節範圍 ════════════════════════
                    col_s, col_e = st.columns(2)
                    with col_s:
                        start_sel = st.selectbox(
                            f"起始章節（共 {len(toc)} 筆）",
                            ["— 請選擇 —"] + toc_options,
                            key="range_start",
                        )
                    with col_e:
                        end_sel = st.selectbox(
                            "結束章節",
                            ["— 請選擇 —"] + toc_options,
                            key="range_end",
                        )

                        if start_sel != "— 請選擇 —" and end_sel != "— 請選擇 —":
                            si = toc_options.index(start_sel)
                            ei = toc_options.index(end_sel)

                            if ei < si:
                                st.error("❌ 結束章節不能在起始章節之前，請重新選擇。")
                                auto_start, auto_end = 1, min(10, total_pages)
                                st.session_state.chapter_id = ""
                                st.session_state.next_chapter_id = ""
                            else:
                                auto_start = toc[si][2]
                                auto_end  = total_pages
                                _next_id  = ""
                                for j in range(ei + 1, len(toc)):
                                    if toc[j][0] <= toc[ei][0]:
                                        auto_end = toc[j][2] - 1
                                        _next_id = toc[j][1].split()[0]
                                        break
                                auto_end = max(auto_end, auto_start)
                                page_span = auto_end - auto_start + 1
                                if page_span > 30:
                                    st.info(
                                        f"ℹ️ 已選 **{toc[si][1]}** → **{toc[ei][1]}**，"
                                        f"共約 **{page_span}** 頁。頁數較多時 AI 改寫可能需要較長時間。"
                                    )
                                st.session_state.chapter_id      = toc[si][1].split()[0]
                                st.session_state.next_chapter_id = _next_id
                        else:
                            auto_start, auto_end = 1, min(10, total_pages)
                            st.session_state.chapter_id = ""
                            st.session_state.next_chapter_id = ""

            else:
                st.warning("⚠️ 未偵測到章節標題，請手動輸入頁碼。")
                auto_start, auto_end = 1, min(10, total_pages)

            # ── 頁碼範圍：自動計算結果直接顯示，進階用戶可展開微調 ──
            page_count_preview = min(auto_end, total_pages) - auto_start + 1
            if page_count_preview > 30:
                st.warning(f"⚠️ 預計萃取 {page_count_preview} 頁，建議不超過 30 頁。")
            else:
                st.success(f"✅ 即將萃取第 {auto_start} ～ {min(auto_end, total_pages)} 頁（共 {page_count_preview} 頁）")

            # 預設值（expander 未展開時使用）
            page_start = auto_start
            page_end   = min(auto_end, total_pages)

            with st.expander("🔧 手動調整頁碼範圍（選用）", expanded=False):
                st.caption("若自動計算的頁碼範圍不正確，可在此手動修改")
                c_s, c_e = st.columns(2)
                with c_s:
                    page_start = st.number_input("起始頁", 1, total_pages, auto_start)
                with c_e:
                    page_end = st.number_input("結束頁", 1, total_pages,
                                               min(auto_end, total_pages))
                page_count = page_end - page_start + 1
                if page_count > 30:
                    st.warning(f"⚠️ 選取 {page_count} 頁，建議不超過 30 頁。")

            if st.button("📥 萃取選定頁面內容", use_container_width=True):
                with st.spinner(f"萃取第 {page_start}～{page_end} 頁..."):
                    try:
                        extracted, _ = extract_pages(pdf_bytes, page_start, page_end)
                        # ── 精確裁切：去除前後相鄰章節的殘留內容 ──
                        ch_id   = st.session_state.get("chapter_id", "")
                        next_id = st.session_state.get("next_chapter_id", "")
                        if ch_id:
                            extracted = trim_to_chapter(extracted, ch_id, next_id or None)
                        st.session_state.extracted_old_text = extracted
                        st.success(f"✅ 萃取完成！共 {len(extracted)} 字元")
                        with st.expander("預覽萃取內容", expanded=False):
                            st.text(extracted[:3000] + ("..." if len(extracted) > 3000 else ""))
                    except Exception as e:
                        st.error(f"❌ 萃取失敗：{e}")

    # ── Tab 2：手動貼上 ──────────────────────────────────
    with tab_paste:
        pasted_text = st.text_area(
            "請複製舊規範條文貼在此處：",
            height=380,
            placeholder="貼上純文字條文，若有表格無法複製請改用其他分頁..."
        )

    # ── Tab 3：截圖辨識 ──────────────────────────────────
    with tab_img:
        st.caption("📷 截圖後上傳，由 Gemini 辨識圖片中的表格")
        cf620_img = st.file_uploader("上傳截圖（PNG / JPG）",
            type=["png", "jpg", "jpeg"], key="cf620_img")
        if cf620_img:
            st.image(cf620_img, use_container_width=True)
            if st.button("🔍 執行圖片辨識", use_container_width=True):
                if not api_key:
                    st.error("⚠️ 請先輸入 API Key！")
                else:
                    with st.spinner("Gemini 辨識中..."):
                        try:
                            mime = "image/png" if cf620_img.name.lower().endswith(".png") else "image/jpeg"
                            result = extract_text_from_image(api_key, selected_model, cf620_img.read(), mime)
                            st.session_state.extracted_old_text = result
                            st.success("✅ 辨識完成！")
                            with st.expander("預覽辨識結果", expanded=True):
                                st.markdown(result)
                        except Exception as e:
                            st.error(f"❌ 辨識失敗：{e}")

    # 整合輸入來源
    if st.session_state.extracted_old_text:
        old_text = st.session_state.extracted_old_text
        st.info("✅ 目前使用：PDF 頁面萃取 / 圖片辨識結果")
    else:
        old_text = pasted_text if "pasted_text" in dir() else ""

    st.markdown("---")
    run_btn = st.button("🚀 執行 AI 改寫",
        disabled=st.session_state.running,
        use_container_width=True, type="primary")
    clear_btn = st.button("✏️ 清除結果", use_container_width=True)

# ══════════════════════════════════════════════════════════
# 右欄：輸出
# ══════════════════════════════════════════════════════════
with col2:
    st.subheader("✨ ② 智慧改寫草稿")

    if clear_btn:
        st.session_state.result_text = ""
        st.session_state.extracted_old_text = ""
        st.rerun()

    if run_btn:
        if not api_key:
            st.error("⚠️ 請先輸入 Gemini API Key！")
        elif not old_text or not old_text.strip():
            if st.session_state.pdf_bytes_cache:
                st.error("⚠️ 請先點擊「📥 萃取選定頁面內容」按鈕，再執行 AI 改寫！")
            else:
                st.error("⚠️ 請先提供待改寫條文（上傳 PDF 並萃取、貼上文字，或上傳截圖辨識）。")
        else:
            st.session_state.running = True
            hint_section = (
                f"\n\n【使用者額外提示語（請優先遵照執行）】\n{user_hint.strip()}"
                if user_hint and user_hint.strip() else ""
            )
            kb_section = (
                f"\n\n【精進文件內容（請優先採用以下參數）】\n{kb_text}"
                if kb_text else "\n\n（本次未上傳精進文件，請依中運量通用規範改寫）"
            )
            prompt = f"""你是一位具備捷運機電系統工程與合約撰寫背景的資深專家。

【任務】
將下方『預修改之PTS（特別技術規範）』改寫為『未來線中運量系統』版本。

【改寫原則——請嚴格遵守，違反任何一條均為失敗】
1. 【範圍限制】你的改寫範圍僅限於「待改寫條文」區塊中所提供的文字，絕對不得自行補充、延伸或改寫未提供的章節內容。若待改寫條文的開頭有前文接續文字（如「（接續前頁）」），請直接從第一個完整條款開始改寫。
2. 嚴格保留原有章節編號與排版格式。
3. 【禁止省略】原文每一條、每一款、每一數值（溫度、面積、載重、時間等）都必須出現在改寫結果中，不得以「如原文」、「同上」或省略號代替。
4. 【表格必須完整保留】若原文有表格，改寫後必須以 Markdown 表格輸出，且嚴守以下格式規定：
   - 表格標題行（| 欄1 | 欄2 | 欄3 |）與分隔線（| --- | --- | --- |）必須完整輸出
   - 每一個測試項目（如 A.車廂地板防火測試、B.車間走道防火測試）都必須是表格中獨立的一列（row），不得拆到表格外
   - 說明欄中若有多個子條件（1. 2. 3.…），全部寫在同一格內，子條件之間用「；」分隔，不得換行成表格外的條列
   - 禁止在表格行之間插入任何非表格文字或空行
5. 優先採用精進文件中的最新中運量規格參數；若未提及，則保留原始數值並標註〔待確認〕。
6. 不適用未來線中運量的設備或條款，保留完整條文內容並在結尾標註【建議刪除，或由機設處重新評估】。
7. 每一條款之間必須空一行，保持段落清晰易讀；禁止將所有條文連成一段文字輸出。
8. 【禁止使用 HTML 標籤】輸出中禁止使用 <br>、<br/>、<p> 等任何 HTML 標籤，段落換行只能用空行（換兩次 Enter）表示。
9. 請勿輸出任何問候語，直接輸出以下四個區塊。{hint_section}{kb_section}

【待改寫條文】
{old_text}

---
請依以下格式輸出，每個區塊用 === 分隔：

=== 一、改寫後條文 ===
（完整改寫版條文，保留編號格式，各條款間空一行，表格用 Markdown 呈現）

=== 二、新增內容摘要 ===
（每項以「▸」開頭，各項間空一行）

=== 三、刪除 / 調整內容摘要 ===
（每項以「✕」開頭，各項間空一行）

=== 四、專業意見與注意事項 ===
（每項以「💡」開頭，各項間空一行）
"""
            with st.spinner(f"⏳ 使用 {selected_model} 改寫中..."):
                retries, success = 3, False
                while not success and retries > 0:
                    try:
                        genai.configure(api_key=api_key)
                        model = genai.GenerativeModel(selected_model)
                        response = model.generate_content(prompt)
                        st.session_state.result_text = response.text
                        success = True
                    except Exception as e:
                        err = str(e)
                        if "429" in err or "quota" in err.lower():
                            retries -= 1
                            if retries > 0:
                                st.warning(f"⚠️ 配額繁忙，15 秒後重試（剩餘 {retries} 次）...")
                                time.sleep(15)
                        elif "API_KEY_INVALID" in err or "api key" in err.lower():
                            st.error("❌ API Key 無效。"); break
                        else:
                            st.error(f"❌ 錯誤：{err}"); break
                if not success and retries == 0:
                    st.error("❌ 已超過重試次數，請稍後手動重試。")

            st.session_state.running = False
            st.rerun()

    if st.session_state.result_text:
        result_text = st.session_state.result_text
        sections = result_text.split("===")
        parsed = {}
        for i in range(1, len(sections) - 1, 2):
            title = sections[i].strip()
            content = sections[i + 1].strip() if i + 1 < len(sections) else ""
            parsed[title] = content

        emoji_map = {"一": "📄", "二": "🆕", "三": "🗑️", "四": "💡"}
        if parsed:
            for title, content in parsed.items():
                emoji = emoji_map.get(title[0] if title else "", "📌")
                with st.expander(f"{emoji} {title}", expanded=title.startswith("一")):
                    # ── 渲染處理：Markdown 表格 → HTML 表格 + 段落換行 ────
                    # Markdown 表格儲存格不支援換行，改用 HTML 表格確保條列正確顯示

                    def parse_md_table_to_html(md_lines):
                        """將連續的 Markdown 表格行轉換為 HTML 表格（支援儲存格內換行）"""
                        html = ['<table border="1" style="border-collapse:collapse;width:100%;font-size:0.9em;">']
                        is_header = True
                        for row in md_lines:
                            # 跳過分隔線
                            if re.match(r"^\|[-:\s|]+\|$", row.strip()):
                                is_header = False
                                continue
                            # 解析儲存格（去掉首尾 |，再 split）
                            cells = [c.strip() for c in row.strip().strip("|").split("|")]
                            tag = "th" if is_header else "td"
                            style = "padding:6px 10px;vertical-align:top;border:1px solid #ccc;"
                            if is_header:
                                style += "background:#f0f0f0;font-weight:bold;"
                            row_html = "<tr>" + "".join(
                                f'<{tag} style="{style}">{format_cell(c)}</{tag}>'
                                for c in cells
                            ) + "</tr>"
                            html.append(row_html)
                        html.append("</table>")
                        return "\n".join(html)

                    def format_cell(text):
                        """
                        儲存格內容格式化：
                        - 修正「1. xxx；2. yyy；3. zzz」時，原本第 1 點被當成前言，
                          導致第 2 點在畫面上變成清單第 1 點的問題。
                        - 將所有連續數字條列轉成同一個 <ol>，並保留原本第一個數字作為 start。
                        """
                        import html as html_mod

                        if text is None:
                            return ""

                        text = str(text)
                        text = text.replace("<br>", "\n").replace("<br/>", "\n").replace("<br />", "\n")
                        text = text.replace("\r\n", "\n").replace("\r", "\n")
                        text = re.sub(r"\s*；\s*", "；", text)

                        # 找出條列編號，例如：1.、2.、3、，但排除 2.1、3.2.1 這類章節或小數。
                        # 舊版 bug 是先用第一個分號切前言，導致「1. ...；2. ...」的第 1 點被當成前言，
                        # 畫面上只剩第 2 點進入 <ol>，因此第 2 點會顯示成清單第 1 點。
                        item_pat = re.compile(r"(?<![\d.])(\d{1,2})[.、]\s*(?!\d)")
                        matches = list(item_pat.finditer(text))

                        if matches:
                            preamble = text[:matches[0].start()].strip("；; \n\t")
                            items = []

                            for idx, match in enumerate(matches):
                                item_start = match.end()
                                item_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
                                item_text = text[item_start:item_end].strip("；; \n\t")

                                if item_text:
                                    items.append(item_text)

                            if items:
                                first_number = matches[0].group(1)
                                prefix = (
                                    f"<p style='margin:0 0 4px 0;'>{html_mod.escape(preamble)}</p>"
                                    if preamble else ""
                                )
                                li_items = "".join(
                                    f"<li>{html_mod.escape(item).replace(chr(10), '<br>')}</li>"
                                    for item in items
                                )
                                return (
                                    f"{prefix}"
                                    f"<ol start='{first_number}' style='margin:0;padding-left:1.4em;'>"
                                    f"{li_items}"
                                    f"</ol>"
                                )

                        return html_mod.escape(text).replace("\n", "<br>")


                    # 逐行分組：表格行 vs 一般段落行
                    lines = content.split("\n")
                    output_parts = []   # 每個元素為 ("md", text) 或 ("table", [行列表])
                    i = 0
                    while i < len(lines):
                        line = lines[i]
                        stripped = line.strip()
                        if stripped.startswith("|"):
                            # 收集完整表格（包含後面緊接的空白行跳過，直到下一個 | 行）
                            tbl_rows = []
                            while i < len(lines) and (lines[i].strip().startswith("|") or
                                  (not lines[i].strip() and i+1 < len(lines) and lines[i+1].strip().startswith("|"))):
                                if lines[i].strip():
                                    tbl_rows.append(lines[i].strip())
                                i += 1
                            output_parts.append(("table", tbl_rows))
                        else:
                            # 收集到下一個表格行或結束
                            md_lines = []
                            while i < len(lines) and not lines[i].strip().startswith("|"):
                                md_lines.append(lines[i])
                                i += 1
                            output_parts.append(("md", md_lines))

                    # 渲染各分組
                    final_html_parts = []
                    for part_type, part_data in output_parts:
                        if part_type == "table":
                            final_html_parts.append(parse_md_table_to_html(part_data))
                        else:
                            # 一般 Markdown 段落：補空行確保換行
                            md_text = ""
                            for ln in part_data:
                                s = ln.strip()
                                md_text += ln + "\n"
                                if s and not s.startswith("---") and not s.startswith("|"):
                                    md_text += "\n"
                            if md_text.strip():
                                # 轉為 HTML 段落（保留 Markdown 格式）
                                import html as html_mod
                                # 直接用 st.markdown 輸出非表格部分
                                final_html_parts.append(f"MARKDOWN_BLOCK:{md_text}")

                    # 輸出：HTML 表格用 st.markdown(unsafe_allow_html=True)，
                    # Markdown 段落仍用 st.markdown
                    for part in final_html_parts:
                        if part.startswith("MARKDOWN_BLOCK:"):
                            st.markdown(part[len("MARKDOWN_BLOCK:"):])
                        else:
                            st.markdown(part, unsafe_allow_html=True)
        else:
            st.markdown(result_text)

        st.success("✅ 改寫完成！")
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                "⬇️ 下載 Word 檔（.docx）",
                data=result_to_docx(result_text),
                file_name="未來線中運量_改寫結果.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with dl_col2:
            st.download_button(
                "⬇️ 下載純文字（.txt）",
                data=result_text.encode("utf-8"),
                file_name="未來線中運量_改寫結果.txt",
                mime="text/plain",
                use_container_width=True,
            )
    else:
        st.info("尚未產出結果。")
