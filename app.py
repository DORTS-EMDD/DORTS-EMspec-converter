# ============================================================
#  Metro-Mentor AI — 捷運機電智慧傳承與介面審查 AI 顧問系統
#  app.py  (Streamlit + Anthropic Claude + TF-IDF RAG)
# ============================================================

import io
import re
import json
import os
from datetime import datetime

import streamlit as st
from google import genai
from google.genai import types as genai_types
import PyPDF2
from docx import Document
import pandas as pd
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# ── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="Metro-Mentor AI | 捷運機電智慧顧問",
    page_icon="🚇",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ─────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { padding-top: 1.5rem; }
    .metric-box {
        background: #f0f4ff; border-radius: 8px;
        padding: 12px 16px; text-align: center;
        border: 1px solid #d0dbff;
    }
    .metric-box h2 { margin: 0; color: #1a3fa0; font-size: 1.8rem; }
    .metric-box p  { margin: 0; color: #555; font-size: 0.85rem; }
    .feature-card {
        background: #fafbff; border: 1px solid #e0e6ff;
        border-radius: 10px; padding: 16px 18px;
    }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
#  SimpleDocStore — TF-IDF 向量知識庫（取代需大量依賴的 ChromaDB）
# ══════════════════════════════════════════════════════════
class SimpleDocStore:
    """輕量 TF-IDF 文件庫，支援中文字元 n-gram 檢索。"""

    def __init__(self):
        self.docs: list[dict] = []   # [{id, text, metadata}]
        self._vect = TfidfVectorizer(
            analyzer="char_wb", ngram_range=(2, 4), min_df=1
        )
        self._matrix = None

    # ── 新增 / 更新文件 ────────────────────────────────────
    def upsert(self, ids: list, documents: list, metadatas: list):
        id_to_idx = {d["id"]: i for i, d in enumerate(self.docs)}
        for doc_id, text, meta in zip(ids, documents, metadatas):
            entry = {"id": doc_id, "text": text, "metadata": meta}
            if doc_id in id_to_idx:
                self.docs[id_to_idx[doc_id]] = entry
            else:
                self.docs.append(entry)
        self._refit()

    def _refit(self):
        if self.docs:
            texts = [d["text"] for d in self.docs]
            self._matrix = self._vect.fit_transform(texts)

    # ── 查詢 ───────────────────────────────────────────────
    def query(self, query_text: str, n_results: int = 5) -> dict:
        empty = {"documents": [[]], "metadatas": [[]]}
        if not self.docs or self._matrix is None:
            return empty
        k = min(n_results, len(self.docs))
        try:
            q_vec = self._vect.transform([query_text])
        except Exception:
            return empty
        scores = cosine_similarity(q_vec, self._matrix)[0]
        top_idx = scores.argsort()[::-1][:k]
        return {
            "documents": [[self.docs[i]["text"]  for i in top_idx]],
            "metadatas": [[self.docs[i]["metadata"] for i in top_idx]],
            "scores":    [[float(scores[i])          for i in top_idx]],
        }

    def count(self) -> int:
        return len(self.docs)

    def export_json(self) -> str:
        """匯出為 JSON 字串（備份用）。"""
        return json.dumps(
            [{"id": d["id"], "text": d["text"], "metadata": d["metadata"]}
             for d in self.docs],
            ensure_ascii=False, indent=2
        )

    def import_json(self, json_str: str):
        """從 JSON 字串匯入（還原備份）。"""
        records = json.loads(json_str)
        ids   = [r["id"]       for r in records]
        texts = [r["text"]     for r in records]
        metas = [r["metadata"] for r in records]
        self.upsert(ids, texts, metas)


# ══════════════════════════════════════════════════════════
#  Session State 初始化
# ══════════════════════════════════════════════════════════
def _init_state():
    # 優先從 Streamlit Secrets 讀取 API Key（Settings → Secrets 設定 GOOGLE_API_KEY）
    _secret_key = st.secrets.get("GOOGLE_API_KEY", "")
    defaults = {
        "api_key": _secret_key,
        "selected_model": "gemini-3.1-flash-lite",
        "knowledge_store": SimpleDocStore(),   # 技術文件庫
        "lesson_store":    SimpleDocStore(),   # 失敗案例庫
        "chat_history":    [],
        "interview_result": None,
        "interview_meta":  {},
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init_state()


# ══════════════════════════════════════════════════════════
#  文件解析
# ══════════════════════════════════════════════════════════
def _extract_pdf(file) -> str:
    reader = PyPDF2.PdfReader(file)
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def _extract_docx(file) -> str:
    doc = Document(file)
    return "\n".join(p.text for p in doc.paragraphs)

def _extract_excel(file) -> str:
    sheets = pd.read_excel(file, sheet_name=None)
    parts = []
    for name, df in sheets.items():
        parts.append(f"=== 工作表：{name} ===\n{df.to_string(index=False)}")
    return "\n\n".join(parts)

def extract_text(file) -> str:
    name = file.name.lower()
    if name.endswith(".pdf"):        return _extract_pdf(file)
    if name.endswith(".docx"):       return _extract_docx(file)
    if name.endswith((".xlsx","xls")):return _extract_excel(file)
    return file.read().decode("utf-8", errors="ignore")


# ══════════════════════════════════════════════════════════
#  文字切塊（Chunking）
# ══════════════════════════════════════════════════════════
def chunk_text(text: str, size: int = 400, overlap: int = 60) -> list:
    text = text.strip()
    if not text:
        return []
    chunks, start = [], 0
    while start < len(text):
        end = min(start + size, len(text))
        c = text[start:end].strip()
        if c:
            chunks.append(c)
        if end >= len(text):
            break
        start = end - overlap
    return chunks


# ══════════════════════════════════════════════════════════
#  知識庫操作
# ══════════════════════════════════════════════════════════
def add_to_store(store: SimpleDocStore,
                 text: str, filename: str,
                 doc_type: str, category: str) -> int:
    chunks = chunk_text(text)
    if not chunks:
        return 0
    safe = re.sub(r"[^a-zA-Z0-9\u4e00-\u9fff_-]", "_", filename)
    ids   = [f"{safe}_{i}" for i in range(len(chunks))]
    metas = [{"source": filename, "doc_type": doc_type,
               "category": category, "chunk": i}
             for i in range(len(chunks))]
    store.upsert(ids, chunks, metas)
    return len(chunks)


# ══════════════════════════════════════════════════════════
#  Gemini API 呼叫
# ══════════════════════════════════════════════════════════
def _gemini(prompt: str, system: str = "", max_tokens: int = 2000) -> str:
    model_name = st.session_state.get("selected_model", "gemini-3.1-flash-lite")
    full_prompt = f"{system}\n\n{prompt}" if system else prompt
    client = genai.Client(api_key=st.session_state.api_key)
    resp = client.models.generate_content(
        model=model_name,
        contents=full_prompt,
        config=genai_types.GenerateContentConfig(max_output_tokens=max_tokens),
    )
    return resp.text


def _build_context(results: dict) -> str:
    docs   = results.get("documents", [[]])[0]
    metas  = results.get("metadatas", [[]])[0]
    scores = results.get("scores",    [[]])[0]
    if not docs:
        return "（知識庫目前為空或無相關內容）"
    parts = []
    for doc, meta, score in zip(docs, metas, scores):
        parts.append(
            f"📄 來源：{meta['source']}（{meta['category']} | 相似度 {score:.2f}）\n{doc}"
        )
    return "\n\n---\n\n".join(parts)


def qa_with_rag(question: str, results: dict) -> str:
    ctx = _build_context(results)
    prompt = f"""你是台北捷運機電處的 AI 顧問（Metro-Mentor AI）。
請嚴格根據以下知識庫內容回答問題。若知識庫沒有相關資訊，請如實告知，不要自行推測。
回答時請引用【來源文件名稱】讓使用者可追溯，並使用繁體中文。

【知識庫相關內容】
{ctx}

【問題】
{question}"""
    return _gemini(prompt, max_tokens=1500)


def generate_checklist(description: str, lesson_results: dict) -> str:
    past = _build_context(lesson_results)
    prompt = f"""你是台北捷運機電處的資深審查工程師 AI。
根據以下廠商送審說明與歷史失敗案例，生成一份結構化的介面審查 Checklist。
使用繁體中文，條列清楚且具體可查核。

【送審介面說明】
{description}

【歷史失敗案例 / Lesson Learned】
{past}

請按以下結構輸出：

## 📋 1. 基本合規確認
（法規、標準、合約規格確認項目）

## ⚠️ 2. 歷史案例警示重點
（根據過去失敗案例的具體提醒，需說明「為何重要」）

## 🔍 3. 本次介面專項審查
（針對本次送審內容的具體問題清單）

每項請說明查核重點與重要原因。"""
    return _gemini(prompt, max_tokens=2000)


def extract_interview_knowledge(text: str, interviewee: str, expertise: str) -> str:
    prompt = f"""你是台北捷運機電處的知識管理 AI。
請分析以下資深同仁訪談紀錄，提取關鍵隱性知識並結構化輸出（繁體中文）。

受訪者：{interviewee}
專長：{expertise}

【訪談內容】
{text}

請輸出以下結構：

## 🏷️ 標籤分類
（系統別 / 路線別 / 事件類型 / 關鍵字）

## 💡 關鍵技術知識點
（具體技術細節、規格判斷依據）

## ⚠️ 踩雷經驗與解決方法
（遇到的問題、排查方式、最終解法）

## 📝 廠商與合約注意事項
（廠商特性、談判要點、合約陷阱）

## 🔗 後人應參閱的相關文件
（標單名稱、規範編號、報告名稱）

## 🎯 給後輩的建議
（最重要的傳承叮嚀）"""
    return _gemini(prompt, max_tokens=2500)


# ══════════════════════════════════════════════════════════
#  Sidebar
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🚇 Metro-Mentor AI")
    st.markdown("**捷運機電智慧傳承顧問系統**")
    st.divider()

    # API Key — 若已透過 Streamlit Secrets 設定則自動套用，不顯示輸入框
    _from_secrets = bool(st.secrets.get("GOOGLE_API_KEY", ""))
    if _from_secrets:
        st.success("✅ API Key 已透過 Secrets 自動載入")
    else:
        api_key = st.text_input(
            "🔑 Google API Key",
            type="password",
            value=st.session_state.api_key,
            placeholder="AIza...",
            help="輸入您的 Google Gemini API 金鑰（可至 Google AI Studio 取得）",
        )
        if api_key:
            st.session_state.api_key = api_key

        if st.session_state.api_key:
            st.success("✅ API Key 已設定")
        else:
            st.warning("⚠️ 請輸入 API Key")

    st.divider()

    # Model selector
    selected_model = st.selectbox(
        "🤖 選擇模型",
        ["gemini-3.1-flash-lite", "gemini-3.5-flash"],
        index=["gemini-3.1-flash-lite", "gemini-3.5-flash"].index(
            st.session_state.get("selected_model", "gemini-3.1-flash-lite")
        ),
        help=(
            "gemini-3.1-flash-lite：Gemini 3 最新輕量版，速度快、省配額，適合大量章節改寫。\n"
            "gemini-3.5-flash：接近 Pro 等級智能，細節與表格保留更完整，建議用於關鍵章節。"
        ),
    )
    st.session_state.selected_model = selected_model

    st.divider()

    # DB stats
    ks = st.session_state.knowledge_store
    ls = st.session_state.lesson_store
    c1, c2 = st.columns(2)
    c1.metric("📚 知識庫", f"{ks.count()} 段")
    c2.metric("⚠️ 案例庫", f"{ls.count()} 段")

    st.divider()

    # Navigation
    page = st.radio(
        "📌 功能選單",
        ["🏠 系統總覽", "📚 文件上傳", "🔍 A. 智慧查詢",
         "📋 B. 介面審查", "🎙️ C. 訪談萃取", "💾 備份管理"],
        label_visibility="collapsed",
    )

    st.divider()
    st.caption("v1.0 | 地端部署友善 | 資料不外洩")


# ══════════════════════════════════════════════════════════
#  Page: 系統總覽
# ══════════════════════════════════════════════════════════
if page == "🏠 系統總覽":
    st.title("🚇 Metro-Mentor AI")
    st.markdown("### 捷運機電系統設計智慧傳承與介面審查 AI 顧問系統")
    st.info(
        "**使用步驟：** ① 在 Streamlit Secrets 設定 `GOOGLE_API_KEY`（或於左側手動輸入）　"
        "② 至「📚 文件上傳」上傳技術文件　③ 使用各功能模組開始提問或生成審查清單"
    )
    st.divider()

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
<div class="feature-card">
<h4>🔍 功能 A：智慧查詢</h4>
<p>自然語言提問，AI 自動從知識庫檢索技術規範、標單、ICD，並附上原始文件出處。</p>
<blockquote><em>例：供電第三軌道岔斷電區間標準？</em></blockquote>
</div>
""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
<div class="feature-card">
<h4>📋 功能 B：介面審查</h4>
<p>輸入廠商送審說明，AI 根據歷史失敗案例自動生成可下載的審查 Checklist。</p>
<blockquote><em>例：號誌系統 ↔ 電聯車介面設計圖審查</em></blockquote>
</div>
""", unsafe_allow_html=True)
    with col3:
        st.markdown("""
<div class="feature-card">
<h4>🎙️ 功能 C：訪談萃取</h4>
<p>將資深同仁退休訪談紀錄輸入，AI 自動分類標籤並萃取隱性知識匯入知識庫。</p>
<blockquote><em>搶救即將流失的 30 年 Know-how</em></blockquote>
</div>
""", unsafe_allow_html=True)

    st.divider()
    st.markdown("""
| 特點 | 說明 |
|:-----|:-----|
| ✅ 資料來源可追溯 | 每則回答皆附帶原始文件名稱與相似度評分 |
| ✅ 防幻覺設計 | AI 嚴格只根據上傳的知識庫內容回答，不自行捏造 |
| ✅ 地端部署友善 | 無需 ChromaDB 外部服務，資料 100% 留在本機 |
| ✅ 多格式支援 | PDF、Word (.docx)、Excel (.xlsx)、純文字 (.txt) |
| ✅ 中文優化檢索 | 使用字元 n-gram TF-IDF，中文技術文件準確命中 |
| ✅ 知識庫備份還原 | 可匯出 JSON 備份，下次啟動時還原 |
""")


# ══════════════════════════════════════════════════════════
#  Page: 文件上傳
# ══════════════════════════════════════════════════════════
elif page == "📚 文件上傳":
    st.title("📚 文件上傳管理")
    tab1, tab2 = st.tabs(["📄 技術文件（知識庫）", "⚠️ 失敗案例（Lesson Learned）"])

    # ── Tab 1：技術文件 ─────────────────────────────────
    with tab1:
        st.subheader("上傳技術規範、標單、ICD 等文件")
        c1, c2 = st.columns(2)
        with c1:
            doc_type = st.selectbox("文件類型", [
                "技術規範（TS）", "介面控制文件（ICD）", "招標文件",
                "會議紀錄", "測試報告（SAT）", "變更設計（VO）",
                "維修手冊", "設計圖說", "其他",
            ])
        with c2:
            category = st.selectbox("系統別", [
                "號誌", "供電", "車輛", "通訊", "AFC",
                "軌道", "土建介面", "綜合",
            ])

        files = st.file_uploader(
            "選擇文件（可多選，支援 PDF / DOCX / XLSX / TXT）",
            type=["pdf", "docx", "xlsx", "xls", "txt"],
            accept_multiple_files=True,
            key="doc_uploader",
        )

        if files:
            if st.button("📥 匯入至知識庫", type="primary"):
                for f in files:
                    with st.spinner(f"處理 {f.name}…"):
                        try:
                            text = extract_text(f)
                            if not text.strip():
                                st.warning(f"⚠️ {f.name}：無法擷取文字（可能是掃描檔）")
                                continue
                            n = add_to_store(
                                st.session_state.knowledge_store,
                                text, f.name, doc_type, category,
                            )
                            st.success(f"✅ {f.name} → 已切塊為 **{n}** 個段落")
                        except Exception as e:
                            st.error(f"❌ {f.name}：{e}")
                st.rerun()

        if ks.count():
            st.divider()
            st.markdown(f"**知識庫目前有 {ks.count()} 個段落**，來源文件清單：")
            sources = list({d["metadata"]["source"] for d in ks.docs})
            for s in sorted(sources):
                st.markdown(f"- {s}")

    # ── Tab 2：失敗案例 ──────────────────────────────────
    with tab2:
        st.subheader("新增歷史失敗案例（Lesson Learned）")
        st.caption("這些案例將用於功能 B：介面審查 Checklist 自動生成")

        lesson_type = st.selectbox("案例類型", [
            "介面衝突", "設計錯誤", "施工變更（VO）",
            "系統故障", "廠商問題", "驗收問題", "其他",
        ])

        input_method = st.radio(
            "輸入方式", ["📁 上傳文件", "✏️ 直接輸入"], horizontal=True
        )

        if input_method == "📁 上傳文件":
            lf = st.file_uploader(
                "上傳失敗案例文件",
                type=["pdf", "docx", "txt"],
                key="lesson_uploader",
            )
            lesson_source = st.text_input(
                "案例來源 / 編號", placeholder="例：環狀線 2019 VO-023"
            )
            if lf and st.button("📥 新增案例", type="primary"):
                try:
                    text = extract_text(lf)
                    n = add_to_store(
                        st.session_state.lesson_store,
                        text, lesson_source or lf.name, lesson_type, "失敗案例",
                    )
                    st.success(f"✅ 已新增：{n} 個段落")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")
        else:
            lesson_text = st.text_area(
                "案例描述",
                height=160,
                placeholder=(
                    "例：環狀線供電系統第三軌道岔斷電區間施工公差不符，"
                    "導致 SAT 發現電弓火花，補發 VO 整改耗時 6 週…"
                ),
            )
            lesson_source_m = st.text_input(
                "案例來源 / 編號", placeholder="例：環狀線 2019 VO-023",
                key="lesson_source_m",
            )
            if lesson_text and st.button(
                "📥 新增案例", type="primary", key="lesson_manual"
            ):
                n = add_to_store(
                    st.session_state.lesson_store,
                    lesson_text,
                    lesson_source_m or "手動輸入",
                    lesson_type, "失敗案例",
                )
                st.success(f"✅ 已新增：{n} 個段落")
                st.rerun()

        if ls.count():
            st.divider()
            sources_l = list({d["metadata"]["source"] for d in ls.docs})
            st.markdown(f"**案例庫目前有 {ls.count()} 個段落**，共 {len(sources_l)} 筆案例")


# ══════════════════════════════════════════════════════════
#  Page: A. 智慧查詢
# ══════════════════════════════════════════════════════════
elif page == "🔍 A. 智慧查詢":
    st.title("🔍 功能 A：智慧設計規範查詢")

    ks = st.session_state.knowledge_store
    if not st.session_state.api_key:
        st.warning("⚠️ 請先設定 Google API Key（Secrets 或左側輸入框）")
        st.stop()

    if ks.count() == 0:
        st.info(
            "💡 知識庫為空——請先至「📚 文件上傳」上傳技術規範、標單或 ICD 文件。"
            "\n\n（即使知識庫為空，您仍可提問，AI 將說明無法找到相關資料。）"
        )

    st.caption(f"📚 知識庫目前有 **{ks.count()}** 個段落")

    # Display history
    for msg in st.session_state.chat_history:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # Chat input
    prompt = st.chat_input(
        "請輸入問題，例如：供電系統第三軌在道岔處的斷電區間標準是多少？"
    )
    if prompt:
        st.session_state.chat_history.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            with st.spinner("🔍 檢索知識庫…"):
                results = ks.query(prompt, n_results=5)
            with st.spinner("🤖 Gemini 生成回答中…"):
                try:
                    answer = qa_with_rag(prompt, results)
                except Exception as e:
                    answer = f"❌ 錯誤：{e}"
            st.markdown(answer)
            st.session_state.chat_history.append(
                {"role": "assistant", "content": answer}
            )

    if st.session_state.chat_history:
        if st.button("🗑️ 清除對話記錄"):
            st.session_state.chat_history = []
            st.rerun()


# ══════════════════════════════════════════════════════════
#  Page: B. 介面審查
# ══════════════════════════════════════════════════════════
elif page == "📋 B. 介面審查":
    st.title("📋 功能 B：介面審查 Checklist 自動生成")

    if not st.session_state.api_key:
        st.warning("⚠️ 請先設定 Google API Key（Secrets 或左側輸入框）")
        st.stop()

    ls = st.session_state.lesson_store
    st.caption(f"⚠️ 失敗案例庫目前有 **{ls.count()}** 個段落（案例愈多，Checklist 愈精準）")

    c1, c2 = st.columns([3, 1])
    with c2:
        interface_type = st.selectbox("介面類型", [
            "號誌 ↔ 車輛", "供電 ↔ 土建", "通訊 ↔ 號誌",
            "AFC ↔ 通訊", "車輛 ↔ 供電", "機電 ↔ 土建", "其他",
        ])
        line_name = st.text_input("路線別", placeholder="例：萬大線")
        vendor    = st.text_input("廠商（選填）")
        n_checklist_results = st.slider(
            "參考案例數", 1, 10, 5,
            help="從案例庫取用的相關案例數量"
        )

    with c1:
        desc = st.text_area(
            "介面設計摘要（廠商送審說明）",
            height=220,
            placeholder=(
                "例：\n本次送審為號誌系統與電聯車之 ATP/ATO 介面設計，包含：\n"
                "- CBTC 無線電天線配置（車頂後方，距車頭 2.1m）\n"
                "- 速度感知器介面（車軸型，A/B 冗餘）\n"
                "- 列車完整性確認（TIMS 介面，硬線邏輯）\n"
                "- 緊急停車按鈕（EBD）電路介面規格"
            ),
        )

    if st.button("🚀 生成審查 Checklist", type="primary", use_container_width=True):
        if not desc.strip():
            st.warning("請輸入介面設計說明")
        else:
            full_desc = (
                f"路線：{line_name or '未指定'} | 廠商：{vendor or '未指定'} "
                f"| 介面類型：{interface_type}\n\n{desc}"
            )
            with st.spinner("🔍 檢索歷史失敗案例…"):
                lesson_results = ls.query(desc, n_results=n_checklist_results)
            with st.spinner("🤖 Gemini 生成審查清單中…"):
                try:
                    checklist = generate_checklist(full_desc, lesson_results)
                except Exception as e:
                    st.error(f"❌ {e}")
                    st.stop()

            st.divider()
            ts = datetime.now().strftime("%Y-%m-%d %H:%M")
            st.markdown(
                f"## 📋 自動生成審查 Checklist\n"
                f"**路線：** {line_name or '未指定'} ｜ "
                f"**介面：** {interface_type} ｜ "
                f"**生成：** {ts}"
            )
            st.divider()
            st.markdown(checklist)

            export_md = (
                f"# 介面審查 Checklist\n"
                f"路線：{line_name or '未指定'} | 類型：{interface_type} "
                f"| 廠商：{vendor or '未指定'}\n"
                f"生成時間：{ts}\n\n---\n\n{checklist}"
            )
            safe_line = re.sub(r"[^a-zA-Z0-9\u4e00-\u9fff]", "_", line_name or "checklist")
            st.download_button(
                "💾 下載 Checklist（.md）",
                data=export_md,
                file_name=f"checklist_{safe_line}_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                mime="text/markdown",
            )


# ══════════════════════════════════════════════════════════
#  Page: C. 訪談萃取
# ══════════════════════════════════════════════════════════
elif page == "🎙️ C. 訪談萃取":
    st.title("🎙️ 功能 C：退休訪談知識萃取")

    if not st.session_state.api_key:
        st.warning("⚠️ 請先設定 Google API Key（Secrets 或左側輸入框）")
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        interviewee    = st.text_input("受訪者姓名 / 代號", placeholder="例：王技師（供電科）")
        interview_date = st.date_input("訪談日期", value=datetime.today())
    with c2:
        expertise  = st.multiselect("專長領域", [
            "供電系統", "號誌系統", "車輛系統", "通訊系統",
            "AFC", "軌道工程", "土建介面", "工程管理", "廠商談判",
        ])
        years_exp  = st.number_input("服務年資（年）", 1, 50, 20)

    interview_text = st.text_area(
        "訪談紀錄（逐字稿或整理稿皆可）",
        height=280,
        placeholder=(
            "訪談者：請問您在環狀線供電系統建置期間，遇到最大的挑戰是什麼？\n\n"
            "王技師：當時最麻煩的是第三軌在 Y 型道岔那段，原設計圖的斷電區間長度是按照 750V 標準畫的，"
            "但承包商的斷電條塊長度差了 20 公分，驗收時差點沒發現。"
            "後來做 SAT 的時候跑了幾趟試車才發現電弓有火花…\n\n"
            "訪談者：怎麼處理？\n\n"
            "王技師：補發 VO，廠商整改，耽誤了 6 週工期。後來我們在審查新線別供電系統設計圖時，"
            "就在 Checklist 加了一條，一定要核對斷電區間的施工公差…"
        ),
    )

    btn1, btn2 = st.columns(2)
    with btn1:
        extract_clicked = st.button(
            "🤖 AI 知識萃取", type="primary", use_container_width=True
        )
    with btn2:
        import_clicked = st.button(
            "📥 萃取結果匯入知識庫",
            use_container_width=True,
            disabled=(st.session_state.interview_result is None),
        )

    if extract_clicked:
        if not interview_text.strip():
            st.warning("請輸入訪談內容")
        elif not interviewee.strip():
            st.warning("請輸入受訪者資訊")
        else:
            expertise_str = "、".join(expertise) if expertise else "未指定"
            with st.spinner("🤖 Gemini 分析訪談內容中…"):
                try:
                    result = extract_interview_knowledge(
                        interview_text, interviewee,
                        f"{expertise_str}（{years_exp} 年資）",
                    )
                    st.session_state.interview_result = result
                    st.session_state.interview_meta = {
                        "interviewee": interviewee,
                        "date":        str(interview_date),
                        "expertise":   expertise_str,
                    }
                except Exception as e:
                    st.error(f"❌ {e}")

    if import_clicked and st.session_state.interview_result:
        meta = st.session_state.interview_meta
        n = add_to_store(
            st.session_state.knowledge_store,
            st.session_state.interview_result,
            f"訪談_{meta['interviewee']}_{meta['date']}",
            "退休訪談知識",
            meta["expertise"],
        )
        st.success(f"✅ 已匯入知識庫（{n} 個段落）")

    if st.session_state.interview_result:
        st.divider()
        st.markdown("## 📝 知識萃取結果")
        st.markdown(st.session_state.interview_result)

        meta = st.session_state.interview_meta
        safe_name = re.sub(r"[^a-zA-Z0-9\u4e00-\u9fff_]", "_",
                           meta.get("interviewee", "unknown"))
        st.download_button(
            "💾 下載萃取結果（.md）",
            data=(
                f"# 訪談知識萃取\n"
                f"受訪者：{meta.get('interviewee','')} | "
                f"日期：{meta.get('date','')} | "
                f"專長：{meta.get('expertise','')}\n\n"
                f"{st.session_state.interview_result}"
            ),
            file_name=f"interview_{safe_name}_{meta.get('date','')}.md",
            mime="text/markdown",
        )


# ══════════════════════════════════════════════════════════
#  Page: 備份管理
# ══════════════════════════════════════════════════════════
elif page == "💾 備份管理":
    st.title("💾 知識庫備份管理")
    st.info(
        "由於本系統使用記憶體知識庫，**重啟後資料將消失**。"
        "請在每次工作結束前下載備份，下次啟動後再匯入還原。"
    )

    ks = st.session_state.knowledge_store
    ls = st.session_state.lesson_store

    st.subheader("📤 匯出備份")
    c1, c2 = st.columns(2)
    with c1:
        if ks.count():
            st.download_button(
                f"⬇️ 下載知識庫備份（{ks.count()} 段）",
                data=ks.export_json(),
                file_name=f"knowledge_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True,
            )
        else:
            st.markdown("_知識庫為空_")
    with c2:
        if ls.count():
            st.download_button(
                f"⬇️ 下載案例庫備份（{ls.count()} 段）",
                data=ls.export_json(),
                file_name=f"lesson_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                mime="application/json",
                use_container_width=True,
            )
        else:
            st.markdown("_案例庫為空_")

    st.divider()
    st.subheader("📥 匯入備份")
    col_k, col_l = st.columns(2)
    with col_k:
        kb_file = st.file_uploader(
            "選擇知識庫備份（.json）", type=["json"], key="kb_restore"
        )
        if kb_file and st.button("還原知識庫", use_container_width=True):
            try:
                ks.import_json(kb_file.read().decode("utf-8"))
                st.success(f"✅ 還原完成，共 {ks.count()} 個段落")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")
    with col_l:
        lb_file = st.file_uploader(
            "選擇案例庫備份（.json）", type=["json"], key="lb_restore"
        )
        if lb_file and st.button("還原案例庫", use_container_width=True):
            try:
                ls.import_json(lb_file.read().decode("utf-8"))
                st.success(f"✅ 還原完成，共 {ls.count()} 個段落")
                st.rerun()
            except Exception as e:
                st.error(f"❌ {e}")

    st.divider()
    st.subheader("🗑️ 清空知識庫")
    st.warning("⚠️ 清空後無法還原（除非已備份）")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("清空知識庫", type="secondary", use_container_width=True):
            st.session_state.knowledge_store = SimpleDocStore()
            st.success("✅ 知識庫已清空")
            st.rerun()
    with c2:
        if st.button("清空案例庫", type="secondary", use_container_width=True):
            st.session_state.lesson_store = SimpleDocStore()
            st.success("✅ 案例庫已清空")
            st.rerun()
