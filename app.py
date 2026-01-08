import streamlit as st
import pandas as pd
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import zipfile
import time

# ==========================================
# 專案：班級讀書建議生成器 (Word 嚴格版 + 模型選擇)
# 功能：
# 1. 讀取 Excel (5分頁)
# 2. 可選擇 Gemini 模型
# 3. AI 生成建議 (GEM 嚴格提示詞)
# 4. 產出 Word 檔 (.docx)
# ==========================================

# --- 1. 網頁設定 ---
st.set_page_config(page_title="班級讀書建議生成器", layout="wide")
st.title("🎓 班級錯題分析與讀書建議生成器 (Word版)")
st.markdown("""
此工具協助老師快速生成全班學生的個別化讀書建議 **Word 檔**。
1. 輸入您的 **Gemini API Key** 並 **選擇模型**。
2. 上傳 **Excel 檔案** (需包含 國文, 英文, 數學, 社會, 自然 5個分頁)。
3. 系統將自動分析並打包 ZIP 下載。
""")

# --- 2. 核心邏輯函式 ---

def process_excel_data(uploaded_file):
    """讀取 Excel 並整理所有學生的錯題"""
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception:
        return None, "檔案格式錯誤，請確認上傳的是 .xlsx Excel 檔案。"

    # 檢查分頁是否齊全
    required_sheets = ["國文", "英文", "數學", "社會", "自然"]
    missing_sheets = [sheet for sheet in required_sheets if sheet not in xls.sheet_names]
    
    if missing_sheets:
        return None, f"Excel 缺少必要分頁：{missing_sheets}，請確認分頁名稱正確。"

    # 讀取所有資料
    data_map = {}
    for sheet in required_sheets:
        # header=None 代表不使用第一列當標題，我們依索引讀取
        data_map[sheet] = pd.read_excel(xls, sheet_name=sheet, header=None)

    # 取得學生名單 (以國文科為準)
    try:
        first_df = data_map["國文"]
        # 假設第 6 列 (Index 5) 的 B 欄 (Index 1) 是姓名
        student_list = first_df.iloc[5:, 1].dropna().unique().tolist()
    except Exception as e:
        return None, f"無法讀取學生名單，請確認 Excel 格式 (錯誤訊息: {e})"
    
    # 整理每位學生的錯題
    all_students_data = {}
    
    for student in student_list:
        student_errors = {}
        for subject in required_sheets:
            df = data_map[subject]
            try:
                # 解析結構
                q_nums = df.iloc[0, 2:].values
                categories = df.iloc[1, 2:].values
                k_points = df.iloc[2, 2:].values
                
                # 找學生列
                student_df_temp = df.iloc[5:, 1:].reset_index(drop=True)
                student_df_temp.columns = ["Name"] + [i for i in range(len(student_df_temp.columns)-1)]
                
                target_row = student_df_temp[student_df_temp["Name"] == student]
                
                if target_row.empty:
                    continue
                
                answers = target_row.iloc[0, 1:].values
                
                errors = []
                for ans, cat, kp, qn in zip(answers, categories, k_points, q_nums):
                    ans_str = str(ans).strip()
                    # 錯題判斷：不是 "-" 且不是空白
                    if ans_str != "-" and pd.notna(ans) and ans_str != "":
                        errors.append({
                            "題號": qn,
                            "領域": str(cat).strip() if pd.notna(cat) else "其他",
                            "知識點": kp
                        })
                student_errors[subject] = errors
            except Exception as e:
                print(f"處理 {student} 的 {subject} 時發生錯誤: {e}")
                
        all_students_data[student] = student_errors
        
    return all_students_data, None

def get_ai_advice(api_key, model_name, student_name, error_data):
    """呼叫 Gemini 生成建議 (使用 GEM 嚴格版 Prompt)"""
    try:
        genai.configure(api_key=api_key)
        # 使用使用者選擇的模型
        model = genai.GenerativeModel(model_name)
        
        # 這是 GEM 嚴格版提示詞
        prompt = f"""
        你是一位專業的台灣國中教育會考升學輔導專家。你的任務是讀取以下學生的錯題數據（九年級第2次複習考，範圍1-4冊），並生成一份精準的讀書建議報告。

        學生姓名：{student_name} (請在報告中一律稱呼為「你」)
        錯題數據：{error_data}

        請嚴格遵守以下規則進行分析與輸出：

        ### 核心規則
        1. **直接開始**：**絕對不要**有任何開場白（如「親愛的同學你好」）。請直接以「## 一、 【整體表現總評】」作為輸出的第一行。
        2. **統一稱呼**：報告中若需提及學生，請一律使用代名詞**「你」**。
        3. **無結尾提問**：報告結束時，請給予一句簡短的鼓勵即可，不要詢問問題。
        4. **格式一致性**：必須嚴格依照下方的【輸出範本】格式進行排版。

        ### 步驟一：資料分類邏輯 (請運用你的專業判斷)
        *   **國文**：文言文 / 白話文
        *   **英文**：聽力 / 閱讀
        *   **數學**：代數 / 幾何 / 機率統計
        *   **社會**：歷史 / 地理 / 公民
        *   **自然**：生物 / 理化 / 地科 (請特別注意地科內容如天文、地質、氣象)

        ### 步驟二：輸出範本 (Output Template)
        請完全依照以下 Markdown 結構輸出內容：

        ## 一、 【整體表現總評】

        * **強弱科分析**：
            * **穩定發展科（強科）**：**[科目名]**（[錯題數]題）。[簡短評語]
            * **急需搶救科（弱科）**：**[科目名]**（[錯題數]題）。[簡短評語]

        * **關鍵弱點領域**：
        [跨科目分析該生的痛點。例如：是「記憶性」較弱，還是「邏輯推演」較弱？]

        ---

        ## 二、 【分科深度分析與建議】

        ### 1. 國文科：[請給予一句該科的總結短評]
        * **弱點診斷 (前三名)**：
            1. **【[領域名]】** [知識點名稱]
            2. **【[領域名]】** [知識點名稱]
            3. **【[領域名]】** [知識點名稱]
        * **領域佔比分析**：[描述佔比]
        * **會考衝刺建議**：[針對弱點提供具體讀書建議]

        ### 2. 英文科：[請給予一句該科的總結短評]
        (格式同上)

        ### 3. 數學科：[請給予一句該科的總結短評]
        (格式同上)

        ### 4. 社會科：[請給予一句該科的總結短評]
        (格式同上)

        ### 5. 自然科：[請給予一句該科的總結短評]
        (格式同上)

        ---
        **[請在此處給予一段總結性的鼓勵話語]**
        """
        
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"AI 分析連線失敗: {e} (請檢查 API Key 或模型權限)"

def create_word(student_name, ai_advice):
    """
    建立 Word 文件 (.docx)
    """
    doc = Document()
    
    # 1. 加入標題
    title = doc.add_heading(f"{student_name} - 讀書建議報告", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 2. 處理 AI 建議內容
    # 簡單清理 Markdown 符號
    clean_text = ai_advice.replace('**', '').replace('## ', '').replace('### ', '')
    
    for paragraph_text in clean_text.split('\n'):
        if paragraph_text.strip():
            p = doc.add_paragraph(paragraph_text)
            p.style.font.size = Pt(12)
            
    # 3. 存入記憶體 Buffer
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- 4. 介面互動邏輯 ---

# 側邊欄：輸入 API Key 與 選擇模型
with st.sidebar:
    st.header("🔑 設定")
    user_api_key = st.text_input("請輸入 Gemini API Key", type="password", help="請前往 Google AI Studio 申請")
    
    # 新增：模型選擇器
    model_options = [
        "gemini-1.5-flash", 
        "gemini-1.5-pro", 
        "gemini-2.0-flash-exp"
    ]
    selected_model = st.selectbox(
        "🤖 選擇 AI 模型", 
        model_options, 
        index=0,
        help="Flash 速度快且免費額度高；Pro 分析能力更強但速度稍慢。"
    )
    
    st.markdown("---")
    st.info("💡 提示：請上傳包含 5 個分頁 (國文, 英文, 數學, 社會, 自然) 的 Excel 檔案。")

# 主畫面：上傳檔案
uploaded_file = st.file_uploader("📂 上傳 Excel 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file and user_api_key:
    if st.button("🚀 開始生成全班報告 (Word)"):
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        # 1. 處理 Excel
        status_text.text("正在讀取 Excel 資料...")
        all_data, error_msg = process_excel_data(uploaded_file)
        
        if error_msg:
            st.error(error_msg)
        else:
            # 2. 準備 ZIP
            zip_buffer = io.BytesIO()
            total_students = len(all_data)
            
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for i, (student, errors) in enumerate(all_data.items()):
                    # 更新進度
                    progress = (i + 1) / total_students
                    progress_bar.progress(progress)
                    status_text.text(f"正在分析：{student} ({i+1}/{total_students})...")
                    
                    # AI 生成 (傳入選擇的模型)
                    advice = get_ai_advice(user_api_key, selected_model, student, str(errors))
                    
                    # Word 生成
                    word_data = create_word(student, advice)
                    
                    # 加入 ZIP
                    zf.writestr(f"{student}_讀書建議.docx", word_data.getvalue())
                    
                    # 稍微休息一下避免 API 限制
                    time.sleep(1)
            
            progress_bar.progress(100)
            status_text.success("✅ 生成完成！")
            
            # 3. 下載按鈕
            st.download_button(
                label="📥 下載全班報告 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="全班讀書建議報告_Word.zip",
                mime="application/zip"
            )

elif uploaded_file and not user_api_key:
    st.warning("請在左側輸入 API Key 才能開始執行。")