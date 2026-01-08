import streamlit as st
import pandas as pd
import google.generativeai as genai
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from reportlab.lib.units import cm
import io
import zipfile
import os
import requests

# ==========================================
# 專案：班級讀書建議生成器 (Excel版)
# 功能：
# 1. 網頁介面，讓不同老師輸入自己的 API Key
# 2. 讀取單一 Excel 檔 (含5個分頁)
# 3. 產出簡潔版 PDF (僅姓名 + 建議)
# ==========================================

# --- 1. 網頁設定 ---
st.set_page_config(page_title="班級讀書建議生成器", layout="wide")
st.title("🎓 班級錯題分析與讀書建議生成器")
st.markdown("""
此工具協助老師快速生成全班學生的個別化讀書建議 PDF。
1. 輸入您的 **Gemini API Key**。
2. 上傳 **Excel 檔案** (需包含 國文, 英文, 數學, 社會, 自然 5個分頁)。
3. 系統將自動分析並打包 PDF 下載。
""")

# --- 2. 系統字型處理 (解決 PDF 中文亂碼) ---
@st.cache_resource
def download_font():
    """下載中文字型到系統暫存區"""
    font_url = "https://github.com/googlefonts/noto-fonts/raw/main/hinted/ttf/NotoSansTC/NotoSansTC-Regular.ttf"
    font_path = "NotoSansTC-Regular.ttf"
    if not os.path.exists(font_path):
        with st.spinner("正在下載中文字型資源..."):
            response = requests.get(font_url)
            with open(font_path, "wb") as f:
                f.write(response.content)
    return font_path

try:
    font_path = download_font()
    pdfmetrics.registerFont(TTFont('NotoSans', font_path))
except Exception as e:
    st.error(f"字型載入失敗: {e}")

# --- 3. 核心邏輯函式 ---

def process_excel_data(uploaded_file):
    """讀取 Excel 並整理所有學生的錯題"""
    # 讀取 Excel 所有分頁
    xls = pd.ExcelFile(uploaded_file)
    
    # 檢查分頁是否齊全
    required_sheets = ["國文", "英文", "數學", "社會", "自然"]
    if not all(sheet in xls.sheet_names for sheet in required_sheets):
        return None, f"Excel 缺少必要分頁，請確認包含：{required_sheets}"

    # 讀取所有資料
    data_map = {}
    for sheet in required_sheets:
        # header=None 代表不使用第一列當標題，我們依索引讀取
        data_map[sheet] = pd.read_excel(xls, sheet_name=sheet, header=None)

    # 取得學生名單 (以國文科為準，假設第6列開始是學生)
    first_df = data_map["國文"]
    # 第 6 列 (Index 5) 的 B 欄 (Index 1) 是姓名
    student_list = first_df.iloc[5:, 1].dropna().unique().tolist()
    
    # 整理每位學生的錯題
    all_students_data = {}
    
    for student in student_list:
        student_errors = {}
        for subject in required_sheets:
            df = data_map[subject]
            try:
                # 解析結構
                # Row 0: 題號, Row 1: 分類, Row 2: 知識點
                q_nums = df.iloc[0, 2:].values
                categories = df.iloc[1, 2:].values
                k_points = df.iloc[2, 2:].values
                
                # 找學生列
                # 先把資料轉成 DataFrame 方便搜尋
                student_df_temp = df.iloc[5:, 1:].reset_index(drop=True)
                # 重新命名欄位以便搜尋：第一欄設為 Name
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

def get_ai_advice(api_key, student_name, error_data):
    """呼叫 Gemini 生成建議"""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    你是一位專業的國中會考升學輔導專家。請根據以下學生的錯題數據，撰寫一份精準的讀書建議。

    學生姓名：{student_name} (請在文中稱呼為「你」)
    錯題數據：{error_data}

    請嚴格遵守以下規則：
    1. **直接開始**：不要有開場白，不要打招呼。
    2. **格式**：請使用 Markdown 格式。
    3. **內容結構**：
       ## 一、 整體表現總評
       (分析強弱科與關鍵弱點)
       ## 二、 分科深度分析與建議
       (針對有錯題的科目，列出弱點領域並給予具體建議)
    4. **語氣**：溫暖、鼓勵且專業。
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"AI 分析連線失敗: {e}"

def create_pdf(student_name, ai_advice):
    """
    繪製 PDF
    修改：移除第一頁錯題表，移除 AI 標題，只保留姓名與建議
    """
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    # --- 標題：學生姓名 ---
    c.setFont("NotoSans", 24)
    # 畫在頁面頂端
    c.drawString(2*cm, height - 3*cm, f"📊 {student_name} - 讀書建議報告")
    
    # --- 內容：AI 建議 ---
    c.setFont("NotoSans", 11)
    
    # 文字換行處理
    text_object = c.beginText(2*cm, height - 5*cm)
    text_object.setFont("NotoSans", 11)
    text_object.setLeading(16) # 行距
    
    # 簡易 Markdown 清理與換行
    max_char = 45 # 每行約 45 個中文字
    
    clean_text = ai_advice.replace('**', '').replace('## ', '').replace('### ', '')
    
    for paragraph in clean_text.split('\n'):
        # 處理過長的段落
        while len(paragraph) > 0:
            line = paragraph[:max_char]
            paragraph = paragraph[max_char:]
            text_object.textLine(line)
            
            # 換頁檢查
            if text_object.getY() < 3*cm:
                c.drawText(text_object)
                c.showPage() # 換頁
                # 新頁面設定
                text_object = c.beginText(2*cm, height - 3*cm)
                text_object.setFont("NotoSans", 11)
                text_object.setLeading(16)
                
    c.drawText(text_object)
    c.save()
    buffer.seek(0)
    return buffer

# --- 4. 介面互動邏輯 ---

# 側邊欄：輸入 API Key
with st.sidebar:
    st.header("🔑 設定")
    user_api_key = st.text_input("請輸入 Gemini API Key", type="password", help="請前往 Google AI Studio 申請")
    st.markdown("---")
    st.info("💡 提示：Excel 檔名建議為「五科數據.xlsx」，且必須包含 國文, 英文, 數學, 社會, 自然 五個分頁。")

# 主畫面：上傳檔案
uploaded_file = st.file_uploader("📂 上傳 Excel 檔案 (.xlsx)", type=['xlsx'])

if uploaded_file and user_api_key:
    if st.button("🚀 開始生成全班報告"):
        
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
                    
                    # AI 生成
                    advice = get_ai_advice(user_api_key, student, str(errors))
                    
                    # PDF 生成
                    pdf_data = create_pdf(student, advice)
                    
                    # 加入 ZIP
                    zf.writestr(f"{student}_讀書建議.pdf", pdf_data.getvalue())
            
            progress_bar.progress(100)
            status_text.success("✅ 生成完成！")
            
            # 3. 下載按鈕
            st.download_button(
                label="📥 下載全班報告 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name="全班讀書建議報告.zip",
                mime="application/zip"
            )

elif uploaded_file and not user_api_key:
    st.warning("請在左側輸入 API Key 才能開始執行。")