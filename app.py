import streamlit as st
import pandas as pd
import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re

# --- 1. 텍스트 간결화 및 중복 제거 함수 ---
def refine_text(text):
    if not text or text == "-": return "-"
    
    # 불필요한 공백 및 반복 기호 정리
    lines = text.split('\n')
    refined_lines = []
    seen = set()

    for line in lines:
        # 불필요한 수식어 제거 및 문구 간결화 (예: 진행 중입니다 -> 진행)
        line = line.strip().replace('•', '').strip()
        line = re.sub(r' 진행 중(입니다)?', ' 진행', line)
        line = re.sub(r' 완료(하였습니다|했습니다)?', ' 완료', line)
        line = re.sub(r' 예정(입니다)?', ' 예정', line)
        line = line.replace(' 팔로업', ' F/U').replace('팔로우업', ' F/U')

        # 중복 라인 제거
        if line and line not in seen:
            refined_lines.append(f"• {line}")
            seen.add(line)
            
    return "\n".join(refined_lines) if refined_lines else "-"

# --- 2. 통합 데이터 처리 (Excel/PDF 공용) ---
def process_report_data(file):
    if file.name.endswith('.pdf'):
        this_week, next_week = [], []
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if not table: continue
                for row in table:
                    # 좌측 이번주(0,1,2) / 우측 다음주(4,5,6)
                    if len(row) >= 3 and row[1] and row[2]: this_week.append([row[0], row[1], row[2]])
                    if len(row) >= 7 and row[5] and row[6]: next_week.append([row[4], row[5], row[6]])
    else:
        df_raw = pd.read_excel(file, sheet_name=0, header=None)
        this_week, next_week = [], []
        # 헤더 찾기
        h_idx = -1
        for i in range(len(df_raw)):
            row = [str(v) for v in df_raw.iloc[i].values]
            if '프로젝트' in row: h_idx = i; break
        
        data_df = df_raw.iloc[h_idx + 1:]
        for _, r in data_df.iterrows():
            if len(r) >= 3: this_week.append([r[0], r[1], r[2]])
            if len(r) >= 7: next_week.append([r[4], r[5], r[6]])

    def summarize(rows):
        df = pd.DataFrame(rows, columns=['팀원', '프로젝트', '내용']).dropna(subset=['프로젝트', '내용'])
        df['프로젝트'] = df['프로젝트'].astype(str).str.strip()
        df = df[~df['프로젝트'].str.contains('프로젝트|팀원|nan', case=False)]
        # 그룹화 및 텍스트 정제 적용
        grouped = df.groupby('프로젝트')['내용'].apply(lambda x: refine_text("\n".join(x))).reset_index()
        return grouped

    res_this = summarize(this_week)
    res_next = summarize(next_week)
    
    merged = pd.merge(res_this, res_next, on='프로젝트', how='outer', suffixes=('_금', '_차')).fillna("-")
    merged.columns = ['프로젝트명', '이번 주 업무내용', '다음 주 업무내용']
    return merged.sort_values('프로젝트명')

# --- 3. PPT 생성 함수 (자동 페이지 분할 기능 포함) ---
def create_split_pptx(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    
    # 한 페이지에 담을 최대 프로젝트(행) 수
    ROWS_PER_PAGE = 5 
    
    # 데이터프레임을 묶음으로 나누기
    for i in range(0, len(df), ROWS_PER_PAGE):
        chunk = df.iloc[i : i + ROWS_PER_PAGE]
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 제목
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
        p = title_box.text_frame.add_paragraph()
        p.text = f"서비스기획팀 주간업무보고 ({i//ROWS_PER_PAGE + 1})"
        p.font.bold, p.font.size = True, Pt(28)

        # 표 생성
        table = slide.shapes.add_table(len(chunk) + 1, 3, Inches(0.5), Inches(1.3), Inches(12.3), Inches(0.8)).table
        table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(2.3), Inches(5.0), Inches(5.0)

        # 헤더 디자인
        headers = ["프로젝트명", "지난 주 진행(MM월 YY주차)", "금주 계획(MM월 YY주차)"]
        for j, h in enumerate(headers):
            cell = table.cell(0, j)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(44, 62, 80)
            para = cell.text_frame.paragraphs[0]
            para.font.color.rgb, para.font.bold, para.font.size = RGBColor(255,255,255), True, Pt(15)
            para.alignment = PP_ALIGN.CENTER

        # 데이터 입력
        for row_idx, (_, data) in enumerate(chunk.iterrows()):
            for col_idx in range(3):
                cell = table.cell(row_idx + 1, col_idx)
                cell.text = str(data.iloc[col_idx])
                for p in cell.text_frame.paragraphs:
                    p.font.size, p.font.name = Pt(11), '맑은 고딕'
                    p.alignment = PP_ALIGN.CENTER if col_idx == 0 else PP_ALIGN.LEFT

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    return ppt_io.getvalue()

# --- 4. Streamlit UI ---
st.set_page_config(page_title="Weekly Report Smart Converter", layout="wide")
st.title("🚀 주간보고 스마트 PPT 변환기")
st.markdown("내용을 **간결하게 요약**하고, 양이 많으면 **슬라이드를 자동으로 분할**합니다.")

file = st.file_uploader("Excel 또는 PDF 파일을 업로드하세요", type=["xlsx", "pdf"])

if file:
    with st.spinner("데이터 정제 및 PPT 생성 중..."):
        final_df = process_report_data(file)
        st.subheader("✅ 정제된 데이터 미리보기")
        st.dataframe(final_df, use_container_width=True)
        
        ppt_binary = create_split_pptx(final_df)
        st.download_button(
            label="📥 정제된 PPT 다운로드",
            data=ppt_binary,
            file_name=f"주간보고_정제본_{file.name.split('.')[0]}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )