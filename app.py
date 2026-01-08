import streamlit as st
import pandas as pd
import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

# --- 1. 데이터 클리닝 및 취합 함수 ---
def clean_and_summarize(data_list):
    """프로젝트별로 중복을 제거하고 불렛 포인트로 묶습니다."""
    if not data_list:
        return pd.DataFrame(columns=['프로젝트', '내용'])
    
    df = pd.DataFrame(data_list)
    df.columns = ['팀원', '프로젝트', '내용']
    
    # 기본 전처리: 공백 제거 및 결측치 제거
    df = df.dropna(subset=['프로젝트', '내용'])
    df['프로젝트'] = df['프로젝트'].astype(str).str.strip()
    df['내용'] = df['내용'].astype(str).str.strip()
    
    # 유효하지 않은 값(헤더 반복 등) 필터링
    invalid_keywords = ['nan', 'none', '', '프로젝트', '주요 업무 내용', '주요업무내용']
    df = df[~df['프로젝트'].str.lower().isin(invalid_keywords)]
    df = df[~df['내용'].str.lower().isin(invalid_keywords)]
    
    # ★ 중복 제거: 동일 프로젝트 내 완전히 같은 업무 내용은 하나만 남김
    df = df.drop_duplicates(subset=['프로젝트', '내용'])
    
    # 프로젝트별 그룹화 (팀원 이름 제외)
    grouped = df.groupby('프로젝트')['내용'].apply(
        lambda x: "\n".join([f"• {val}" for val in x if val])
    ).reset_index()
    return grouped

# --- 2. PDF 분석 함수 ---
def parse_pdf(pdf_file):
    this_week_all = []
    next_week_all = []
    
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if not table: continue
            
            df_raw = pd.DataFrame(table)
            # 헤더 찾기
            header_idx = -1
            for i, row in df_raw.iterrows():
                row_str = "".join([str(c) for c in row if c])
                if '프로젝트' in row_str or '팀원' in row_str:
                    header_idx = i
                    break
            
            if header_idx == -1: continue
            
            data_rows = df_raw.iloc[header_idx + 1:]
            for _, row in data_rows.iterrows():
                # 좌측 3열: 이번 주 / 우측 3열: 다음 주 (중간 빈 칸 고려)
                if len(row) >= 3 and row[1] and row[2]:
                    this_week_all.append([row[0], row[1], row[2]])
                if len(row) >= 7 and row[5] and row[6]:
                    next_week_all.append([row[4], row[5], row[6]])
                    
    return clean_and_summarize(this_week_all), clean_and_summarize(next_week_all)

# --- 3. 엑셀 분석 함수 ---
def parse_excel(excel_file):
    df_raw = pd.read_excel(excel_file, sheet_name=0, header=None)
    
    header_idx = -1
    for i in range(len(df_raw)):
        row_values = [str(val).strip() for val in df_raw.iloc[i].values]
        if '프로젝트' in row_values or '팀원' in row_values:
            header_idx = i
            break
            
    if header_idx == -1: return None, None
    
    data_df = df_raw.iloc[header_idx + 1:].copy()
    this_week_raw = data_df.iloc[:, [0, 1, 2]].values.tolist()
    next_week_raw = data_df.iloc[:, [4, 5, 6]].values.tolist()
    
    return clean_and_summarize(this_week_raw), clean_and_summarize(next_week_raw)

# --- 4. PPT 생성 함수 ---
def create_pptx(merged_df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    p = title_box.text_frame.add_paragraph()
    p.text = "서비스기획팀 주간업무보고"
    p.font.bold, p.font.size = True, Pt(32)

    rows, cols = len(merged_df) + 1, 3
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.3), Inches(12.3), Inches(0.8)).table
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(2.3), Inches(5.0), Inches(5.0)

    headers = ["프로젝트명", "이번 주 업무내용", "다음 주 업무내용"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(44, 62, 80)
        para = cell.text_frame.paragraphs[0]
        para.font.color.rgb, para.font.bold, para.font.size = RGBColor(255, 255, 255), True, Pt(16)
        para.alignment = PP_ALIGN.CENTER

    for i, row in merged_df.iterrows():
        for j in range(3):
            cell = table.cell(i+1, j)
            cell.text = str(row.iloc[j])
            for para in cell.text_frame.paragraphs:
                para.font.size, para.font.name = Pt(11), '맑은 고딕'
                para.alignment = PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- 5. Streamlit UI ---
st.set_page_config(page_title="Weekly Report Converter", layout="wide")
st.title("📊 주간보고 통합 변환기 (PDF/Excel 지원)")
st.write("PDF 또는 엑셀 파일을 업로드하면 프로젝트별로 자동 취합하여 PPT를 생성합니다.")

uploaded_file = st.file_uploader("파일을 업로드하세요", type=["xlsx", "pdf", "csv"])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.pdf'):
            sum_this, sum_next = parse_pdf(uploaded_file)
        else:
            sum_this, sum_next = parse_excel(uploaded_file)
        
        if sum_this is not None:
            # 병합
            merged = pd.merge(sum_this, sum_next, on='프로젝트', how='outer', suffixes=('_이번', '_다음'))
            merged.columns = ['프로젝트명', '이번 주 업무내용', '다음 주 업무내용']
            merged = merged.fillna("-").sort_values('프로젝트명')
            
            st.subheader("✅ 데이터 취합 결과 확인")
            st.dataframe(merged, use_container_width=True)
            
            ppt_data = create_pptx(merged)
            st.download_button(
                label="📥 PPT 파일 다운로드",
                data=ppt_data,
                file_name=f"주간보고_통합_{uploaded_file.name.split('.')[0]}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            st.success("데이터 취합이 완료되었습니다!")
        else:
            st.error("데이터를 분석할 수 없습니다. 파일 양식을 확인해 주세요.")
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")