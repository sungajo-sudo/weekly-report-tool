import streamlit as st
import pandas as pd
import pdfplumber
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

# --- 1. PDF 데이터 추출 함수 ---
def extract_data_from_pdf(pdf_file):
    all_data = []
    with pdfplumber.open(pdf_file) as pdf:
        # 첫 번째 페이지 분석
        page = pdf.pages[0]
        table = page.extract_table()
        
        if not table:
            return None
        
        # 데이터프레임으로 변환
        df_raw = pd.DataFrame(table)
        
        # 헤더 행 찾기 ('프로젝트' 또는 '업무' 키워드 기준)
        header_idx = -1
        for i, row in df_raw.iterrows():
            row_str = [str(cell) for cell in row if cell]
            if any('프로젝트' in s or '업무' in s for s in row_str):
                header_idx = i
                break
        
        if header_idx == -1:
            return None
            
        # 데이터 영역 추출
        data_rows = df_raw.iloc[header_idx + 1:]
        
        # PDF 표 구조 분석 (7개 컬럼 가정: 0,1,2(이번주) / 3(공백) / 4,5,6(다음주))
        # 만약 컬럼 수가 다르면 아래 인덱스를 조정합니다.
        col_count = len(df_raw.columns)
        
        this_week_list = []
        next_week_list = []
        
        for _, row in data_rows.iterrows():
            # 이번 주 데이터 (컬럼 0:팀원, 1:프로젝트, 2:내용)
            if row[1] and row[2]:
                this_week_list.append({'프로젝트': str(row[1]).strip(), '내용': str(row[2]).strip()})
            # 다음 주 데이터 (컬럼 4:팀원, 5:프로젝트, 6:내용)
            if col_count > 5 and row[5] and row[6]:
                next_week_list.append({'프로젝트': str(row[5]).strip(), '내용': str(row[6]).strip()})

        # 데이터 클리닝 및 그룹화 함수
        def clean_and_group(data_list):
            if not data_list:
                return pd.DataFrame(columns=['프로젝트명', '업무내용'])
            
            df = pd.DataFrame(data_list)
            # 불필요한 텍스트 및 중복 제거
            df = df[~df['프로젝트'].str.lower().isin(['nan', 'none', '', '프로젝트'])]
            df = df.drop_duplicates()
            
            # 프로젝트별 통합
            grouped = df.groupby('프로젝트')['내용'].apply(
                lambda x: "\n".join([f"• {val.replace('\\n', ' ')}" for val in x if val])
            ).reset_index()
            grouped.columns = ['프로젝트명', '업무내용']
            return grouped

        summary_this = clean_and_group(this_week_list)
        summary_next = clean_and_group(next_week_list)

        # 금주/차주 통합
        merged = pd.merge(summary_this, summary_next, on='프로젝트명', how='outer', suffixes=('_이번', '_다음'))
        return merged.fillna("-").sort_values('프로젝트명')

# --- 2. PPT 생성 함수 ---
def create_pptx(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    p = title_box.text_frame.add_paragraph()
    p.text = "서비스기획팀 주간업무보고"
    p.font.bold, p.font.size = True, Pt(30)

    # 표 생성
    rows, cols = len(df) + 1, 3
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.3), Inches(12.3), Inches(0.8)).table
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(2.3), Inches(5.0), Inches(5.0)

    # 헤더 스타일
    headers = ["프로젝트명", "이번 주 업무내용", "다음 주 업무내용"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(44, 62, 80)
        p = cell.text_frame.paragraphs[0]
        p.font.color.rgb, p.font.bold, p.font.size = RGBColor(255, 255, 255), True, Pt(16)
        p.alignment = PP_ALIGN.CENTER

    # 데이터 입력
    for i, row in df.iterrows():
        for j in range(3):
            cell = table.cell(i+1, j)
            cell.text = str(row.iloc[j])
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(11)
                para.alignment = PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- 3. UI ---
st.set_page_config(page_title="PDF to PPT Converter", layout="wide")
st.title("📄 PDF 주간보고 PPT 변환기")
st.info("PDF 파일의 왼쪽 표(이번 주)와 오른쪽 표(다음 주)를 자동으로 인식하여 취합합니다.")

uploaded_pdf = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])

if uploaded_pdf:
    with st.spinner("PDF 표 데이터를 분석 중..."):
        final_df = extract_data_from_pdf(uploaded_pdf)
        
        if final_df is not None:
            st.subheader("✅ 취합 데이터 확인")
            st.dataframe(final_df, use_container_width=True)

            if st.button("🚀 PPT 파일 생성 및 다운로드"):
                ppt_data = create_pptx(final_df)
                st.download_button(
                    label="📥 PPT 다운로드",
                    data=ppt_data,
                    file_name=f"주간보고_{uploaded_pdf.name.replace('.pdf', '')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
        else:
            st.error("PDF에서 표 형식을 찾을 수 없습니다. 파일의 텍스트가 추출 가능한 형태인지 확인해 주세요.")