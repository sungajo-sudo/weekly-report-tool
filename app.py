import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

# --- 1. 데이터 클리닝 및 헤더 자동 찾기 함수 ---
def get_clean_df(df):
    """상단의 빈 줄을 건너뛰고 실제 데이터 시작점(팀원/프로젝트 컬럼)을 찾습니다."""
    for i in range(len(df)):
        # 행의 값 중 '프로젝트'나 '팀원'이라는 글자가 포함된 행을 찾음
        row_values = [str(val) for val in df.iloc[i].values]
        if any('프로젝트' in val or '팀원' in val for val in row_values):
            new_df = df.iloc[i+1:].copy()
            new_df.columns = row_values
            return new_df.reset_index(drop=True)
    return df

# --- 2. 시트별 요약 함수 (중복 제거 포함) ---
def summarize_sheet(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=['프로젝트명', '내용'])
    
    # 헤더 정리
    df = get_clean_df(df)
    df.columns = [str(c).strip() for c in df.columns]
    
    # 필요한 컬럼 찾기
    proj_col = next((c for c in df.columns if '프로젝트' in c), None)
    task_col = next((c for c in df.columns if '업무' in c or '내용' in c), None)
    
    if not proj_col or not task_col:
        return pd.DataFrame(columns=['프로젝트명', '내용'])

    # 데이터 정리: 공백 제거, 결측치 제거
    df[proj_col] = df[proj_col].astype(str).str.strip()
    df[task_col] = df[task_col].astype(str).str.strip()
    df = df[df[proj_col].str.lower() != 'nan']
    df = df[df[task_col].str.lower() != 'nan']
    df = df[df[task_col] != '']

    # ★ 중복 내용 제거 (동일 프로젝트 내 같은 문구는 하나만 남김)
    df = df.drop_duplicates(subset=[proj_col, task_col])

    # 프로젝트별 통합
    summary = df.groupby(proj_col)[task_col].apply(
        lambda x: "\n".join([f"• {val}" for val in x])
    ).reset_index()
    
    summary.columns = ['프로젝트명', '내용']
    return summary

# --- 3. 메인 데이터 통합 로직 ---
def merge_data(uploaded_file):
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        # 금주/차주 시트 이름 매칭
        this_week_name = next((s for s in sheet_names if '금주' in s), None)
        next_week_name = next((s for s in sheet_names if '차주' in s), None)
        
        if not this_week_name or not next_week_name:
            st.error(f"시트 이름을 찾을 수 없습니다. (현재 시트: {sheet_names})")
            return None

        # 데이터 읽기
        df_this_raw = pd.read_excel(uploaded_file, sheet_name=this_week_name, header=None)
        df_next_raw = pd.read_excel(uploaded_file, sheet_name=next_week_name, header=None)
        
        # 시트별 요약 (팀원 제외, 중복 제거 적용)
        summary_this = summarize_sheet(df_this_raw)
        summary_next = summarize_sheet(df_next_raw)

        # 프로젝트 기준 통합
        merged = pd.merge(summary_this, summary_next, on='프로젝트명', how='outer', suffixes=('_금주', '_차주'))
        merged.columns = ['프로젝트명', '금주 업무내용', '차주 업무내용']
        return merged.fillna("-").sort_values('프로젝트명')

    except Exception as e:
        st.error(f"파일을 읽는 중 오류가 발생했습니다: {e}")
        return None

# --- 4. PPT 생성 함수 ---
def create_pptx(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(0.8))
    p = title_box.text_frame.add_paragraph()
    p.text = "서비스기획팀 주간업무보고"
    p.font.bold, p.font.size = True, Pt(28)

    # 표 (3열)
    rows, cols = len(df) + 1, 3
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.2), Inches(12.3), Inches(0.8)).table
    table.columns[0].width = Inches(2.3)
    table.columns[1].width = Inches(5.0)
    table.columns[2].width = Inches(5.0)

    # 헤더 스타일
    headers = ["프로젝트명", "금주 업무내용", "차주 업무내용"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(44, 62, 80)
        p = cell.text_frame.paragraphs[0]
        p.font.color.rgb, p.font.bold, p.font.size = RGBColor(255,255,255), True, Pt(15)
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

# --- 웹 UI ---
st.set_page_config(page_title="Weekly Report Tool", layout="wide")
st.title("📊 주간업무보고 PPT 생성기")

file = st.file_uploader("금주/차주 시트가 포함된 엑셀파일(.xlsx)을 업로드하세요", type=["xlsx"])

if file:
    with st.spinner("데이터 분석 중..."):
        merged_df = merge_data(file)
        
        if merged_df is not None:
            st.success("데이터를 성공적으로 통합했습니다.")
            st.dataframe(merged_df, use_container_width=True)
            
            if st.button("🪄 PPT 다운로드"):
                ppt_file = create_pptx(merged_df)
                st.download_button("📥 파일 받기", ppt_file, "주간업무보고.pptx")