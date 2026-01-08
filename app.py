import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

# --- 1. 데이터 추출 및 클리닝 함수 ---
def parse_parallel_columns(df):
    """한 시트 내의 이번 주(좌측 0-2열) / 다음 주(우측 4-6열) 데이터를 추출합니다."""
    
    # 헤더('팀원', '프로젝트')가 있는 행 찾기
    header_idx = -1
    for i in range(len(df)):
        row_values = [str(val).strip() for val in df.iloc[i].values]
        if '팀원' in row_values and '프로젝트' in row_values:
            header_idx = i
            break
    
    if header_idx == -1:
        st.error("파일에서 '팀원' 및 '프로젝트' 헤더를 찾을 수 없습니다.")
        return None

    # 데이터 영역 슬라이싱
    data_df = df.iloc[header_idx + 1:].copy()
    
    # 0,1,2열 -> 이번 주 / 4,5,6열 -> 다음 주
    this_week_raw = data_df.iloc[:, [0, 1, 2]].copy()
    this_week_raw.columns = ['팀원', '프로젝트', '내용']
    
    next_week_raw = data_df.iloc[:, [4, 5, 6]].copy()
    next_week_raw.columns = ['팀원', '프로젝트', '내용']

    def clean_and_summarize(target_df):
        target_df = target_df.dropna(subset=['프로젝트', '내용'])
        target_df['프로젝트'] = target_df['프로젝트'].astype(str).str.strip()
        target_df['내용'] = target_df['내용'].astype(str).str.strip()
        
        # 유효하지 않은 행 제거
        target_df = target_df[~target_df['프로젝트'].str.lower().isin(['nan', 'none', '', '프로젝트'])]
        target_df = target_df[~target_df['내용'].str.lower().isin(['nan', 'none', '', '주요 업무 내용'])]
        
        # ★ 중복 제거: 동일 프로젝트 내 같은 내용은 하나만 남김
        target_df = target_df.drop_duplicates(subset=['프로젝트', '내용'])
        
        # 프로젝트별 그룹화 (팀원 제외)
        return target_df.groupby('프로젝트')['내용'].apply(
            lambda x: "\n".join([f"• {val}" for val in x if val])
        ).reset_index()

    summary_this = clean_and_summarize(this_week_raw)
    summary_next = clean_and_summarize(next_week_raw)

    # 프로젝트명 기준 병합
    merged = pd.merge(summary_this, summary_next, on='프로젝트', how='outer', suffixes=('_이번', '_다음'))
    merged.columns = ['프로젝트명', '이번 주 업무내용', '다음 주 업무내용']
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
    p.font.bold, p.font.size = True, Pt(32)

    # 표 구성
    rows, cols = len(df) + 1, 3
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.3), Inches(12.3), Inches(0.8)).table
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(2.3), Inches(5.0), Inches(5.0)

    # 헤더 디자인
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
                para.font.size, para.font.name = Pt(11), '맑은 고딕'
                para.alignment = PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- 3. Streamlit UI ---
st.set_page_config(page_title="Weekly Report Converter", layout="wide")
st.title("📊 주간업무보고 PPT 생성 도구")
st.write("구글 드라이브의 파일을 PC로 다운로드한 뒤 아래에 업로드해주세요.")

file = st.file_uploader("파일 업로드 (.xlsx, .csv)", type=["xlsx", "csv"])

if file:
    try:
        # 데이터 읽기
        df_raw = pd.read_csv(file, header=None) if file.name.endswith('.csv') else pd.read_excel(file, header=None)
        final_df = parse_parallel_columns(df_raw)
        
        if final_df is not None:
            st.subheader("✅ 취합 데이터 확인")
            st.dataframe(final_df, use_container_width=True)

            # 다운로드 버튼
            ppt_data = create_pptx(final_df)
            st.download_button(
                label="📥 PPT 파일 다운로드 (클릭)",
                data=ppt_data,
                file_name=f"서비스기획팀_주간보고_{file.name.split('.')[0]}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")