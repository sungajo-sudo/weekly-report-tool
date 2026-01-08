import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io

# --- 1. 데이터 추출 및 클리닝 함수 ---
def parse_multi_column_sheet(df):
    """한 시트 내의 이번 주(좌측 3열) / 다음 주(우측 3열) 데이터를 분리하여 취합합니다."""
    
    # 실제 헤더('팀원', '프로젝트')가 있는 행 번호 찾기
    header_idx = -1
    for i in range(len(df)):
        row_values = [str(val).strip() for val in df.iloc[i].values]
        if '팀원' in row_values and '프로젝트' in row_values:
            header_idx = i
            break
    
    if header_idx == -1:
        st.error("파일에서 '팀원' 및 '프로젝트' 헤더를 찾을 수 없습니다. 양식을 확인해주세요.")
        return None

    # 데이터 시작 부분부터 슬라이싱
    data_df = df.iloc[header_idx + 1:].copy()
    
    # 열 인덱스 설정 (왼쪽: 0,1,2 / 오른쪽: 4,5,6)
    # 3번 열은 보통 비어있는 구분 열입니다.
    this_week_raw = data_df.iloc[:, [0, 1, 2]].copy()
    this_week_raw.columns = ['팀원', '프로젝트', '내용']
    
    next_week_raw = data_df.iloc[:, [4, 5, 6]].copy()
    next_week_raw.columns = ['팀원', '프로젝트', '내용']

    def clean_data(target_df):
        # 내용이 없는 행 제거 및 문자열 정리
        target_df = target_df.dropna(subset=['프로젝트', '내용'])
        target_df['프로젝트'] = target_df['프로젝트'].astype(str).str.strip()
        target_df['내용'] = target_df['내용'].astype(str).str.strip()
        
        # 유효하지 않은 값 필터링
        target_df = target_df[~target_df['프로젝트'].str.lower().isin(['nan', 'none', '', '프로젝트'])]
        target_df = target_df[~target_df['내용'].str.lower().isin(['nan', 'none', '', '주요 업무 내용'])]
        
        # ★ 중복 제거: 동일 프로젝트 내 완전히 같은 내용은 하나만 남김
        target_df = target_df.drop_duplicates(subset=['프로젝트', '내용'])
        
        # 프로젝트별 그룹화 (불렛 포인트 적용)
        grouped = target_df.groupby('프로젝트')['내용'].apply(
            lambda x: "\n".join([f"• {val}" for val in x if val])
        ).reset_index()
        return grouped

    summary_this = clean_data(this_week_raw)
    summary_next = clean_data(next_week_raw)

    # 두 표를 프로젝트명 기준으로 합침 (어느 한쪽만 있어도 표시되게 Outer Join)
    merged = pd.merge(summary_this, summary_next, on='프로젝트', how='outer', suffixes=('_이번', '_다음'))
    merged.columns = ['프로젝트명', '이번 주 업무내용', '다음 주 업무내용']
    
    # 빈 값은 대시(-)로 채우고 프로젝트명으로 정렬
    return merged.fillna("-").sort_values('프로젝트명')

# --- 2. PPT 생성 함수 ---
def create_pptx(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5) # 16:9
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 제목 상자 (이미지 양식 반영)
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.8))
    p = title_box.text_frame.add_paragraph()
    p.text = "서비스기획팀 주간업무보고"
    p.font.bold = True
    p.font.size = Pt(32)
    p.font.color.rgb = RGBColor(0, 0, 0)

    # 표 생성 (3열)
    rows, cols = len(df) + 1, 3
    left, top = Inches(0.5), Inches(1.3)
    width, height = Inches(12.3), Inches(0.6)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    # 열 너비 설정
    table.columns[0].width = Inches(2.3) # 프로젝트
    table.columns[1].width = Inches(5.0) # 이번 주
    table.columns[2].width = Inches(5.0) # 다음 주

    # 헤더 스타일 (진네이비 배경 + 흰색 글씨)
    headers = ["프로젝트명", "이번 주 업무내용", "다음 주 업무내용"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(44, 62, 80)
        para = cell.text_frame.paragraphs[0]
        para.font.color.rgb, para.font.bold, para.font.size = RGBColor(255, 255, 255), True, Pt(16)
        para.alignment = PP_ALIGN.CENTER

    # 데이터 입력
    for i, row in df.iterrows():
        for j in range(3):
            cell = table.cell(i+1, j)
            cell.text = str(row.iloc[j])
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(11)
                para.font.name = '맑은 고딕'
                # 프로젝트명은 중앙, 내용은 왼쪽 정렬
                para.alignment = PP_ALIGN.CENTER if j == 0 else PP_ALIGN.LEFT

    ppt_io = io.BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- 3. Streamlit UI 구성 ---
st.set_page_config(page_title="Weekly Report Tool", layout="wide")
st.title("📊 주간업무보고 PPT 자동 변환기")
st.write("엑셀의 첫 번째 시트에서 '이번 주'와 '다음 주' 데이터를 취합합니다.")

uploaded_file = st.file_uploader("파일을 업로드하세요 (.xlsx 또는 .csv)", type=["xlsx", "csv"])

if uploaded_file:
    try:
        # 파일 타입에 따른 로드
        if uploaded_file.name.endswith('.csv'):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        
        # 데이터 처리
        final_df = parse_multi_column_sheet(df_raw)
        
        if final_df is not None and not final_df.empty:
            st.subheader("✅ 취합된 데이터 미리보기")
            st.dataframe(final_df, use_container_width=True)

            # PPT 생성 (버튼 클릭 전 미리 생성하여 안정성 확보)
            ppt_data = create_pptx(final_df)
            
            st.download_button(
                label="📥 PPT 파일 다운로드",
                data=ppt_data,
                file_name=f"주간업무보고_{uploaded_file.name.split('.')[0]}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            st.success("데이터 취합이 완료되었습니다. 위 버튼을 눌러 다운로드하세요!")
        else:
            st.warning("분석할 수 있는 데이터가 없습니다. 시트의 구성을 확인해주세요.")

    except Exception as e:
        st.error(f"오류가 발생했습니다: {e}")