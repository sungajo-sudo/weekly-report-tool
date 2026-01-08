import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import io
import re
import requests

# --- 1. 데이터 추출 및 클리닝 함수 (기존 로직 유지) ---
def parse_multi_column_sheet(df):
    header_idx = -1
    for i in range(len(df)):
        row_str = [str(val) for val in df.iloc[i].values]
        if any('팀원' in s or '프로젝트' in s for s in row_str):
            header_idx = i
            break
    
    if header_idx == -1:
        return None

    data_df = df.iloc[header_idx + 1:].copy()
    
    # 0,1,2열 -> 이번 주 / 4,5,6열 -> 다음 주
    this_week_raw = data_df.iloc[:, [0, 1, 2]]
    this_week_raw.columns = ['팀원', '프로젝트', '내용']
    next_week_raw = data_df.iloc[:, [4, 5, 6]]
    next_week_raw.columns = ['팀원', '프로젝트', '내용']

    def clean_and_group(target_df):
        target_df = target_df.dropna(subset=['프로젝트', '내용'])
        target_df['프로젝트'] = target_df['프로젝트'].astype(str).str.strip()
        target_df['내용'] = target_df['내용'].astype(str).str.strip()
        target_df = target_df[~target_df['프로젝트'].str.lower().isin(['nan', 'none', ''])]
        target_df = target_df.drop_duplicates(subset=['프로젝트', '내용'])
        
        return target_df.groupby('프로젝트')['내용'].apply(
            lambda x: "\n".join([f"• {val}" for v in x if (val := str(v).strip())])
        ).reset_index()

    summary_this = clean_and_group(this_week_raw)
    summary_next = clean_and_group(next_week_raw)

    merged = pd.merge(summary_this, summary_next, on='프로젝트', how='outer', suffixes=('_이번주', '_다음주'))
    merged.columns = ['프로젝트명', '이번 주 업무내용', '다음 주 업무내용']
    return merged.fillna("-").sort_values('프로젝트명')

# --- 2. 구글 드라이브 파일 다운로드 함수 ---
def download_from_drive(url):
    """공유된 구글 드라이브 링크에서 파일을 다운로드합니다."""
    try:
        # 파일 ID 추출
        file_id_match = re.search(r'd/([^/]+)', url)
        if not file_id_match:
            st.error("올바른 구글 드라이브 링크가 아닙니다.")
            return None
        
        file_id = file_id_match.group(1)
        # 구글 드라이브 직다운로드 URL (CSV로 내보내기 방식)
        download_url = f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
        
        response = requests.get(download_url)
        if response.status_code == 200:
            return io.BytesIO(response.content)
        else:
            st.error("파일을 불러올 수 없습니다. 링크가 '링크가 있는 모든 사용자에게 공개' 상태인지 확인하세요.")
            return None
    except Exception as e:
        st.error(f"드라이브 연결 오류: {e}")
        return None

# --- 3. PPT 생성 함수 (기존 로직 유지) ---
def create_pptx(df):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.33), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12), Inches(0.8))
    p = title_box.text_frame.add_paragraph()
    p.text = "서비스기획팀 주간업무보고"
    p.font.bold, p.font.size = True, Pt(28)

    rows, cols = len(df) + 1, 3
    table = slide.shapes.add_table(rows, cols, Inches(0.5), Inches(1.3), Inches(12.3), Inches(0.8)).table
    table.columns[0].width, table.columns[1].width, table.columns[2].width = Inches(2.3), Inches(5.0), Inches(5.0)

    headers = ["프로젝트명", "이번 주 업무내용", "다음 주 업무내용"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(44, 62, 80)
        p = cell.text_frame.paragraphs[0]
        p.font.color.rgb, p.font.bold, p.font.size = RGBColor(255, 255, 255), True, Pt(16)
        p.alignment = PP_ALIGN.CENTER

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

# --- Streamlit UI ---
st.set_page_config(page_title="Weekly Report Tool", layout="wide")
st.title("📊 주간보고 PPT 자동 변환기")

# 입력 방식 선택
option = st.radio("파일 선택 방식", ["내 컴퓨터에서 업로드", "구글 드라이브 링크로 가져오기"])

input_file = None

if option == "내 컴퓨터에서 업로드":
    input_file = st.file_uploader("엑셀 파일을 업로드하세요", type=["xlsx"])
else:
    drive_url = st.text_input("구글 스프레드시트 공유 링크를 입력하세요", placeholder="https://docs.google.com/spreadsheets/d/...")
    if drive_url:
        input_file = download_from_drive(drive_url)

if input_file:
    try:
        df_raw = pd.read_excel(input_file, sheet_name=0, header=None)
        merged_df = parse_multi_column_sheet(df_raw)
        
        if merged_df is not None:
            st.success("데이터 취합 성공!")
            st.dataframe(merged_df, use_container_width=True)
            
            if st.button("🪄 PPT 생성 및 다운로드"):
                ppt_file = create_pptx(merged_df)
                st.download_button("📥 PPT 받기", ppt_file, "주간업무보고.pptx")
    except Exception as e:
        st.error(f"오류: {e}")