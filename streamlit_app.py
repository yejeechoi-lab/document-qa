import streamlit as st
import pandas as pd
import re
import io
from io import BytesIO
import xlsxwriter
import gdown
import os
import cv2

# 스트림릿 페이지 설정 (가로를 넓게 쓰기 위함)
st.set_page_config(layout="wide")

st.title("🎬 소재 자동 인덱싱 및 정형화 대시보드")
st.write("구글 드라이브 링크를 입력하고 원하는 열(가로 칸) 개수를 설정하면, 예쁘게 정렬된 엑셀 리포트를 생성합니다.")

# --- 가로 배치 열 개수 조절 슬라이더 ---
columns_count = st.slider("📊 가로로 배치할 소재(열) 개수를 설정하세요:", min_value=2, max_value=15, value=10, step=1)

# 구글 드라이브 폴더 ID 추출 함수
def extract_folder_id(url):
    match = re.search(r"folders/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    return None

# 이미지 다운로드 및 설정된 열 개수 배열 엑셀 생성 핵심 로직
def create_indexed_excel(folder_url, cols_cnt):
    folder_id = extract_folder_id(folder_url)
    if not folder_id:
        st.error("올바른 구글 드라이브 폴더 주소가 아닙니다. 다시 확인해 주세요.")
        return None

    # 다운로드용 임시 폴더 빌드
    output_dir = "drive_download"
    try:
        with st.spinner("구글 드라이브에서 소재를 안전하게 가져오는 중..."):
            if os.path.exists(output_dir):
                import shutil
                shutil.rmtree(output_dir)
            
            # 어제 쓰던 안전한 순정 다운로드 방식
            gdown.download_folder(id=folder_id, output=output_dir, quiet=True)
    except Exception as e:
        st.error(f"구글 드라이브 연결 실패: {e}\n폴더가 '링크가 있는 모든 사용자'에게 공개되어 있는지 확인해 주세요.")
        return None

    # 엑셀 생성을 위한 메모리 버퍼 생성
    output_excel = io.BytesIO()
    workbook = xlsxwriter.Workbook(output_excel, {'in_memory': True})
    worksheet = workbook.add_worksheet("소재인덱싱")
    
    # 엑셀 스타일 설정
    header_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 
        'bg_color': '#DCE6F1', 'border': 1
    })
    text_format = workbook.add_format({
        'align': 'center', 'valign': 'vcenter', 'border': 1
    })

    # 설정된 열 개수만큼 칸 너비 조절 (15 지정)
    for col_idx in range(cols_cnt):
        worksheet.set_column(col_idx, col_idx, 15)

    # 이미지 파일만 필터링해서 가져오기
    all_files = os.listdir(output_dir)
    img_files = [f for f in all_files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.webp'))]

    if not img_files:
        st.warning("폴더 내에 이미지 파일(.png, .jpg 등)이 발견되지 않았습니다.")
        workbook.close()
        return None

    row_pointer = 0
    
    # 사용자가 설정한 cols_cnt 개수씩 쪼개서 이미지 배치 시작
    for i in range(0, len(img_files), cols_cnt):
        chunk = img_files[i:i+cols_cnt]
        
        # 1층: 파일명 작성 행
        worksheet.set_row(row_pointer, 25)
        for c_idx, filename in enumerate(chunk):
            display_name = os.path.splitext(filename)[0]
            worksheet.write(row_pointer, c_idx, display_name, header_format)
        # 빈 칸도 테두리 채우기
        for c_idx in range(len(chunk), cols_cnt):
            worksheet.write(row_pointer, c_idx, "", text_format)
            
        # 2층: 실제 이미지 삽입 행
        worksheet.set_row(row_pointer + 1, 100) # 이미지 행 높이 100으로 고정
        for c_idx, filename in enumerate(chunk):
            file_path = os.path.join(output_dir, filename)
            
            try:
                img = cv2.imread(file_path)
                if img is not None:
                    is_success, buffer = cv2.imencode(".png", img)
                    if is_success:
                        img_data = BytesIO(buffer)
                        worksheet.insert_image(
                            row_pointer + 1, c_idx, filename,
                            {
                                'image_data': img_data,
                                'x_scale': 0.14,
                                'y_scale': 0.14,
                                'x_offset': 5,
                                'y_offset': 5,
                                'object_position': 2 # 셀 크기에 맞춰 고정
                            }
                        )
                    else:
                        worksheet.write(row_pointer + 1, c_idx, "(변환 실패)", text_format)
                else:
                    worksheet.write(row_pointer + 1, c_idx, "(읽기 불가)", text_format)
            except:
                worksheet.write(row_pointer + 1, c_idx, "(에러)", text_format)
                
        # 빈 이미지 칸 테두리 처리
        for c_idx in range(len(chunk), cols_cnt):
            worksheet.write(row_pointer + 1, c_idx, "", text_format)
            
        row_pointer += 3 # 다음 세트를 위해 3줄 아래로 이동 (공백 1줄 포함)

    workbook.close()
    output_excel.seek(0)
    return output_excel

# 대시보드 UI 화면단
folder_url_input = st.text_input("🔗 구글 드라이브 폴더 링크를 입력하세요:", "")

if folder_url_input:
    excel_data = create_indexed_excel(folder_url_input, columns_count)
    
    if excel_data:
        st.success(f"🎉 소재 리포트 생성 완료! (가로 배열: {columns_count}칸)")
        st.download_button(
            label="📊 완벽 정형화된 엑셀 파일 PC로 저장하기",
            data=excel_data,
            file_name="소재인덱싱_리포트.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
