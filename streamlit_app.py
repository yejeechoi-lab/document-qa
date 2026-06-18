import streamlit as st
import os
import cv2
import xlsxwriter
import unicodedata
import tempfile
import gdown
from io import BytesIO

# 페이지 기본 설정
st.set_page_config(page_title='소재 인덱싱 자동화 대시보드 (구글 드라이브 버전)', page_icon='✨', layout='centered')

st.markdown("<p class='main-title'>✨ 소재 인덱싱 자동화 대시보드</p>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>대용량 영상 업로드 없이, 구글 드라이브 폴더 링크만 넣으면 자모음 깨짐 없는 엑셀을 즉시 생성합니다.</p>", unsafe_allow_html=True)

# CSS 스타일 적용
st.markdown("""
    <style>
    .main-title { font-size:30px; font-weight:bold; color:#2E4A3F; margin-bottom:5px; }
    .sub-title { font-size:14px; color:#666666; margin-bottom:25px; }
    div.stButton > button:first-child { background-color: #2E4A3F; color: white; border-radius: 6px; font-weight: bold; width: 100%; height: 45px; border: none; }
    div.stButton > button:first-child:hover { background-color: #1F332B; color: white; }
    </style>
""", unsafe_allow_html=True)

# 1. 구글 드라이브 폴더 URL 입력 받기
folder_url = st.text_input("🔗 구글 드라이브 폴더 공유 링크를 입력해주세요", placeholder="https://drive.google.com/drive/folders/...")

# 2. 설정 옵션들 (⭐ 가로 바둑판 최대 10개까지 대폭 확장!)
col1, col2 = st.columns(2)
with col1:
    # 3개부터 10개까지 마케터님이 원하시는 대로 다 고르실 수 있습니다.
    grid_cols = st.selectbox("📊 가로 바둑판 배열 개수", [3, 4, 5, 6, 7, 8, 9, 10], index=0)
with col2:
    excel_name = st.text_input("📝 저장될 엑셀 파일명", value="소재_자동_인덱싱_리포트")

# 자모음 분리 해결 함수
def fix_normalization(text):
    return unicodedata.normalize('NFC', text)

# 영상 캡처 함수
def extract_frame(video_path):
    cap = cv2.VideoCapture(video_path)
    success, frame = cap.read()
    if success:
        _, img_encoded = cv2.imencode('.jpg', frame)
        cap.release()
        return img_encoded.tobytes()
    cap.release()
    return None

# 실행 버튼 클릭 시 구동
if st.button("🚀 엑셀 리포트 추출하기"):
    if not folder_url:
        st.warning("구글 드라이브 폴더 링크를 입력해주세요!")
    else:
        with st.spinner("구글 드라이브에서 파일을 안전하게 가져와 연산하는 중입니다... ☕"):
            try:
                # 임시 디렉토리 생성하여 구글 드라이브 다운로드 진행
                with tempfile.TemporaryDirectory() as tmpdir:
                    downloaded_files = gdown.download_folder(url=folder_url, output=tmpdir, quiet=True)
                    
                    if not downloaded_files:
                        st.error("폴더에 접근할 수 없거나 파일이 없습니다. 구글 드라이브 폴더가 '링크가 있는 모든 사용자에게 공개' 상태인지 확인해주세요!")
                    else:
                        output_excel = f"{excel_name}.xlsx"
                        workbook = xlsxwriter.Workbook(output_excel)
                        worksheet = workbook.add_worksheet()
                        
                        # 🎨 테두리와 투명 배경 세팅
                        text_format = workbook.add_format({
                            'align': 'center', 
                            'valign': 'vcenter', 
                            'text_wrap': True, 
                            'font_size': 9,           # 글자 크기 9pt 고정
                            'border': 1,
                            'border_color': '#D3D3D3' # 투명하고 옅은 회색 테두리
                        })
                        
                        # 파일 목록 정렬
                        local_files = sorted([f for f in os.listdir(tmpdir) if os.path.isfile(os.path.join(tmpdir, f))])
                        col_count = grid_cols     
                        
                        # 모든 열의 너비를 깔끔하게 통일 (너비 22)
                        for c in range(col_count):
                            worksheet.set_column(c, c, 22)
                        
                        current_file_idx = 0
                        row_pointer = 0
                        
                        while current_file_idx < len(local_files):
                            # ① 상단 소재명 행: 높이 20으로 아주 슬림하게
                            worksheet.set_row(row_pointer, 20) 
                            
                            # ② 하단 이미지 행: 높이 120으로 안정적으로 삽입
                            worksheet.set_row(row_pointer + 1, 120) 
                            
                            for c in range(col_count):
                                if current_file_idx >= len(local_files):
                                    break
                                    
                                filename = local_files[current_file_idx]
                                file_path = os.path.join(tmpdir, filename)
                                
                                display_name = os.path.splitext(filename)[0]
                                
                                # [상단 셀] 배경색 없이 투명하게 글자 입력
                                worksheet.write(row_pointer, c, fix_normalization(display_name), text_format)
                                
                                # [하단 셀] 이미지를 완전히 셀에 꽉 가두는 옵션
                                img_data = None
                                if filename.lower().endswith(('.mp4', '.mov', '.avi', '.mkv')):
                                    img_data = extract_frame(file_path)
                                elif filename.lower().endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
                                    with open(file_path, 'rb') as f:
                                        img_data = f.read()
                                        
                                if img_data:
                                    try:
                                        # 비어있는 이미지 셀 공간에도 테두리를 입혀서 격자를 완성해 줍니다.
                                        worksheet.write(row_pointer + 1, c, "", text_format)
                                        
                                        # ⭐ [핵심 변경] object_position 옵션을 2로 주어 셀 크기 변화에 맞춰 이미지가 무조건 내부에 박히도록 고정합니다.
                                        worksheet.insert_image(
                                            row_pointer + 1, c, filename, 
                                            {
                                                'image_data': BytesIO(img_data),
                                                'x_scale': 0.14,  
                                                'y_scale': 0.14, 
                                                'x_offset': 5,    
                                                'y_offset': 5,
                                                'object_position': 2  # 2번 옵션: Move and size with cells (셀 안에 고정되어 함께 움직임)
                                            }
                                        )
                                    except Exception as e:
                                        worksheet.write(row_pointer + 1, c, "(이미지 삽입 실패)", text_format)
                                else:
                                    worksheet.write(row_pointer + 1, c, "(미리보기 불가)", text_format)
                                    
                                current_file_idx += 1
                                
                            row_pointer += 2
                        
                        workbook.close()
                        
                        with open(output_excel, 'rb') as f:
                            st.success("🎉 대폭 업그레이드된 최종 양식 리포트 생성 완료!")
                            st.download_button(
                                label="📥 완벽 정형화된 엑셀 파일 PC로 저장하기",
                                data=f,
                                file_name=f"{excel_name}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
