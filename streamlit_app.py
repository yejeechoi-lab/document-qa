import streamlit as st# 🚨 [용량 제한 완벽 해제] 업로드 제한을 1000MB(1GB)로 강제 설정합니다.
import streamlit.components.v1 as components
st.config.set_option("server.maxUploadSize", 1000)
import os
import re
import cv2
import xlsxwriter
import unicodedata
from io import BytesIO
import tempfile

# 페이지 설정
st.set_page_config(page_title='소재 인덱싱 자동화 대시보드', page_icon='✨', layout='centered')

st.markdown("""
    <style>
    .main-title { font-size:30px; font-weight:bold; color:#2E4A3F; margin-bottom:5px; }
    .sub-title { font-size:14px; color:#666666; margin-bottom:25px; }
    div.stButton > button:first-child {
        background-color: #2E4A3F; color: white; border-radius: 6px;
        font-weight: bold; width: 100%; height: 45px; border: none;
    }
    div.stButton > button:first-child:hover { background-color: #1F332B; color: white; }
    </style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">✨ 소재 인덱싱 자동화 대시보드</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-title">광고 소재들을 업로드하면 자모음 깨짐 없는 바둑판 엑셀을 즉시 생성합니다.</p>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "광고 소재 파일들을 한 번에 선택하여 올려주세요 (이미지 및 영상 가능)", 
    accept_multiple_files=True,
    type=['jpg', 'jpeg', 'png', 'webp', 'gif', 'mp4', 'mov', 'avi', 'mkv']
)

col1, col2 = st.columns(2)
with col1:
    columns_count = st.selectbox('📊 가로 바둑판 배열 개수', [3, 4, 5], index=0)
with col2:
    excel_name = st.text_input('📝 저장될 엑셀 파일명', value='소재_자동_인덱싱_리포트')

def clean_and_sort_filename(filename):
    normalized_name = unicodedata.normalize('NFC', filename)
    numbers = re.findall(r'\d+', normalized_name)
    sort_key = int(numbers[0]) if numbers else 9999
    return sort_key, normalized_name

if st.button('🚀 엑셀 리포트 추출하기'):
    if not uploaded_files:
        st.error('⚠️ 업로드된 파일이 없습니다. 파일을 넣어주세요!')
    else:
        media_targets = []
        temp_dir = tempfile.TemporaryDirectory()
        
        for f in uploaded_files:
            _, safe_name = clean_and_sort_filename(f.name)
            t_path = os.path.join(temp_dir.name, safe_name)
            with open(t_path, "wb") as temp_f:
                temp_f.write(f.read())
            media_targets.append((safe_name, t_path))
            
        media_targets.sort(key=lambda x: clean_and_sort_filename(x[0])[0])
        
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('소재 인덱싱')
        
        worksheet.set_column('A:A', 5)
        worksheet.set_column('B:Z', 28)
        name_format = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter', 
            'text_wrap': True, 'font_size': 10, 'border': 1, 'bg_color': '#F4F6F4'
        })
        
        start_row, start_col = 1, 1
        VIDEO_EXTENSIONS = ('.mp4', '.mov', '.avi', '.mkv')
        
        for index, (filename, file_path) in enumerate(media_targets):
            status_text.text(f'🔄 깨짐 교정 및 인덱싱 중 ({index+1}/{len(media_targets)}): {filename}')
            progress_bar.progress((index + 1) / len(media_targets))
            
            upload_path = file_path
            if filename.lower().endswith(VIDEO_EXTENSIONS):
                cap = cv2.VideoCapture(file_path)
                success, frame = cap.read()
                if success:
                    preview_path = os.path.join(tempfile.gettempdir(), f"{os.path.splitext(filename)[0]}_preview.jpg")
                    cv2.imwrite(preview_path, frame)
                    cap.release()
                    upload_path = preview_path
                else:
                    cap.release()
                    continue
            
            grid_idx = index % columns_count
            grid_row_block = index // columns_count
            
            name_row = start_row + (grid_row_block * 3)
            img_row = name_row + 1
            current_col = start_col + grid_idx
            
            worksheet.set_row(name_row, 30)
            worksheet.set_row(img_row, 160)
            worksheet.write(name_row, current_col, filename, name_format)
            
            try:
                worksheet.insert_image(img_row, current_col, upload_path, {
                    'x_scale': 0.18, 'y_scale': 0.18, 'x_offset': 5, 'y_offset': 5, 'object_position': 1
                })
            except:
                worksheet.write(img_row, current_col, '(이미지 삽입 실패)')
                
        workbook.close()
        output.seek(0)
        
        status_text.success(f'🎉 교정 완료! 총 {len(media_targets)}개의 광고 소재 리포트가 완성되었습니다.')
        
        st.download_button(
            label='📥 완성된 한글 정형 엑셀 파일 PC로 저장하기', 
            data=output, 
            file_name=f'{excel_name}.xlsx', 
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )