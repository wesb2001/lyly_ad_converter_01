from flask import Flask, request, send_file, render_template, after_this_request
import os
from werkzeug.utils import secure_filename
import pandas as pd
from datetime import datetime
import glob

app = Flask(__name__)

# 업로드된 파일을 저장할 디렉토리
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_next_version(base_filename):
    # 오늘 생성된 같은 날짜의 파일들 검색
    pattern = f"{base_filename}_v*.xlsx"
    existing_files = glob.glob(os.path.join(app.config['CONVERTED_FOLDER'], pattern))
    
    if not existing_files:
        return "v01"
    
    # 기존 파일들의 버전 번호 추출
    versions = []
    for file in existing_files:
        try:
            version = file.split("_")[-1].replace(".xlsx", "")  # v01, v02 등 추출
            version_num = int(version.replace("v", ""))
            versions.append(version_num)
        except:
            continue
    
    if not versions:
        return "v01"
    
    # 최대 버전 번호에 1을 더해서 반환
    next_version = max(versions) + 1
    return f"v{next_version:02d}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return '파일이 없습니다.'
    file = request.files['file']
    if file.filename == '':
        return '선택된 파일이 없습니다.'
    
    if file and allowed_file(file.filename):
        # 업로드된 파일 저장
        filename = secure_filename(file.filename)
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(input_path)
        
        try:
            # 데이터프레임 생성하여 날짜 정보 추출
            df = pd.read_excel(input_path)
            start_date = pd.to_datetime(df['보고 시작'].iloc[0]).strftime('%y%m%d')
            end_date = pd.to_datetime(df['보고 종료'].iloc[0]).strftime('%y%m%d')
            
            # 기본 파일명 생성 (버전 제외)
            base_filename = f"LYLYL_{start_date}_{end_date}"
            
            # 다음 버전 번호 가져오기
            version = get_next_version(base_filename)
            
            # 최종 출력 파일명 생성
            output_filename = f"{base_filename}_{version}.xlsx"
            output_path = os.path.join(app.config['CONVERTED_FOLDER'], output_filename)
            
            # 파일 변환
            from auto_convert_excel import convert_excel_file
            convert_excel_file(input_path, output_path)
            
            # 파일 다운로드 후 임시 파일 삭제
            @after_this_request
            def remove_files(response):
                try:
                    os.remove(input_path)
                    os.remove(output_path)
                except:
                    pass
                return response
            
            return send_file(output_path, as_attachment=True, download_name=output_filename)
            
        except Exception as e:
            return f'파일 처리 중 오류가 발생했습니다: {str(e)}'
    
    return '허용되지 않는 파일 형식입니다.'

if __name__ == '__main__':
    # 업로드 및 변환 폴더가 없으면 생성
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['CONVERTED_FOLDER'], exist_ok=True)
    app.run(debug=True)