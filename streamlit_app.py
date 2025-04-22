import streamlit as st
import pandas as pd
import os
from auto_convert_excel import convert_excel_file
from datetime import datetime
import tempfile

st.set_page_config(
    page_title="LYLYL 광고 데이터 변환기",
    page_icon="📊",
    layout="centered"
)

st.title("LYLYL 광고 데이터 변환기")

st.markdown("""
### 사용 방법
1. Meta 광고 데이터 Excel 파일(.xlsx 또는 .xls)을 선택하세요.
2. '변환하기' 버튼을 클릭하면 자동으로 데이터가 변환됩니다.
3. 변환된 파일은 자동으로 다운로드됩니다.
4. 파일명은 'LYLYL_시작일_종료일_버전.xlsx' 형식으로 저장됩니다.
""")

uploaded_file = st.file_uploader("Excel 파일을 선택하세요", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # 임시 파일로 저장
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_input:
            tmp_input.write(uploaded_file.getvalue())
            input_path = tmp_input.name

        # 임시 출력 파일 경로 생성
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_output:
            output_path = tmp_output.name

        # 파일 변환
        convert_excel_file(input_path, output_path)

        # 변환된 파일 다운로드 버튼 생성
        with open(output_path, 'rb') as f:
            df = pd.read_excel(input_path)
            start_date = pd.to_datetime(df['보고 시작'].iloc[0]).strftime('%y%m%d')
            end_date = pd.to_datetime(df['보고 종료'].iloc[0]).strftime('%y%m%d')
            output_filename = f"LYLYL_{start_date}_{end_date}.xlsx"
            
            st.download_button(
                label="변환된 파일 다운로드",
                data=f,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        # 임시 파일 삭제
        os.unlink(input_path)
        os.unlink(output_path)

        st.success("✅ 변환이 완료되었습니다!")
        
    except Exception as e:
        st.error(f"❌ 오류가 발생했습니다: {str(e)}")
        st.error("파일 형식을 확인해주세요.") 