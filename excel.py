import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="마이리얼트립 리포트 변환기", layout="centered")

# --- 세션 상태(Session State) 초기화 ---
# 업로더 위젯을 초기화하기 위한 키(key)
if 'uploader_key' not in st.session_state:
    st.session_state['uploader_key'] = str(0)
# 변환된 데이터를 임시 저장할 공간
if 'processed_data' not in st.session_state:
    st.session_state['processed_data'] = None

st.title("📊 마이리얼트립 광고 리포트 변환기")
st.write("다운로드한 엑셀 리포트를 업로드하세요. (변환이 완료되면 원본 파일은 자동으로 화면에서 지워집니다.)")

# 파일 업로드 (key 값을 세션 상태와 연결하여 자동 초기화가 가능하게 설정)
uploaded_files = st.file_uploader(
    "여기로 엑셀 파일을 드래그하거나 클릭하여 업로드하세요.", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True,
    key=st.session_state['uploader_key']
)

# 파일이 업로드된 상태일 때만 추출 버튼 표시
if uploaded_files:
    if st.button("데이터 추출 및 변환하기", type="primary"):
        all_data = []
        
        with st.spinner("데이터를 변환하는 중입니다... 잠시만 기다려주세요."):
            for file in uploaded_files:
                try:
                    # 1. 메타 정보 추출
                    df_meta = pd.read_excel(file, header=None, nrows=2)
                    a1_val = str(df_meta.iloc[0, 0])
                    a2_val = str(df_meta.iloc[1, 0])
                    
                    product_id = re.sub(r'[^0-9]', '', a1_val)
                    product_name = a2_val.replace('상품명:', '').strip()
                    
                    # 2. 메인 데이터 읽기
                    df_data = pd.read_excel(file, header=None, skiprows=8)
                    
                    # 3. 새 데이터프레임 구성
                    df_new = pd.DataFrame()
                    df_new['상품 ID'] = [product_id] * len(df_data)
                    df_new['MRT 상품명'] = [product_name] * len(df_data)
                    df_new['날짜'] = df_data.iloc[:, 0]
                    df_new['광고유형'] = df_data.iloc[:, 1]
                    df_new['키워드'] = df_data.iloc[:, 2]
                    df_new['노출수'] = df_data.iloc[:, 3]
                    df_new['클릭수'] = df_data.iloc[:, 4]
                    df_new['광고비'] = df_data.iloc[:, 6]
                    df_new['광고예약수'] = df_data.iloc[:, 7]
                    df_new['광고전환매출'] = df_data.iloc[:, 9]
                    
                    # 빈 날짜 제거
                    df_new = df_new.dropna(subset=['날짜'])
                    
                    all_data.append(df_new)
                    
                except Exception as e:
                    st.error(f"'{file.name}' 처리 중 오류가 발생했습니다: {e}")
            
            # 4. 변환 완료 시 데이터 저장 및 자동 새로고침(파일 삭제)
            if all_data:
                final_df = pd.concat(all_data, ignore_index=True)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='RawData')
                
                # 변환된 데이터를 메모리에 저장
                st.session_state['processed_data'] = output.getvalue()
                
                # uploader_key 값을 변경하여 업로더를 강제 초기화(빈 상태로 만듦)
                st.session_state['uploader_key'] = str(int(st.session_state['uploader_key']) + 1)
                
                # 즉시 화면을 새로고침 (이 코드가 실행되면 파일 목록이 자동으로 싹 사라집니다)
                st.rerun()

# 화면 새로고침 후, 변환된 데이터가 저장되어 있다면 다운로드 버튼을 표시
if st.session_state['processed_data'] is not None:
    st.success("✨ 변환 완료! 업로드된 원본 파일은 자동으로 삭제되었습니다.")
    
    st.download_button(
        label="📥 변환된 통합 엑셀 파일 다운로드",
        data=st.session_state['processed_data'],
        file_name="마이리얼트립_통합데이터.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # 다음 작업을 위해 모든 상태를 초기화하는 버튼
    if st.button("🔄 새로운 작업 시작하기"):
        st.session_state['processed_data'] = None
        st.rerun()