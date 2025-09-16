import streamlit as st
import pandas as pd
import os
from pathlib import Path

# 페이지 설정
st.set_page_config(
    page_title="위험성평가 시스템",
    page_icon="⚠️",
    layout="wide",
    initial_sidebar_state="expanded"
)

def load_html_with_modifications(filename):
    """HTML 파일을 로드하고 Streamlit에 맞게 수정"""
    html_path = Path("html_files") / filename
    
    if not html_path.exists():
        return None
        
    with open(html_path, "r", encoding="utf-8") as f:
        html_content = f.read()
    
    # localStorage 사용 부분을 sessionStorage로 변경하거나 제거
    # (Streamlit 환경에서는 sessionStorage도 제한적)
    html_content = html_content.replace("localStorage", "sessionStorage")
    
    # 외부 CDN 링크가 HTTPS인지 확인
    html_content = html_content.replace("http://", "https://")
    
    return html_content

def inject_streamlit_communication():
    """HTML과 Streamlit 간 통신을 위한 JavaScript 주입"""
    return """
    <script>
    // Streamlit과 통신하기 위한 함수들
    function sendDataToStreamlit(data) {
        // PostMessage API 사용하여 부모 프레임(Streamlit)에 데이터 전송
        window.parent.postMessage({
            type: 'streamlit:setComponentValue',
            value: data
        }, '*');
    }
    
    // 저장 버튼 클릭 시 데이터를 Streamlit으로 전송
    function interceptSaveButtons() {
        const saveButtons = document.querySelectorAll('[onclick*="saveData"], .btn-primary');
        saveButtons.forEach(btn => {
            btn.addEventListener('click', function(e) {
                // 기존 동작 실행 후 데이터 수집
                setTimeout(() => {
                    const formData = collectFormData();
                    sendDataToStreamlit(formData);
                }, 100);
            });
        });
    }
    
    function collectFormData() {
        const data = {};
        
        // 모든 input, textarea, select 요소의 값 수집
        document.querySelectorAll('input, textarea, select').forEach(element => {
            if (element.id) {
                data[element.id] = element.value;
            }
        });
        
        return data;
    }
    
    // 페이지 로드 완료 후 실행
    document.addEventListener('DOMContentLoaded', interceptSaveButtons);
    </script>
    """

def main():
    st.title("위험성평가 시스템")
    
    # 사이드바 메뉴
    st.sidebar.title("메뉴")
    
    # HTML 파일 목록
    html_files = {
        "통합시스템": "integrated-risk-assessment.html",
        "표지": "RA(1)-title.html", 
        "작업공정": "RA(2)-WorkProcess.html",
        "위험정보": "RA(3)-RiskInfo.html",
        "위험성평가표": "RA(4)_riskAssessment.html"
    }
    
    selected_page = st.sidebar.selectbox("페이지 선택", list(html_files.keys()))
    
    # 선택된 HTML 파일 표시
    if selected_page in html_files:
        filename = html_files[selected_page]
        html_content = load_html_with_modifications(filename)
        
        if html_content:
            # JavaScript 통신 코드 주입
            enhanced_html = html_content + inject_streamlit_communication()
            
            # HTML 컴포넌트로 표시
            component_value = st.components.v1.html(
                enhanced_html, 
                height=800, 
                scrolling=True
            )
            
            # HTML에서 전송된 데이터 처리
            if component_value:
                st.sidebar.write("받은 데이터:")
                st.sidebar.json(component_value)
                
                # 세션 상태에 저장
                if 'form_data' not in st.session_state:
                    st.session_state.form_data = {}
                st.session_state.form_data[selected_page] = component_value
                
        else:
            st.error(f"HTML 파일을 찾을 수 없습니다: {filename}")
            st.info("html_files 폴더에 해당 파일이 있는지 확인해주세요.")
    
    # 데이터 내보내기 섹션
    if st.sidebar.button("Excel 보고서 생성"):
        generate_excel_report()
    
    # 저장된 데이터 표시
    if 'form_data' in st.session_state and st.session_state.form_data:
        with st.sidebar.expander("저장된 데이터 보기"):
            st.json(st.session_state.form_data)

def generate_excel_report():
    """Excel 보고서 생성"""
    try:
        from io import BytesIO
        import xlsxwriter
        
        buffer = BytesIO()
        workbook = xlsxwriter.Workbook(buffer)
        
        # 저장된 데이터가 있으면 시트로 생성
        if 'form_data' in st.session_state:
            for page_name, data in st.session_state.form_data.items():
                worksheet = workbook.add_worksheet(page_name)
                
                # 데이터를 행별로 작성
                row = 0
                for key, value in data.items():
                    worksheet.write(row, 0, key)
                    worksheet.write(row, 1, str(value))
                    row += 1
        
        workbook.close()
        buffer.seek(0)
        
        st.sidebar.download_button(
            label="Excel 파일 다운로드",
            data=buffer.getvalue(),
            file_name=f"위험성평가_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.sidebar.error(f"Excel 생성 오류: {e}")

if __name__ == "__main__":
    main()