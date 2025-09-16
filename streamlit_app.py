import streamlit as st
import pandas as pd
import os
from pathlib import Path
import json
from datetime import datetime

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
        
    try:
        with open(html_path, "r", encoding="utf-8") as f:
            html_content = f.read()
        
        # 브라우저 저장소 사용 부분 제거 (Streamlit 환경에서 제한적)
        html_content = html_content.replace("localStorage", "{}") 
        html_content = html_content.replace("sessionStorage", "{}")
        
        # 외부 CDN 링크가 HTTPS인지 확인
        html_content = html_content.replace("http://", "https://")
        
        return html_content
    except Exception as e:
        st.error(f"HTML 파일 로드 중 오류: {e}")
        return None

def inject_streamlit_communication():
    """HTML과 Streamlit 간 통신을 위한 JavaScript 주입"""
    return """
    <script>
    // Streamlit과 통신하기 위한 함수들
    function sendDataToStreamlit(data) {
        try {
            // PostMessage API 사용하여 부모 프레임(Streamlit)에 데이터 전송
            window.parent.postMessage({
                type: 'streamlit:setComponentValue',
                value: data
            }, '*');
        } catch (e) {
            console.log('데이터 전송 오류:', e);
        }
    }
    
    // 저장 버튼 클릭 시 데이터를 Streamlit으로 전송
    function interceptSaveButtons() {
        const saveButtons = document.querySelectorAll('[onclick*="saveData"], .btn-primary, .save-btn');
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
            if (element.id || element.name) {
                const key = element.id || element.name;
                data[key] = element.value || '';
            }
        });
        
        return data;
    }
    
    // 페이지 로드 완료 후 실행
    document.addEventListener('DOMContentLoaded', function() {
        interceptSaveButtons();
        
        // 주기적으로 데이터 수집 (자동 저장)
        setInterval(() => {
            const formData = collectFormData();
            if (Object.keys(formData).length > 0) {
                sendDataToStreamlit(formData);
            }
        }, 5000); // 5초마다
    });
    </script>
    """

def create_sample_form():
    """HTML 파일이 없을 때 사용할 샘플 폼"""
    return """
    <div style="padding: 20px; font-family: Arial, sans-serif;">
        <h2>위험성평가 입력 폼</h2>
        <form id="risk-form">
            <div style="margin-bottom: 15px;">
                <label>회사명:</label><br>
                <input type="text" id="company_name" style="width: 100%; padding: 8px;" placeholder="회사명을 입력하세요">
            </div>
            <div style="margin-bottom: 15px;">
                <label>작업내용:</label><br>
                <textarea id="work_content" rows="3" style="width: 100%; padding: 8px;" placeholder="작업내용을 입력하세요"></textarea>
            </div>
            <div style="margin-bottom: 15px;">
                <label>위험요인:</label><br>
                <input type="text" id="hazard_factor" style="width: 100%; padding: 8px;" placeholder="위험요인을 입력하세요">
            </div>
            <div style="margin-bottom: 15px;">
                <label>가능성 (1-5):</label><br>
                <select id="possibility" style="width: 100%; padding: 8px;">
                    <option value="">선택하세요</option>
                    <option value="1">1 - 매우 낮음</option>
                    <option value="2">2 - 낮음</option>
                    <option value="3">3 - 보통</option>
                    <option value="4">4 - 높음</option>
                    <option value="5">5 - 매우 높음</option>
                </select>
            </div>
            <div style="margin-bottom: 15px;">
                <label>중대성 (1-4):</label><br>
                <select id="severity" style="width: 100%; padding: 8px;">
                    <option value="">선택하세요</option>
                    <option value="1">1 - 경미</option>
                    <option value="2">2 - 보통</option>
                    <option value="3">3 - 심각</option>
                    <option value="4">4 - 치명적</option>
                </select>
            </div>
            <button type="button" class="save-btn" style="background: #0066cc; color: white; padding: 10px 20px; border: none; border-radius: 4px; cursor: pointer;">저장</button>
        </form>
    </div>
    """

def calculate_risk_score(possibility, severity):
    """위험성 점수 계산"""
    try:
        p = int(possibility) if possibility else 0
        s = int(severity) if severity else 0
        score = p * s
        
        if score >= 12:
            level = "높음"
            color = "red"
        elif score >= 6:
            level = "보통"
            color = "orange"
        else:
            level = "낮음"
            color = "green"
            
        return score, level, color
    except:
        return 0, "미평가", "gray"

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
        "위험성평가표": "RA(4)_riskAssessment.html",
        "샘플폼": "sample"  # 샘플 폼 추가
    }
    
    selected_page = st.sidebar.selectbox("페이지 선택", list(html_files.keys()))
    
    # 세션 상태 초기화
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {}
    if 'assessment_data' not in st.session_state:
        st.session_state.assessment_data = []
    
    # 선택된 페이지 표시
    if selected_page in html_files:
        if selected_page == "샘플폼":
            # 샘플 폼 사용
            html_content = create_sample_form()
            st.info("HTML 파일이 없어 샘플 폼을 표시합니다.")
        else:
            filename = html_files[selected_page]
            html_content = load_html_with_modifications(filename)
            
            if not html_content:
                st.warning(f"HTML 파일을 찾을 수 없습니다: {filename}")
                st.info("html_files 폴더에 해당 파일이 있는지 확인하거나 '샘플폼'을 선택해주세요.")
                html_content = create_sample_form()
        
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
                st.sidebar.success("데이터가 수신되었습니다!")
                
                # 위험성 점수 계산
                if 'possibility' in component_value and 'severity' in component_value:
                    score, level, color = calculate_risk_score(
                        component_value.get('possibility'), 
                        component_value.get('severity')
                    )
                    component_value['risk_score'] = score
                    component_value['risk_level'] = level
                    component_value['timestamp'] = datetime.now().isoformat()
                
                # 세션 상태에 저장
                st.session_state.form_data[selected_page] = component_value
                
                # 위험성평가 데이터인 경우 별도 저장
                if 'work_content' in component_value or 'hazard_factor' in component_value:
                    # 중복 체크
                    existing = next((item for item in st.session_state.assessment_data 
                                   if item.get('work_content') == component_value.get('work_content')), None)
                    
                    if not existing:
                        st.session_state.assessment_data.append(component_value.copy())
                    else:
                        # 기존 데이터 업데이트
                        for i, item in enumerate(st.session_state.assessment_data):
                            if item.get('work_content') == component_value.get('work_content'):
                                st.session_state.assessment_data[i] = component_value.copy()
                                break
                
                # 받은 데이터 표시
                with st.sidebar.expander("수신된 데이터"):
                    st.json(component_value)
    
    # 데이터 관리 섹션
    st.sidebar.markdown("---")
    st.sidebar.subheader("데이터 관리")
    
    # 저장된 평가 데이터 표시
    if st.session_state.assessment_data:
        st.sidebar.info(f"저장된 평가 항목: {len(st.session_state.assessment_data)}개")
        
        if st.sidebar.button("평가 데이터 보기"):
            st.subheader("저장된 위험성평가 데이터")
            df = pd.DataFrame(st.session_state.assessment_data)
            st.dataframe(df, use_container_width=True)
    
    # Excel 보고서 생성
    if st.sidebar.button("Excel 보고서 생성"):
        if st.session_state.assessment_data or st.session_state.form_data:
            generate_excel_report()
        else:
            st.sidebar.warning("저장된 데이터가 없습니다.")
    
    # 데이터 초기화
    if st.sidebar.button("데이터 초기화", type="secondary"):
        if st.sidebar.checkbox("정말로 초기화하시겠습니까?"):
            st.session_state.form_data = {}
            st.session_state.assessment_data = []
            st.sidebar.success("데이터가 초기화되었습니다.")
            st.rerun()

def generate_excel_report():
    """Excel 보고서 생성"""
    try:
        from io import BytesIO
        import xlsxwriter
        
        buffer = BytesIO()
        workbook = xlsxwriter.Workbook(buffer)
        
        # 헤더 포맷
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#D7E4BD',
            'border': 1
        })
        
        # 위험도별 색상 포맷
        high_risk_format = workbook.add_format({'bg_color': '#FFB3BA', 'border': 1})
        medium_risk_format = workbook.add_format({'bg_color': '#FFDFBA', 'border': 1})
        low_risk_format = workbook.add_format({'bg_color': '#BAFFC9', 'border': 1})
        
        # 위험성평가 데이터가 있으면 시트 생성
        if st.session_state.assessment_data:
            df_assessment = pd.DataFrame(st.session_state.assessment_data)
            worksheet = workbook.add_worksheet('위험성평가')
            
            # 헤더 작성
            for col, header in enumerate(df_assessment.columns):
                worksheet.write(0, col, header, header_format)
            
            # 데이터 작성 및 조건부 서식 적용
            for row, data in enumerate(df_assessment.to_dict('records'), 1):
                for col, (key, value) in enumerate(data.items()):
                    # 위험성 점수에 따른 색상 적용
                    if key == 'risk_score' and isinstance(value, (int, float)):
                        if value >= 12:
                            format_to_use = high_risk_format
                        elif value >= 6:
                            format_to_use = medium_risk_format
                        else:
                            format_to_use = low_risk_format
                        worksheet.write(row, col, value, format_to_use)
                    else:
                        worksheet.write(row, col, str(value) if value is not None else '')
        
        # 기타 폼 데이터 시트 생성
        for page_name, data in st.session_state.form_data.items():
            if data:  # 데이터가 있는 경우만
                worksheet = workbook.add_worksheet(page_name[:31])  # 시트명 길이 제한
                
                row = 0
                for key, value in data.items():
                    worksheet.write(row, 0, key, header_format)
                    worksheet.write(row, 1, str(value) if value is not None else '')
                    row += 1
        
        workbook.close()
        buffer.seek(0)
        
        filename = f"위험성평가_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        st.sidebar.download_button(
            label="📥 Excel 파일 다운로드",
            data=buffer.getvalue(),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.sidebar.success("Excel 보고서가 생성되었습니다!")
        
    except Exception as e:
        st.sidebar.error(f"Excel 생성 오류: {e}")

if __name__ == "__main__":
    main()
