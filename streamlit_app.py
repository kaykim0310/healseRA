import streamlit as st
import pandas as pd
import os
from pathlib import Path
import json
from datetime import datetime, date
from io import BytesIO
import xlsxwriter

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
    """HTML과 Streamlit 간 통신을 위한 JavaScript 주입 - 오류 수정 버전"""
    return """
    <script>
    // Streamlit과 통신하기 위한 함수들
    function sendDataToStreamlit(data) {
        try {
            // 데이터 검증 및 정리
            const cleanData = {};
            
            for (const [key, value] of Object.entries(data)) {
                // 값이 존재하고 유효한 경우만 포함
                if (value !== null && value !== undefined && value !== '') {
                    cleanData[key] = String(value).trim();
                }
            }
            
            // 빈 객체인 경우 전송하지 않음
            if (Object.keys(cleanData).length === 0) {
                console.log('전송할 데이터가 없습니다.');
                return;
            }
            
            console.log('전송할 데이터:', cleanData);
            
            // PostMessage API 사용하여 부모 프레임(Streamlit)에 데이터 전송
            window.parent.postMessage({
                type: 'streamlit:setComponentValue',
                value: cleanData
            }, '*');
            
        } catch (e) {
            console.error('데이터 전송 오류:', e);
        }
    }
    
    // 저장 버튼 클릭 시 데이터를 Streamlit으로 전송
    function interceptSaveButtons() {
        const saveButtons = document.querySelectorAll('[onclick*="saveData"], .btn-primary, .save-btn, button[type="submit"], input[type="submit"]');
        console.log('발견된 저장 버튼 수:', saveButtons.length);
        
        saveButtons.forEach((btn, index) => {
            btn.addEventListener('click', function(e) {
                console.log('저장 버튼 클릭됨:', index);
                // 기존 동작 실행 후 데이터 수집
                setTimeout(() => {
                    const formData = collectFormData();
                    sendDataToStreamlit(formData);
                }, 200);
            });
        });
    }
    
    function collectFormData() {
        const data = {};
        
        try {
            // 모든 input, textarea, select 요소의 값 수집
            document.querySelectorAll('input, textarea, select').forEach(element => {
                const key = element.id || element.name || element.getAttribute('data-name');
                const value = element.value;
                
                if (key && value !== null && value !== undefined) {
                    // 특수 문자 제거 및 정리
                    const cleanKey = key.replace(/[^a-zA-Z0-9_가-힣]/g, '_');
                    const cleanValue = String(value).trim();
                    
                    if (cleanValue !== '') {
                        data[cleanKey] = cleanValue;
                    }
                }
            });
            
            // 현재 시간 추가
            data['timestamp'] = new Date().toISOString();
            data['page_type'] = 'form_data';
            
            console.log('수집된 데이터:', data);
            return data;
            
        } catch (e) {
            console.error('데이터 수집 오류:', e);
            return {};
        }
    }
    
    // 자동 데이터 수집 (변경 감지)
    function setupAutoCollection() {
        const inputs = document.querySelectorAll('input, textarea, select');
        
        inputs.forEach(input => {
            input.addEventListener('change', function() {
                setTimeout(() => {
                    const formData = collectFormData();
                    if (Object.keys(formData).length > 2) { // timestamp, page_type 외에 데이터가 있는 경우
                        sendDataToStreamlit(formData);
                    }
                }, 300);
            });
        });
    }
    
    // 페이지 로드 완료 후 실행
    document.addEventListener('DOMContentLoaded', function() {
        console.log('페이지 로드 완료');
        
        // 버튼 이벤트 설정
        interceptSaveButtons();
        
        // 자동 수집 설정
        setupAutoCollection();
        
        // 초기 데이터 전송 (페이지 정보)
        setTimeout(() => {
            const initialData = {
                'page_loaded': 'true',
                'timestamp': new Date().toISOString(),
                'page_type': 'page_load'
            };
            sendDataToStreamlit(initialData);
        }, 1000);
    });
    
    // 에러 처리
    window.addEventListener('error', function(e) {
        console.error('JavaScript 오류:', e.error);
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

def show_native_form(selected_page):
    """Streamlit 네이티브 위젯을 사용한 폼"""
    
    if selected_page in ["표지", "통합시스템"]:
        st.subheader("1. 사업장 개요")
        
        with st.form(f"form_{selected_page}"):
            col1, col2 = st.columns(2)
            
            with col1:
                company_name = st.text_input("사업장명", value="", help="회사명을 입력하세요")
                main_products = st.text_area("주요생산품", value="", help="주요 생산품을 입력하세요")
                evaluation_date = st.date_input("평가일자", value=date.today())
                
            with col2:
                business_type = st.selectbox("업종", 
                    ["제조업", "건설업", "서비스업", "운송업", "기타"])
                employee_count = st.number_input("근로자 수", min_value=1, value=10)
                ceo_name = st.text_input("대표자명", value="")
                
            address = st.text_area("사업장 주소")
            telephone = st.text_input("전화번호")
            
            submitted = st.form_submit_button("저장")
            
            if submitted:
                form_data = {
                    'company_name': company_name,
                    'main_products': main_products,
                    'evaluation_date': str(evaluation_date),
                    'business_type': business_type,
                    'employee_count': employee_count,
                    'ceo_name': ceo_name,
                    'address': address,
                    'telephone': telephone,
                    'timestamp': datetime.now().isoformat()
                }
                
                st.session_state.form_data[selected_page] = form_data
                st.success("데이터가 저장되었습니다!")
                return form_data
                
    elif selected_page == "작업공정":
        st.subheader("2. 작업공정 정보")
        
        with st.form("process_form"):
            process_name = st.text_input("공정명")
            process_desc = st.text_area("공정 설명")
            
            col1, col2 = st.columns(2)
            with col1:
                equipment = st.text_area("사용 장비 (한 줄에 하나씩)", help="예:\n프레스\n컨베이어\n호퍼")
            with col2:
                chemicals = st.text_area("화학물질 (한 줄에 하나씩)", help="예:\n유기용제\n접착제\n윤활유")
                
            hazards = st.text_area("위험요인 (한 줄에 하나씩)", help="예:\n끼임\n절단\n화상")
            
            submitted = st.form_submit_button("공정 저장")
            
            if submitted and process_name:
                process_data = {
                    'name': process_name,
                    'description': process_desc,
                    'equipment': [item.strip() for item in equipment.split('\n') if item.strip()],
                    'chemicals': [item.strip() for item in chemicals.split('\n') if item.strip()],
                    'hazards': [item.strip() for item in hazards.split('\n') if item.strip()],
                    'timestamp': datetime.now().isoformat()
                }
                
                if 'processes' not in st.session_state:
                    st.session_state.processes = []
                st.session_state.processes.append(process_data)
                
                st.session_state.form_data[selected_page] = {'processes': st.session_state.processes}
                st.success("공정 정보가 저장되었습니다!")
                return process_data
                
    elif selected_page in ["위험정보", "위험성평가표"]:
        st.subheader("위험성평가 입력")
        
        with st.form("risk_assessment_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                work_content = st.text_area("작업내용", help="수행하는 작업을 상세히 기술")
                hazard_factor = st.text_input("위험요인", help="예: 끼임, 절단, 화상 등")
                hazard_cause = st.text_area("위험요인 발생원인")
                
            with col2:
                possibility = st.selectbox("가능성 (1-5)", 
                    options=[1, 2, 3, 4, 5],
                    format_func=lambda x: f"{x} - {'매우낮음,낮음,보통,높음,매우높음'.split(',')[x-1]}")
                
                severity = st.selectbox("중대성 (1-4)", 
                    options=[1, 2, 3, 4],
                    format_func=lambda x: f"{x} - {'경미,보통,심각,치명적'.split(',')[x-1]}")
                
                legal_basis = st.text_input("법적근거", help="관련 법령 조항")
            
            current_measures = st.text_area("현재 안전조치")
            improvement_measures = st.text_area("개선대책")
            manager_name = st.text_input("담당자")
            
            submitted = st.form_submit_button("위험성평가 저장")
            
            if submitted and work_content and hazard_factor:
                # 위험성 점수 계산
                risk_score = possibility * severity
                if risk_score >= 12:
                    risk_level = "높음"
                    risk_color = "red"
                    action = "즉시 개선"
                elif risk_score >= 6:
                    risk_level = "보통" 
                    risk_color = "orange"
                    action = "계획적 개선"
                else:
                    risk_level = "낮음"
                    risk_color = "green"
                    action = "모니터링"
                
                assessment_item = {
                    'work_content': work_content,
                    'hazard_factor': hazard_factor,
                    'hazard_cause': hazard_cause,
                    'possibility': possibility,
                    'severity': severity,
                    'risk_score': risk_score,
                    'risk_level': risk_level,
                    'risk_color': risk_color,
                    'action': action,
                    'legal_basis': legal_basis,
                    'current_measures': current_measures,
                    'improvement_measures': improvement_measures,
                    'manager_name': manager_name,
                    'timestamp': datetime.now().isoformat()
                }
                
                # 평가 데이터에 추가
                if 'assessment_data' not in st.session_state:
                    st.session_state.assessment_data = []
                st.session_state.assessment_data.append(assessment_item)
                
                st.success(f"위험성평가가 저장되었습니다! (위험도: {risk_level}, 점수: {risk_score})")
                
                # 색상에 따라 알림 표시
                if risk_level == "높음":
                    st.error("⚠️ 높은 위험도입니다. 즉시 개선이 필요합니다!")
                elif risk_level == "보통":
                    st.warning("⚡ 보통 위험도입니다. 계획적 개선이 필요합니다.")
                else:
                    st.info("✅ 낮은 위험도입니다. 지속적인 모니터링을 실시하세요.")
                
                return assessment_item
                
    return None

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

def process_form_data(raw_data):
    """폼 데이터 처리 함수"""
    processed = raw_data.copy()
    
    try:
        # 위험성 점수 계산
        if 'possibility' in processed and 'severity' in processed:
            score, level, color = calculate_risk_score(
                processed.get('possibility'), 
                processed.get('severity')
            )
            processed['risk_score'] = score
            processed['risk_level'] = level
            processed['risk_color'] = color
        
        # 타임스탬프 추가
        if 'timestamp' not in processed:
            processed['timestamp'] = datetime.now().isoformat()
            
        # 데이터 정리
        for key, value in list(processed.items()):
            if isinstance(value, str):
                processed[key] = value.strip()
                
    except Exception as e:
        st.sidebar.error(f"폼 데이터 처리 오류: {e}")
        
    return processed

def generate_excel_report():
    """Excel 보고서 생성"""
    try:
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
        
        # 1. 표지 정보
        if st.session_state.form_data:
            title_sheet = workbook.add_worksheet('사업장정보')
            row = 0
            for page_name, data in st.session_state.form_data.items():
                title_sheet.write(row, 0, f"=== {page_name} ===", header_format)
                row += 1
                if isinstance(data, dict):
                    for key, value in data.items():
                        title_sheet.write(row, 0, key)
                        title_sheet.write(row, 1, str(value))
                        row += 1
                row += 1
        
        # 2. 위험성평가 데이터가 있으면 시트 생성
        if hasattr(st.session_state, 'assessment_data') and st.session_state.assessment_data:
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
                        
            # 열 너비 자동 조정
            for col in range(len(df_assessment.columns)):
                worksheet.set_column(col, col, 15)
        
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

def main():
    st.title("위험성평가 시스템")
    
    # 사이드바 메뉴
    st.sidebar.title("📋 메뉴")
    
    # HTML 파일 목록
    html_files = {
        "통합시스템": "integrated-risk-assessment.html",
        "표지": "RA(1)-title.html", 
        "작업공정": "RA(2)-WorkProcess.html",
        "위험정보": "RA(3)-RiskInfo.html",
        "위험성평가표": "RA(4)_riskAssessment.html"
    }
    
    selected_page = st.sidebar.selectbox("페이지 선택", list(html_files.keys()))
    
    # 세션 상태 초기화
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {}
    if 'assessment_data' not in st.session_state:
        st.session_state.assessment_data = []
    if 'processes' not in st.session_state:
        st.session_state.processes = []
    
    # 폼 타입 선택
    form_type = st.sidebar.radio(
        "폼 형식 선택",
        ["네이티브 폼 (권장)", "HTML 폼 (실험적)"]
    )
    
    # 선택된 페이지 표시
    if form_type == "네이티브 폼 (권장)":
        # 네이티브 Streamlit 폼 사용
        result = show_native_form(selected_page)
        
    else:
        # HTML 컴포넌트 사용 (기존 방식)
        filename = html_files.get(selected_page, "")
        html_content = load_html_with_modifications(filename)
        
        if not html_content:
            st.warning(f"HTML 파일을 찾을 수 없습니다: {filename}")
            st.info("네이티브 폼을 사용하는 것을 권장합니다.")
            html_content = create_sample_form()
        
        if html_content:
            # JavaScript 통신 코드 주입
            enhanced_html = html_content + inject_streamlit_communication()
            
            # 디버깅 모드
            debug_mode = st.sidebar.checkbox("디버깅 모드")
            
            # HTML 컴포넌트로 표시
            component_value = st.components.v1.html(
                enhanced_html, 
                height=800, 
                scrolling=True
            )
            
            # 디버깅: 받은 데이터의 원시 형태 표시
            if debug_mode:
                st.sidebar.write("받은 데이터 타입:", type(component_value))
                st.sidebar.write("받은 데이터 내용:")
                st.sidebar.code(str(component_value))
                
                if component_value:
                    st.sidebar.write("JSON 직렬화 테스트:")
                    try:
                        json_str = json.dumps(component_value, ensure_ascii=False, indent=2)
                        st.sidebar.success("JSON 직렬화 성공!")
                        if st.sidebar.checkbox("JSON 내용 보기"):
                            st.sidebar.code(json_str)
                    except Exception as e:
                        st.sidebar.error(f"JSON 직렬화 실패: {e}")
            
            # HTML에서 전송된 데이터 처리
            if component_value:
                try:
                    # 데이터 타입 검증
                    if isinstance(component_value, dict):
                        # 빈 딕셔너리나 무의미한 데이터 필터링
                        meaningful_data = {k: v for k, v in component_value.items() 
                                         if v is not None and str(v).strip() != ''}
                        
                        if meaningful_data:
                            st.sidebar.success(f"유효한 데이터 수신: {len(meaningful_data)}개 필드")
                            
                            # 페이지 로드 정보는 별도 처리
                            if component_value.get('page_type') == 'page_load':
                                st.sidebar.info("페이지가 로드되었습니다")
                            else:
                                # 실제 폼 데이터 처리
                                processed_data = process_form_data(component_value)
                                st.session_state.form_data[selected_page] = processed_data
                                
                                # 위험성평가 데이터인 경우 별도 저장
                                if 'work_content' in processed_data and 'hazard_factor' in processed_data:
                                    st.session_state.assessment_data.append(processed_data.copy())
                                
                                # 받은 데이터 표시
                                with st.sidebar.expander("수신된 데이터", expanded=False):
                                    st.json(meaningful_data)
                        else:
                            st.sidebar.info("빈 데이터 수신됨")
                    else:
                        st.sidebar.warning(f"예상하지 못한 데이터 타입: {type(component_value)}")
                        st.sidebar.write(f"데이터: {component_value}")
                        
                except Exception as e:
                    st.sidebar.error(f"데이터 처리 오류: {e}")
                    import traceback
                    if debug_mode:
                        st.sidebar.code(traceback.format_exc())
    
    # 사이드바 - 데이터 관리 섹션
    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 데이터 관리")
    
    # 저장된 데이터 현황
    total_pages = len([k for k, v in st.session_state.form_data.items() if v])
    total_assessments = len(st.session_state.assessment_data) if hasattr(st.session_state, 'assessment_data') else 0
    total_processes = len(st.session_state.processes) if hasattr(st.session_state, 'processes') else 0
    
    st.sidebar.metric("저장된 페이지", total_pages)
    st.sidebar.metric("위험성평가 항목", total_assessments)
    st.sidebar.metric("작업공정", total_processes)
    
    # 저장된 평가 데이터 표시
    if total_assessments > 0:
        if st.sidebar.button("위험성평가 데이터 보기"):
            st.subheader("📋 저장된 위험성평가 데이터")
            df = pd.DataFrame(st.session_state.assessment_data)
            
            # 위험도별 색상 표시
            def highlight_risk(row):
                if row['risk_level'] == '높음':
                    return ['background-color: #ffcccc'] * len(row)
                elif row['risk_level'] == '보통':
                    return ['background-color: #fff3cd'] * len(row)
                elif row['risk_level'] == '낮음':
                    return ['background-color: #d4edda'] * len(row)
                else:
                    return [''] * len(row)
            
            st.dataframe(
                df.style.apply(highlight_risk, axis=1),
                use_container_width=True
            )
            
            # 통계 정보
            st.subheader("📈 위험성 평가 통계")
            col1, col2, col3 = st.columns(3)
            
            high_count = len([item for item in st.session_state.assessment_data if item['risk_level'] == '높음'])
            medium_count = len([item for item in st.session_state.assessment_data if item['risk_level'] == '보통'])
            low_count = len([item for item in st.session_state.assessment_data if item['risk_level'] == '낮음'])
            
            col1.metric("높음 위험", high_count, delta=None)
            col2.metric("보통 위험", medium_count, delta=None)
            col3.metric("낮음 위험", low_count, delta=None)
    
    # Excel 보고서 생성
    if st.sidebar.button("📊 Excel 보고서 생성"):
        if st.session_state.assessment_data or st.session_state.form_data:
            generate_excel_report()
        else:
            st.sidebar.warning("저장된 데이터가 없습니다.")
    
    # 데이터 초기화
    if st.sidebar.button("🗑️ 데이터 초기화", type="secondary"):
        if st.sidebar.checkbox("⚠️ 정말로 모든 데이터를 초기화하시겠습니까?"):
            st.session_state.form_data = {}
            st.session_state.assessment_data = []
            st.session_state.processes = []
            st.sidebar.success("데이터가 초기화되었습니다.")
            st.rerun()
    
    # 도움말
    with st.sidebar.expander("📖 사용법"):
        st.write("""
        **1단계**: 페이지 선택
        - 표지: 사업장 기본정보 입력
        - 작업공정: 공정별 상세정보 입력
        - 위험정보/위험성평가표: 위험성평가 수행
        
        **2단계**: 데이터 입력
        - 네이티브 폼 사용 권장
        - 모든 필드 입력 후 저장 버튼 클릭
        
        **3단계**: 보고서 생성
        - Excel 보고서 생성 버튼 클릭
        - 다운로드 링크를 통해 파일 저장
        """)

if __name__ == "__main__":
    main()
