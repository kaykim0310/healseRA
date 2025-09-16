import pandas as pd
import json
import streamlit as st
from typing import Dict, List, Any
import openpyxl
from datetime import datetime

class RiskAssessmentData:
    """위험성평가 데이터 관리 클래스"""
    
    def __init__(self):
        self.title_data = {}
        self.process_data = {}
        self.risk_info_data = {}
        self.assessment_data = []
        
    def save_title_data(self, data: Dict[str, Any]):
        """표지 데이터 저장"""
        self.title_data = data
        self._save_to_session('title_data', data)
        
    def save_process_data(self, data: Dict[str, Any]):
        """작업공정 데이터 저장"""
        self.process_data = data
        self._save_to_session('process_data', data)
        
    def save_risk_info_data(self, data: Dict[str, Any]):
        """위험정보 데이터 저장"""
        self.risk_info_data = data
        self._save_to_session('risk_info_data', data)
        
    def add_assessment_item(self, item: Dict[str, Any]):
        """위험성평가 항목 추가"""
        self.assessment_data.append(item)
        self._save_to_session('assessment_data', self.assessment_data)
        
    def calculate_risk_score(self, possibility: int, severity: int) -> Dict[str, Any]:
        """위험성 점수 계산"""
        score = possibility * severity
        
        if score >= 12:
            level = "높음"
            action = "즉시 개선/작업중지"
            color = "red"
        elif score >= 6:
            level = "보통" 
            action = "개선"
            color = "orange"
        else:
            level = "낮음"
            action = "모니터링"
            color = "green"
            
        return {
            "score": score,
            "level": level,
            "action": action,
            "color": color
        }
        
    def load_legal_basis_excel(self, file_path: str) -> pd.DataFrame:
        """법적근거 엑셀 파일 로드"""
        try:
            df = pd.read_excel(file_path)
            return df
        except Exception as e:
            st.error(f"엑셀 파일 로드 중 오류: {e}")
            return pd.DataFrame()
            
    def get_legal_basis_for_hazard(self, hazard_type: str, df: pd.DataFrame) -> List[str]:
        """위험요인에 대한 법적근거 찾기"""
        if df.empty:
            return []
            
        # 위험요인 유형에 따른 법적근거 필터링
        filtered = df[df['분류'].str.contains(hazard_type, na=False)]
        return filtered['법적근거'].dropna().tolist()
        
    def export_to_excel(self) -> bytes:
        """데이터를 엑셀로 내보내기"""
        from io import BytesIO
        
        buffer = BytesIO()
        
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # 표지 데이터
            if self.title_data:
                title_df = pd.DataFrame([self.title_data])
                title_df.to_excel(writer, sheet_name='표지', index=False)
                
            # 작업공정 데이터
            if self.process_data:
                process_df = pd.DataFrame(self.process_data.get('processes', []))
                process_df.to_excel(writer, sheet_name='작업공정', index=False)
                
            # 위험성평가 데이터
            if self.assessment_data:
                assessment_df = pd.DataFrame(self.assessment_data)
                assessment_df.to_excel(writer, sheet_name='위험성평가', index=False)
                
        buffer.seek(0)
        return buffer.getvalue()
        
    def _save_to_session(self, key: str, data: Any):
        """세션 상태에 데이터 저장"""
        if 'risk_assessment_data' not in st.session_state:
            st.session_state.risk_assessment_data = {}
        st.session_state.risk_assessment_data[key] = data
        
    def _load_from_session(self, key: str, default=None):
        """세션 상태에서 데이터 로드"""
        if 'risk_assessment_data' in st.session_state:
            return st.session_state.risk_assessment_data.get(key, default)
        return default
        
    def load_all_data(self):
        """모든 저장된 데이터 로드"""
        self.title_data = self._load_from_session('title_data', {})
        self.process_data = self._load_from_session('process_data', {})
        self.risk_info_data = self._load_from_session('risk_info_data', {})
        self.assessment_data = self._load_from_session('assessment_data', [])

def create_sample_data():
    """샘플 데이터 생성"""
    return {
        'title_data': {
            'company_name': '(주)안전기업',
            'company_address': '서울시 강남구 테헤란로 123',
            'company_tel': '02-1234-5678',
            'ceo_name': '김안전'
        },
        'process_data': {
            'processes': [
                {
                    'name': '원료투입',
                    'description': '원재료를 투입하는 공정',
                    'equipment': ['컨베이어', '호퍼'],
                    'chemicals': ['유기용제']
                }
            ]
        }
    }