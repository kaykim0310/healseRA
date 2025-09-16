# -*- coding: utf-8 -*-
"""
위험성평가 데이터 처리 모듈
Risk Assessment Data Handler Module

안전보건컨설팅을 위한 위험성평가 데이터 관리 시스템
"""

import pandas as pd
import json
import streamlit as st
from typing import Dict, List, Any, Optional, Union
import openpyxl
from datetime import datetime
import logging
import sqlite3
from pathlib import Path

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class RiskAssessmentDataHandler:
    """위험성평가 데이터 핸들러 클래스"""
    
    def __init__(self, db_path: str = "risk_assessment.db"):
        """
        초기화
        Args:
            db_path: SQLite 데이터베이스 파일 경로
        """
        self.db_path = db_path
        self.init_database()
        
        # 위험성 등급 기준
        self.risk_levels = {
            "매우높음": {"min": 16, "max": 20, "color": "#ff4444", "action": "즉시작업중지"},
            "높음": {"min": 12, "max": 15, "color": "#ff8888", "action": "즉시개선"},
            "보통": {"min": 6, "max": 11, "color": "#ffaa44", "action": "계획적개선"},
            "낮음": {"min": 3, "max": 5, "color": "#44ff88", "action": "모니터링"},
            "매우낮음": {"min": 1, "max": 2, "color": "#88ff88", "action": "현상유지"}
        }
        
    def init_database(self):
        """데이터베이스 초기화"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 회사정보 테이블
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS company_info (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    company_name TEXT NOT NULL,
                    address TEXT,
                    tel TEXT,
                    fax TEXT,
                    ceo_name TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # 작업공정 테이블
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS work_processes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    company_id INTEGER,
                    process_name TEXT NOT NULL,
                    description TEXT,
                    equipment TEXT,  -- JSON 형태로 저장
                    chemicals TEXT,  -- JSON 형태로 저장
                    hazards TEXT,    -- JSON 형태로 저장
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (company_id) REFERENCES company_info (id)
                )
            ''')
            
            # 위험성평가 테이블
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS risk_assessments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    process_id INTEGER,
                    work_content TEXT,
                    hazard_classification TEXT,
                    hazard_cause TEXT,
                    hazard_factor TEXT,
                    legal_basis TEXT,
                    current_state TEXT,
                    possibility INTEGER,  -- 가능성 (1-5)
                    severity INTEGER,     -- 중대성 (1-4)
                    risk_score INTEGER,   -- 위험성 점수
                    risk_level TEXT,      -- 위험성 등급
                    reduction_measures TEXT,
                    improved_risk TEXT,
                    improvement_date DATE,
                    completion_date DATE,
                    manager_name TEXT,
                    notes TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (process_id) REFERENCES work_processes (id)
                )
            ''')
            
            # 법적근거 테이블
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS legal_basis (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    category TEXT,
                    subcategory TEXT,
                    legal_text TEXT,
                    regulation_number TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            conn.commit()
            conn.close()
            logger.info("데이터베이스 초기화 완료")
            
        except Exception as e:
            logger.error(f"데이터베이스 초기화 오류: {e}")
            raise
    
    def save_company_info(self, company_data: Dict[str, Any]) -> int:
        """
        회사 정보 저장
        Args:
            company_data: 회사 정보 딕셔너리
        Returns:
            company_id: 저장된 회사 ID
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO company_info (company_name, address, tel, fax, ceo_name)
                VALUES (?, ?, ?, ?, ?)
            ''', (
                company_data.get('name', ''),
                company_data.get('address', ''),
                company_data.get('tel', ''),
                company_data.get('fax', ''),
                company_data.get('ceo', '')
            ))
            
            company_id = cursor.lastrowid
            conn.commit()
            conn.close()
            
            logger.info(f"회사 정보 저장 완료 (ID: {company_id})")
            return company_id
            
        except Exception as e:
            logger.error(f"회사 정보 저장 오류: {e}")
            raise
    
    def save_work_process(self, company_id: int, process_data: Dict[str, Any]) -> int:
        """
        작업공정 저장
        Args:
            company_id: 회사 ID
            process_data: 공정 데이터
        Returns:
            process_id: 저장된 공정 ID
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO work_processes (company_id, process_name, description, equipment, chemicals, hazards)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (
                company_id,
                process_data.get('name', ''),
                process_data.get('description', ''),
                json.dumps(process_data.get('equipment', []), ensure_ascii=False),
                json.dumps(process_data.get('chemicals', []), ensure_ascii=False),
                json.dumps(process_data.get('hazards', []), ensure_ascii=False)
            ))
            
            process_id = cursor.lastrowid
            conn.commit()
            conn.close()
            
            logger.info(f"작업공정 저장 완료 (ID: {process_id})")
            return process_id
            
        except Exception as e:
            logger.error(f"작업공정 저장 오류: {e}")
            raise
    
    def calculate_risk_score(self, possibility: int, severity: int) -> Dict[str, Any]:
        """
        위험성 점수 및 등급 계산
        Args:
            possibility: 가능성 (1-5)
            severity: 중대성 (1-4)
        Returns:
            위험성 정보 딕셔너리
        """
        score = possibility * severity
        level = "낮음"
        
        for level_name, criteria in self.risk_levels.items():
            if criteria["min"] <= score <= criteria["max"]:
                level = level_name
                break
        
        level_info = self.risk_levels.get(level, self.risk_levels["낮음"])
        
        return {
            "score": score,
            "level": level,
            "color": level_info["color"],
            "action": level_info["action"],
            "possibility": possibility,
            "severity": severity
        }
    
    def save_risk_assessment(self, process_id: int, assessment_data: Dict[str, Any]) -> int:
        """
        위험성평가 저장
        Args:
            process_id: 공정 ID
            assessment_data: 평가 데이터
        Returns:
            assessment_id: 저장된 평가 ID
        """
        try:
            # 위험성 계산
            risk_info = self.calculate_risk_score(
                assessment_data.get('possibility', 1),
                assessment_data.get('severity', 1)
            )
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                INSERT INTO risk_assessments (
                    process_id, work_content, hazard_classification, hazard_cause,
                    hazard_factor, legal_basis, current_state, possibility, severity,
                    risk_score, risk_level, reduction_measures, improved_risk,
                    improvement_date, completion_date, manager_name, notes
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                process_id,
                assessment_data.get('work_content', ''),
                assessment_data.get('classification', ''),
                assessment_data.get('cause', ''),
                assessment_data.get('hazard', ''),
                assessment_data.get('legal_basis', ''),
                assessment_data.get('current_state', ''),
                assessment_data.get('possibility', 1),
                assessment_data.get('severity', 1),
                risk_info['score'],
                risk_info['level'],
                assessment_data.get('measures', ''),
                assessment_data.get('improved_risk', ''),
                assessment_data.get('improvement_date'),
                assessment_data.get('completion_date'),
                assessment_data.get('manager', ''),
                assessment_data.get('note', '')
            ))
            
            assessment_id = cursor.lastrowid
            conn.commit()
            conn.close()
            
            logger.info(f"위험성평가 저장 완료 (ID: {assessment_id})")
            return assessment_id
            
        except Exception as e:
            logger.error(f"위험성평가 저장 오류: {e}")
            raise
    
    def load_legal_basis_from_excel(self, file_path: str) -> bool:
        """
        엑셀 파일에서 법적근거 로드
        Args:
            file_path: 엑셀 파일 경로
        Returns:
            성공 여부
        """
        try:
            df = pd.read_excel(file_path)
            
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # 기존 법적근거 삭제
            cursor.execute('DELETE FROM legal_basis')
            
            # 새 데이터 삽입
            for _, row in df.iterrows():
                cursor.execute('''
                    INSERT INTO legal_basis (category, subcategory, legal_text, regulation_number)
                    VALUES (?, ?, ?, ?)
                ''', (
                    row.get('대분류', ''),
                    row.get('중분류', ''),
                    row.get('법적근거', ''),
                    row.get('조항', '')
                ))
            
            conn.commit()
            conn.close()
            
            logger.info("법적근거 데이터 로드 완료")
            return True
            
        except Exception as e:
            logger.error(f"법적근거 로드 오류: {e}")
            return False
    
    def get_legal_basis_by_hazard(self, hazard_type: str) -> List[str]:
        """
        위험요인에 따른 법적근거 조회
        Args:
            hazard_type: 위험요인 유형
        Returns:
            관련 법적근거 리스트
        """
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('''
                SELECT legal_text FROM legal_basis 
                WHERE subcategory LIKE ? OR category LIKE ?
            ''', (f'%{hazard_type}%', f'%{hazard_type}%'))
            
            results = [row[0] for row in cursor.fetchall()]
            conn.close()
            
            return results
            
        except Exception as e:
            logger.error(f"법적근거 조회 오류: {e}")
            return []
    
    def get_company_data(self, company_id: int) -> Dict[str, Any]:
        """회사 데이터 조회"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM company_info WHERE id = ?', (company_id,))
            row = cursor.fetchone()
            conn.close()
            
            if row:
                return {
                    'id': row[0],
                    'name': row[1],
                    'address': row[2],
                    'tel': row[3],
                    'fax': row[4],
                    'ceo': row[5],
                    'created_at': row[6]
                }
            return {}
            
        except Exception as e:
            logger.error(f"회사 데이터 조회 오류: {e}")
            return {}
    
    def get_process_data(self, company_id: int) -> List[Dict[str, Any]]:
        """공정 데이터 조회"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM work_processes WHERE company_id = ?', (company_id,))
            rows = cursor.fetchall()
            conn.close()
            
            processes = []
            for row in rows:
                processes.append({
                    'id': row[0],
                    'company_id': row[1],
                    'name': row[2],
                    'description': row[3],
                    'equipment': json.loads(row[4]) if row[4] else [],
                    'chemicals': json.loads(row[5]) if row[5] else [],
                    'hazards': json.loads(row[6]) if row[6] else [],
                    'created_at': row[7]
                })
            
            return processes
            
        except Exception as e:
            logger.error(f"공정 데이터 조회 오류: {e}")
            return []
    
    def get_assessment_data(self, process_id: int) -> List[Dict[str, Any]]:
        """위험성평가 데이터 조회"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('SELECT * FROM risk_assessments WHERE process_id = ?', (process_id,))
            rows = cursor.fetchall()
            conn.close()
            
            assessments = []
            for row in rows:
                assessments.append({
                    'id': row[0],
                    'process_id': row[1],
                    'work_content': row[2],
                    'classification': row[3],
                    'cause': row[4],
                    'hazard': row[5],
                    'legal_basis': row[6],
                    'current_state': row[7],
                    'possibility': row[8],
                    'severity': row[9],
                    'risk_score': row[10],
                    'risk_level': row[11],
                    'measures': row[12],
                    'improved_risk': row[13],
                    'improvement_date': row[14],
                    'completion_date': row[15],
                    'manager': row[16],
                    'note': row[17],
                    'created_at': row[18]
                })
            
            return assessments
            
        except Exception as e:
            logger.error(f"위험성평가 데이터 조회 오류: {e}")
            return []
    
    def delete_assessment(self, assessment_id: int) -> bool:
        """위험성평가 항목 삭제"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            cursor.execute('DELETE FROM risk_assessments WHERE id = ?', (assessment_id,))
            conn.commit()
            conn.close()
            
            logger.info(f"위험성평가 항목 삭제 완료 (ID: {assessment_id})")
            return True
            
        except Exception as e:
            logger.error(f"위험성평가 삭제 오류: {e}")
            return False
    
    def backup_database(self, backup_path: str) -> bool:
        """데이터베이스 백업"""
        try:
            import shutil
            shutil.copy2(self.db_path, backup_path)
            logger.info(f"데이터베이스 백업 완료: {backup_path}")
            return True
        except Exception as e:
            logger.error(f"데이터베이스 백업 오류: {e}")
            return False
    
    def restore_database(self, backup_path: str) -> bool:
        """데이터베이스 복원"""
        try:
            import shutil
            shutil.copy2(backup_path, self.db_path)
            logger.info(f"데이터베이스 복원 완료: {backup_path}")
            return True
        except Exception as e:
            logger.error(f"데이터베이스 복원 오류: {e}")
            return False

# Streamlit 전용 세션 상태 관리
class StreamlitDataManager:
    """Streamlit 세션 상태 관리 클래스"""
    
    @staticmethod
    def save_to_session(key: str, data: Any):
        """세션에 데이터 저장"""
        if 'risk_assessment_data' not in st.session_state:
            st.session_state.risk_assessment_data = {}
        st.session_state.risk_assessment_data[key] = data
    
    @staticmethod
    def load_from_session(key: str, default=None):
        """세션에서 데이터 로드"""
        if 'risk_assessment_data' in st.session_state:
            return st.session_state.risk_assessment_data.get(key, default)
        return default
    
    @staticmethod
    def clear_session():
        """세션 데이터 초기화"""
        if 'risk_assessment_data' in st.session_state:
            del st.session_state.risk_assessment_data

# 유틸리티 함수들
def validate_risk_data(data: Dict[str, Any]) -> List[str]:
    """위험성평가 데이터 유효성 검사"""
    errors = []
    
    required_fields = ['work_content', 'hazard', 'possibility', 'severity']
    for field in required_fields:
        if not data.get(field):
            errors.append(f"{field} 필드가 필요합니다.")
    
    possibility = data.get('possibility')
    if possibility and (not isinstance(possibility, int) or possibility < 1 or possibility > 5):
        errors.append("가능성은 1-5 사이의 정수여야 합니다.")
    
    severity = data.get('severity')
    if severity and (not isinstance(severity, int) or severity < 1 or severity > 4):
        errors.append("중대성은 1-4 사이의 정수여야 합니다.")
    
    return errors

def create_sample_data() -> Dict[str, Any]:
    """샘플 데이터 생성"""
    return {
        'company': {
            'name': '(주)안전기업',
            'address': '서울시 강남구 테헤란로 123',
            'tel': '02-1234-5678',
            'fax': '02-1234-5679',
            'ceo': '김안전'
        },
        'processes': [
            {
                'name': '원료투입',
                'description': '원재료를 투입하는 공정',
                'equipment': ['컨베이어', '호퍼', '계량기'],
                'chemicals': ['유기용제', '접착제'],
                'hazards': ['끼임', '화재', '흡입']
            },
            {
                'name': '가공작업',
                'description': '제품을 가공하는 공정',
                'equipment': ['프레스', '절단기', '연삭기'],
                'chemicals': ['냉각유', '윤활유'],
                'hazards': ['절단', '화상', '소음']
            }
        ]
    }
