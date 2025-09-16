# -*- coding: utf-8 -*-
"""
위험성평가 데이터 처리 모듈 (개선된 버전)
Risk Assessment Data Handler Module - Improved

주요 개선사항:
1. 예외 처리 강화
2. 타입 힌트 개선
3. 메모리 효율성 향상
4. SQLite 연결 관리 개선
"""

import pandas as pd
import json
import streamlit as st
from typing import Dict, List, Any, Optional, Union
import sqlite3
from datetime import datetime, date
import logging
from pathlib import Path
from contextlib import contextmanager

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class RiskAssessmentDataHandler:
    """위험성평가 데이터 핸들러 클래스 - 개선된 버전"""
    
    def __init__(self, db_path: str = "risk_assessment.db"):
        """
        초기화
        Args:
            db_path: SQLite 데이터베이스 파일 경로
        """
        self.db_path = Path(db_path)
        self.init_database()
        
        # 위험성 등급 기준
        self.risk_levels = {
            "매우높음": {"min": 16, "max": 20, "color": "#ff4444", "action": "즉시작업중지"},
            "높음": {"min": 12, "max": 15, "color": "#ff8888", "action": "즉시개선"},
            "보통": {"min": 6, "max": 11, "color": "#ffaa44", "action": "계획적개선"},
            "낮음": {"min": 3, "max": 5, "color": "#44ff88", "action": "모니터링"},
            "매우낮음": {"min": 1, "max": 2, "color": "#88ff88", "action": "현상유지"}
        }

    @contextmanager
    def get_connection(self):
        """데이터베이스 연결 컨텍스트 매니저"""
        conn = None
        try:
            conn = sqlite3.connect(str(self.db_path))
            conn.execute("PRAGMA foreign_keys = ON")  # 외래키 제약 조건 활성화
            yield conn
        except Exception as e:
            if conn:
                conn.rollback()
            logger.error(f"데이터베이스 오류: {e}")
            raise
        finally:
            if conn:
                conn.close()
    
    def init_database(self):
        """데이터베이스 초기화 - 개선된 버전"""
        try:
            with self.get_connection() as conn:
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
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                ''')
                
                # 나머지 테이블들... (원본과 동일하되 업데이트 트리거 추가)
                cursor.execute('''
                    CREATE TRIGGER IF NOT EXISTS update_company_info_timestamp 
                    AFTER UPDATE ON company_info
                    BEGIN
                        UPDATE company_info SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
                    END;
                ''')
                
                conn.commit()
                logger.info("데이터베이스 초기화 완료")
                
        except Exception as e:
            logger.error(f"데이터베이스 초기화 오류: {e}")
            raise

    def calculate_risk_score(self, possibility: int, severity: int) -> Dict[str, Any]:
        """
        위험성 점수 및 등급 계산 - 입력값 검증 강화
        """
        try:
            # 입력값 검증
            if not isinstance(possibility, int) or not isinstance(severity, int):
                raise ValueError("가능성과 중대성은 정수여야 합니다.")
            
            if not (1 <= possibility <= 5):
                raise ValueError("가능성은 1-5 사이의 값이어야 합니다.")
                
            if not (1 <= severity <= 4):
                raise ValueError("중대성은 1-4 사이의 값이어야 합니다.")
            
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
            
        except Exception as e:
            logger.error(f"위험성 계산 오류: {e}")
            return {
                "score": 0,
                "level": "오류",
                "color": "#cccccc",
                "action": "재계산필요",
                "possibility": possibility,
                "severity": severity,
                "error": str(e)
            }

    def save_company_info(self, company_data: Dict[str, Any]) -> Optional[int]:
        """
        회사 정보 저장 - 개선된 오류 처리
        """
        try:
            with self.get_connection() as conn:
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
                
                logger.info(f"회사 정보 저장 완료 (ID: {company_id})")
                return company_id
                
        except Exception as e:
            logger.error(f"회사 정보 저장 오류: {e}")
            return None

    def get_statistics(self) -> Dict[str, Any]:
        """전체 통계 정보 조회"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                # 기본 통계
                cursor.execute('SELECT COUNT(*) FROM company_info')
                company_count = cursor.fetchone()[0]
                
                cursor.execute('SELECT COUNT(*) FROM work_processes')
                process_count = cursor.fetchone()[0]
                
                cursor.execute('SELECT COUNT(*) FROM risk_assessments')
                assessment_count = cursor.fetchone()[0]
                
                # 위험도별 통계
                cursor.execute('''
                    SELECT risk_level, COUNT(*) 
                    FROM risk_assessments 
                    GROUP BY risk_level
                ''')
                risk_stats = dict(cursor.fetchall())
                
                return {
                    'companies': company_count,
                    'processes': process_count,
                    'assessments': assessment_count,
                    'risk_distribution': risk_stats,
                    'generated_at': datetime.now().isoformat()
                }
                
        except Exception as e:
            logger.error(f"통계 조회 오류: {e}")
            return {}

# Streamlit 전용 세션 상태 관리 - 개선된 버전
class StreamlitDataManager:
    """Streamlit 세션 상태 관리 클래스 - 개선된 버전"""
    
    @staticmethod
    def save_to_session(key: str, data: Any) -> bool:
        """세션에 데이터 저장 - 성공 여부 반환"""
        try:
            if 'risk_assessment_data' not in st.session_state:
                st.session_state.risk_assessment_data = {}
            st.session_state.risk_assessment_data[key] = data
            return True
        except Exception as e:
            logger.error(f"세션 저장 오류: {e}")
            return False
    
    @staticmethod
    def load_from_session(key: str, default=None) -> Any:
        """세션에서 데이터 로드 - 개선된 오류 처리"""
        try:
            if 'risk_assessment_data' in st.session_state:
                return st.session_state.risk_assessment_data.get(key, default)
            return default
        except Exception as e:
            logger.error(f"세션 로드 오류: {e}")
            return default
    
    @staticmethod
    def get_session_size() -> Dict[str, Any]:
        """세션 데이터 크기 정보"""
        try:
            if 'risk_assessment_data' not in st.session_state:
                return {'items': 0, 'size_estimate': 0}
            
            data = st.session_state.risk_assessment_data
            items = len(data)
            size_estimate = len(str(data))  # 대략적인 크기
            
            return {
                'items': items,
                'size_estimate': size_estimate,
                'keys': list(data.keys())
            }
        except Exception as e:
            logger.error(f"세션 크기 계산 오류: {e}")
            return {'items': 0, 'size_estimate': 0, 'error': str(e)}

# 유틸리티 함수들 - 개선된 버전
def validate_risk_data(data: Dict[str, Any]) -> List[str]:
    """위험성평가 데이터 유효성 검사 - 강화된 버전"""
    errors = []
    
    # 필수 필드 검사
    required_fields = {
        'work_content': '작업내용',
        'hazard_factor': '위험요인',
        'possibility': '가능성',
        'severity': '중대성'
    }
    
    for field, korean_name in required_fields.items():
        value = data.get(field)
        if not value or (isinstance(value, str) and not value.strip()):
            errors.append(f"{korean_name} 필드가 필요합니다.")
    
    # 숫자 필드 검사
    possibility = data.get('possibility')
    if possibility is not None:
        try:
            p = int(possibility)
            if p < 1 or p > 5:
                errors.append("가능성은 1-5 사이의 정수여야 합니다.")
        except (ValueError, TypeError):
            errors.append("가능성은 숫자여야 합니다.")
    
    severity = data.get('severity')
    if severity is not None:
        try:
            s = int(severity)
            if s < 1 or s > 4:
                errors.append("중대성은 1-4 사이의 정수여야 합니다.")
        except (ValueError, TypeError):
            errors.append("중대성은 숫자여야 합니다.")
    
    return errors

def create_test_data() -> Dict[str, Any]:
    """테스트용 샘플 데이터 생성"""
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
            }
        ],
        'assessments': [
            {
                'work_content': '원료 투입 작업',
                'hazard_factor': '컨베이어 벨트 끼임',
                'possibility': 3,
                'severity': 3,
                'legal_basis': '산업안전보건법 제23조',
                'measures': '안전커버 설치, 비상정지장치 설치'
            }
        ]
    }
