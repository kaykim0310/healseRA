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
    """위험성평가 데이터 핸들러 클래스"""
    
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
        """데이터베이스 초기화"""
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
                        business_type TEXT,
                        employee_count INTEGER,
                        main_products TEXT,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
                        risk_color TEXT,      -- 위험성 색상
                        action_required TEXT, -- 필요조치
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
                
                # 업데이트 트리거 생성
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
        위험성 점수 및 등급 계산
        """
        try:
            # 입력값 검증
            if not isinstance(possibility, int) or not isinstance(severity, int):
                possibility = int(possibility) if str(possibility).isdigit() else 1
                severity = int(severity) if str(severity).isdigit() else 1
            
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
        회사 정보 저장
        """
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute('''
                    INSERT INTO company_info (
                        company_name, address, tel, fax, ceo_name, 
                        business_type, employee_count, main_products
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    company_data.get('company_name', ''),
                    company_data.get('address', ''),
                    company_data.get('telephone', company_data.get('tel', '')),
                    company_data.get('fax', ''),
                    company_data.get('ceo_name', ''),
                    company_data.get('business_type', ''),
                    company_data.get('employee_count', 0),
                    company_data.get('main_products', '')
                ))
                
                company_id = cursor.lastrowid
                conn.commit()
                
                logger.info(f"회사 정보 저장 완료 (ID: {company_id})")
                return company_id
                
        except Exception as e:
            logger.error(f"회사 정보 저장 오류: {e}")
            return None

    def save_work_process(self, company_id: int, process_data: Dict[str, Any]) -> Optional[int]:
        """
        작업공정 저장
        """
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute('''
                    INSERT INTO work_processes (
                        company_id, process_name, description, equipment, chemicals, hazards
                    ) VALUES (?, ?, ?, ?, ?, ?)
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
                
                logger.info(f"작업공정 저장 완료 (ID: {process_id})")
                return process_id
                
        except Exception as e:
            logger.error(f"작업공정 저장 오류: {e}")
            return None

    def save_risk_assessment(self, process_id: int, assessment_data: Dict[str, Any]) -> Optional[int]:
        """
        위험성평가 저장
        """
        try:
            # 위험성 계산
            risk_info = self.calculate_risk_score(
                assessment_data.get('possibility', 1),
                assessment_data.get('severity', 1)
            )
            
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute('''
                    INSERT INTO risk_assessments (
                        process_id, work_content, hazard_classification, hazard_cause,
                        hazard_factor, legal_basis, current_state, possibility, severity,
                        risk_score, risk_level, risk_color, action_required, 
                        reduction_measures, improved_risk, improvement_date, 
                        completion_date, manager_name, notes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    process_id,
                    assessment_data.get('work_content', ''),
                    assessment_data.get('hazard_classification', ''),
                    assessment_data.get('hazard_cause', ''),
                    assessment_data.get('hazard_factor', ''),
                    assessment_data.get('legal_basis', ''),
                    assessment_data.get('current_measures', ''),
                    assessment_data.get('possibility', 1),
                    assessment_data.get('severity', 1),
                    risk_info['score'],
                    risk_info['level'],
                    risk_info['color'],
                    risk_info['action'],
                    assessment_data.get('improvement_measures', ''),
                    assessment_data.get('improved_risk', ''),
                    assessment_data.get('improvement_date'),
                    assessment_data.get('completion_date'),
                    assessment_data.get('manager_name', ''),
                    assessment_data.get('notes', '')
                ))
                
                assessment_id = cursor.lastrowid
                conn.commit()
                
                logger.info(f"위험성평가 저장 완료 (ID: {assessment_id})")
                return assessment_id
                
        except Exception as e:
            logger.error(f"위험성평가 저장 오류: {e}")
            return None

    def load_legal_basis_from_excel(self, file_path: str) -> bool:
        """
        엑셀 파일에서 법적근거 로드
        """
        try:
            df = pd.read_excel(file_path)
            
            with self.get_connection() as conn:
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
                logger.info("법적근거 데이터 로드 완료")
                return True
                
        except Exception as e:
            logger.error(f"법적근거 로드 오류: {e}")
            return False

    def get_legal_basis_by_hazard(self, hazard_type: str) -> List[str]:
        """
        위험요인에 따른 법적근거 조회
        """
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute('''
                    SELECT legal_text FROM legal_basis 
                    WHERE subcategory LIKE ? OR category LIKE ?
                ''', (f'%{hazard_type}%', f'%{hazard_type}%'))
                
                results = [row[0] for row in cursor.fetchall()]
                return results
                
        except Exception as e:
            logger.error(f"법적근거 조회 오류: {e}")
            return []

    def get_company_data(self, company_id: int) -> Dict[str, Any]:
        """회사 데이터 조회"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute('SELECT * FROM company_info WHERE id = ?', (company_id,))
                row = cursor.fetchone()
                
                if row:
                    columns = [description[0] for description in cursor.description]
                    return dict(zip(columns, row))
                return {}
                
        except Exception as e:
            logger.error(f"회사 데이터 조회 오류: {e}")
            return {}

    def get_process_data(self, company_id: int) -> List[Dict[str, Any]]:
        """공정 데이터 조회"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                cursor.execute('SELECT * FROM work_processes WHERE company_id = ?', (company_id,))
                rows = cursor.fetchall()
                columns = [description[0] for description in cursor.description]
                
                processes = []
                for row in rows:
                    process_dict = dict(zip(columns, row))
                    # JSON 필드 파싱
                    for field in ['equipment', 'chemicals', 'hazards']:
                        if process_dict.get(field):
                            try:
                                process_dict[field] = json.loads(process_dict[field])
                            except:
                                process_dict[field] = []
                    processes.append(process_dict)
                
                return processes
                
        except Exception as e:
            logger.error(f"공정 데이터 조회 오류: {e}")
            return []

    def get_assessment_data(self, process_id: Optional[int] = None) -> List[Dict[str, Any]]:
        """위험성평가 데이터 조회"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                
                if process_id:
                    cursor.execute('SELECT * FROM risk_assessments WHERE process_id = ?', (process_id,))
                else:
                    cursor.execute('SELECT * FROM risk_assessments')
                    
                rows = cursor.fetchall()
                columns = [description[0] for description in cursor.description]
                
                assessments = []
                for row in rows:
                    assessments.append(dict(zip(columns, row)))
                
                return assessments
                
        except Exception as e:
            logger.error(f"위험성평가 데이터 조회 오류: {e}")
            return []

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

    def delete_assessment(self, assessment_id: int) -> bool:
        """위험성평가 항목 삭제"""
        try:
            with self.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM risk_assessments WHERE id = ?', (assessment_id,))
                conn.commit()
                logger.info(f"위험성평가 항목 삭제 완료 (ID: {assessment_id})")
                return True
                
        except Exception as e:
            logger.error(f"위험성평가 삭제 오류: {e}")
            return False

    def backup_database(self, backup_path: str) -> bool:
        """데이터베이스 백업"""
        try:
            import shutil
            shutil.copy2(str(self.db_path), backup_path)
            logger.info(f"데이터베이스 백업 완료: {backup_path}")
            return True
        except Exception as e:
            logger.error(f"데이터베이스 백업 오류: {e}")
            return False

    def restore_database(self, backup_path: str) -> bool:
        """데이터베이스 복원"""
        try:
            import shutil
            shutil.copy2(backup_path, str(self.db_path))
            logger.info(f"데이터베이스 복원 완료: {backup_path}")
            return True
        except Exception as e:
            logger.error(f"데이터베이스 복원 오류: {e}")
            return False

# Streamlit 전용 세션 상태 관리
class StreamlitDataManager:
    """Streamlit 세션 상태 관리 클래스"""
    
    @staticmethod
    def save_to_session(key: str, data: Any) -> bool:
        """세션에 데이터 저장"""
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
        """세션에서 데이터 로드"""
        try:
            if 'risk_assessment_data' in st.session_state:
                return st.session_state.risk_assessment_data.get(key, default)
            return default
        except Exception as e:
            logger.error(f"세션 로드 오류: {e}")
            return default
    
    @staticmethod
    def clear_session():
        """세션 데이터 초기화"""
        if 'risk_assessment_data' in st.session_state:
            del st.session_state.risk_assessment_data

    @staticmethod
    def get_session_size() -> Dict[str, Any]:
        """세션 데이터 크기 정보"""
        try:
            if 'risk_assessment_data' not in st.session_state:
                return {'items': 0, 'size_estimate': 0}
            
            data = st.session_state.risk_assessment_data
            items = len(data)
            size_estimate = len(str(data))
            
            return {
                'items': items,
                'size_estimate': size_estimate,
                'keys': list(data.keys())
            }
        except Exception as e:
            logger.error(f"세션 크기 계산 오류: {e}")
            return {'items': 0, 'size_estimate': 0, 'error': str(e)}

# 유틸리티 함수들
def validate_risk_data(data: Dict[str, Any]) -> List[str]:
    """위험성평가 데이터 유효성 검사"""
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

def create_sample_data() -> Dict[str, Any]:
    """샘플 데이터 생성"""
    return {
        'company': {
            'company_name': '(주)안전기업',
            'address': '서울시 강남구 테헤란로 123',
            'telephone': '02-1234-5678',
            'fax': '02-1234-5679',
            'ceo_name': '김안전',
            'business_type': '제조업',
            'employee_count': 50,
            'main_products': '자동차부품'
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
        ],
        'assessments': [
            {
                'work_content': '원료 투입 작업',
                'hazard_factor': '컨베이어 벨트 끼임',
                'hazard_cause': '안전커버 미설치',
                'possibility': 3,
                'severity': 3,
                'legal_basis': '산업안전보건법 제23조',
                'current_measures': '작업자 교육',
                'improvement_measures': '안전커버 설치, 비상정지장치 설치',
                'manager_name': '김안전'
            }
        ]
    }

def export_to_excel(data_handler: RiskAssessmentDataHandler, output_path: str) -> bool:
    """데이터를 엑셀로 내보내기"""
    try:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # 회사 정보 시트
            companies = []
            # 실제 구현에서는 데이터베이스에서 조회
            
            # 공정 정보 시트
            processes = []
            # 실제 구현에서는 데이터베이스에서 조회
            
            # 위험성평가 시트
            assessments = data_handler.get_assessment_data()
            if assessments:
                df_assessments = pd.DataFrame(assessments)
                df_assessments.to_excel(writer, sheet_name='위험성평가', index=False)
        
        logger.info(f"엑셀 파일 생성 완료: {output_path}")
        return True
        
    except Exception as e:
        logger.error(f"엑셀 내보내기 오류: {e}")
        return False
