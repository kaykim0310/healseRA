# -*- coding: utf-8 -*-
"""
위험성평가 보고서 생성 모듈
Risk Assessment Report Generator Module

안전보건컨설팅을 위한 위험성평가 보고서 생성 시스템
"""

import pandas as pd
import numpy as np
from io import BytesIO
import streamlit as st
from datetime import datetime, date
import json
from typing import Dict, List, Any, Optional
import base64
import logging

# Excel 관련
import xlsxwriter
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils.dataframe import dataframe_to_rows
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# PDF 관련
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Word 문서 관련
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# 차트 관련
try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class RiskAssessmentReportGenerator:
    """위험성평가 보고서 생성기"""
    
    def __init__(self, data_handler=None):
        self.data_handler = data_handler
        self.report_date = datetime.now()
        
        # 위험도별 색상 설정
        self.risk_colors = {
            "매우높음": "#ff4444",
            "높음": "#ff8888",
            "보통": "#ffaa44",
            "낮음": "#44ff88",
            "매우낮음": "#88ff88"
        }
        
    def generate_excel_report(self, data: Dict[str, Any]) -> bytes:
        """엑셀 보고서 생성"""
        try:
            buffer = BytesIO()
            
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # 스타일 정의
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'center',
                    'fg_color': '#D7E4BD',
                    'border': 1
                })
                
                cell_format = workbook.add_format({
                    'text_wrap': True,
                    'valign': 'top',
                    'border': 1
                })
                
                # 1. 표지 시트
                self._create_cover_sheet(writer, data, header_format, cell_format)
                
                # 2. 작업공정 시트
                self._create_process_sheet(writer, data, header_format, cell_format)
                
                # 3. 위험성평가 시트
                self._create_assessment_sheet(writer, data, header_format, cell_format)
                
                # 4. 통계 및 차트 시트
                self._create_statistics_sheet(writer, data, header_format, cell_format)
                
            buffer.seek(0)
            return buffer.getvalue()
            
        except Exception as e:
            logger.error(f"엑셀 보고서 생성 오류: {e}")
            raise
    
    def _create_cover_sheet(self, writer, data, header_format, cell_format):
        """표지 시트 생성"""
        try:
            worksheet = writer.book.add_worksheet('표지')
            
            # 제목
            title_format = writer.book.add_format({
                'bold': True,
                'font_size': 18,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            worksheet.merge_range('A1:D3', f'{self.report_date.year}년도 위험성평가 결과서', title_format)
            
            # 기본 정보
            title_data = data.get('title_data', data.get('표지', {}))
            info_data = [
                ['회사명', title_data.get('company_name', '')],
                ['주소', title_data.get('address', title_data.get('company_address', ''))],
                ['전화번호', title_data.get('telephone', title_data.get('company_tel', ''))],
                ['대표자', title_data.get('ceo_name', '')],
                ['업종', title_data.get('business_type', '')],
                ['근로자 수', str(title_data.get('employee_count', ''))],
                ['작성일', self.report_date.strftime('%Y-%m-%d')]
            ]
            
            row = 5
            for label, value in info_data:
                worksheet.write(row, 0, label, header_format)
                worksheet.write(row, 1, value, cell_format)
                row += 1
                
            # 열 너비 조정
            worksheet.set_column('A:A', 15)
            worksheet.set_column('B:D', 25)
            
        except Exception as e:
            logger.error(f"표지 시트 생성 오류: {e}")
    
    def _create_process_sheet(self, writer, data, header_format, cell_format):
        """작업공정 시트 생성"""
        try:
            # 여러 소스에서 공정 데이터 찾기
            processes = []
            
            # 세션 상태에서 공정 데이터 찾기
            if hasattr(st.session_state, 'processes') and st.session_state.processes:
                processes = st.session_state.processes
            elif 'process_data' in data and data['process_data'].get('processes'):
                processes = data['process_data']['processes']
            elif '작업공정' in data and isinstance(data['작업공정'], list):
                processes = data['작업공정']
            
            if not processes:
                return
                
            # 프로세스 데이터를 DataFrame으로 변환
            process_list = []
            for i, process in enumerate(processes):
                process_row = {
                    '번호': i + 1,
                    '공정명': process.get('name', ''),
                    '설명': process.get('description', ''),
                    '장비': ', '.join(process.get('equipment', [])) if isinstance(process.get('equipment'), list) else str(process.get('equipment', '')),
                    '화학물질': ', '.join(process.get('chemicals', [])) if isinstance(process.get('chemicals'), list) else str(process.get('chemicals', '')),
                    '위험요인': ', '.join(process.get('hazards', [])) if isinstance(process.get('hazards'), list) else str(process.get('hazards', ''))
                }
                process_list.append(process_row)
            
            if process_list:
                df_processes = pd.DataFrame(process_list)
                df_processes.to_excel(writer, sheet_name='작업공정', index=False, startrow=1)
                
                worksheet = writer.sheets['작업공정']
                
                # 헤더 스타일 적용
                for col_num, value in enumerate(df_processes.columns.values):
                    worksheet.write(1, col_num, value, header_format)
                    
                # 열 너비 조정
                worksheet.set_column('A:A', 8)   # 번호
                worksheet.set_column('B:B', 20)  # 공정명
                worksheet.set_column('C:C', 30)  # 설명
                worksheet.set_column('D:F', 25)  # 장비, 화학물질, 위험요인
                
        except Exception as e:
            logger.error(f"작업공정 시트 생성 오류: {e}")
    
    def _create_assessment_sheet(self, writer, data, header_format, cell_format):
        """위험성평가 시트 생성"""
        try:
            # 여러 소스에서 평가 데이터 찾기
            assessment_data = []
            
            if hasattr(st.session_state, 'assessment_data') and st.session_state.assessment_data:
                assessment_data = st.session_state.assessment_data
            elif 'assessment_data' in data:
                assessment_data = data['assessment_data']
            elif 'risk_assessments' in data:
                assessment_data = data['risk_assessments']
            
            if not assessment_data:
                return
                
            df_assessment = pd.DataFrame(assessment_data)
            
            # 컬럼 순서 정리
            desired_columns = [
                'work_content', 'hazard_factor', 'hazard_cause', 'possibility', 
                'severity', 'risk_score', 'risk_level', 'legal_basis',
                'current_measures', 'improvement_measures', 'manager_name'
            ]
            
            # 존재하는 컬럼만 선택
            available_columns = [col for col in desired_columns if col in df_assessment.columns]
            if available_columns:
                df_assessment = df_assessment[available_columns]
            
            # 컬럼명 한글화
            column_mapping = {
                'work_content': '작업내용',
                'hazard_factor': '위험요인',
                'hazard_cause': '위험요인 원인',
                'possibility': '가능성',
                'severity': '중대성',
                'risk_score': '위험성점수',
                'risk_level': '위험성등급',
                'legal_basis': '법적근거',
                'current_measures': '현재 안전조치',
                'improvement_measures': '개선대책',
                'manager_name': '담당자'
            }
            
            df_assessment = df_assessment.rename(columns=column_mapping)
            df_assessment.to_excel(writer, sheet_name='위험성평가표', index=False, startrow=1)
            
            worksheet = writer.sheets['위험성평가표']
            
            # 위험성 레벨에 따른 색상 적용
            high_risk_format = writer.book.add_format({
                'bg_color': '#FFB3BA',
                'border': 1
            })
            medium_risk_format = writer.book.add_format({
                'bg_color': '#FFDFBA',
                'border': 1
            })
            low_risk_format = writer.book.add_format({
                'bg_color': '#BAFFC9',
                'border': 1
            })
            
            # 헤더 스타일 적용
            for col_num, value in enumerate(df_assessment.columns.values):
                worksheet.write(1, col_num, value, header_format)
            
            # 위험성 점수에 따른 조건부 서식
            for row_num, row_data in enumerate(assessment_data, start=2):
                risk_score = row_data.get('risk_score', 0)
                risk_level = row_data.get('risk_level', '')
                
                # 위험성 점수와 등급 셀에 색상 적용
                if isinstance(risk_score, (int, float)) or str(risk_score).isdigit():
                    score_val = int(risk_score) if risk_score else 0
                    if score_val >= 12 or risk_level == '높음':
                        format_to_use = high_risk_format
                    elif score_val >= 6 or risk_level == '보통':
                        format_to_use = medium_risk_format
                    else:
                        format_to_use = low_risk_format
                        
                    # 위험성점수 열 찾기
                    score_cols = [i for i, col in enumerate(df_assessment.columns) if '위험성점수' in str(col)]
                    level_cols = [i for i, col in enumerate(df_assessment.columns) if '위험성등급' in str(col)]
                    
                    for col in score_cols + level_cols:
                        if col < len(df_assessment.columns):
                            cell_value = list(row_data.values())[col] if col < len(row_data) else ''
                            worksheet.write(row_num, col, cell_value, format_to_use)
            
            # 열 너비 자동 조정
            worksheet.set_column('A:A', 25)  # 작업내용
            worksheet.set_column('B:B', 20)  # 위험요인
            worksheet.set_column('C:C', 20)  # 위험요인 원인
            worksheet.set_column('D:F', 12)  # 가능성, 중대성, 점수
            worksheet.set_column('G:G', 12)  # 등급
            worksheet.set_column('H:H', 25)  # 법적근거
            worksheet.set_column('I:K', 30)  # 조치사항들
                        
        except Exception as e:
            logger.error(f"위험성평가 시트 생성 오류: {e}")
    
    def _create_statistics_sheet(self, writer, data, header_format, cell_format):
        """통계 시트 생성"""
        try:
            # 평가 데이터 가져오기
            assessment_data = []
            if hasattr(st.session_state, 'assessment_data') and st.session_state.assessment_data:
                assessment_data = st.session_state.assessment_data
            elif 'assessment_data' in data:
                assessment_data = data['assessment_data']
            
            if not assessment_data:
                return
                
            worksheet = writer.book.add_worksheet('통계')
            
            # 위험성 통계 계산
            risk_stats = self._calculate_risk_statistics(assessment_data)
            
            # 통계 테이블 작성
            worksheet.write(0, 0, '위험성 평가 통계', header_format)
            
            stats_data = [
                ['구분', '수량', '비율'],
                ['전체 항목', risk_stats['total'], '100%'],
                ['높음 (12-20)', risk_stats['high'], f"{risk_stats['high_percent']:.1f}%"],
                ['보통 (6-11)', risk_stats['medium'], f"{risk_stats['medium_percent']:.1f}%"],
                ['낮음 (1-5)', risk_stats['low'], f"{risk_stats['low_percent']:.1f}%"]
            ]
            
            for row_num, row_data in enumerate(stats_data):
                for col_num, cell_data in enumerate(row_data):
                    format_to_use = header_format if row_num == 0 else cell_format
                    worksheet.write(row_num + 2, col_num, cell_data, format_to_use)
                    
            # 열 너비 조정
            worksheet.set_column('A:C', 15)
            
            # 상세 분석 추가
            worksheet.write(8, 0, '위험요인별 분석', header_format)
            
            # 위험요인별 통계
            hazard_stats = {}
            for item in assessment_data:
                hazard = item.get('hazard_factor', '기타')
                level = item.get('risk_level', '미정')
                
                if hazard not in hazard_stats:
                    hazard_stats[hazard] = {'높음': 0, '보통': 0, '낮음': 0, '총계': 0}
                
                hazard_stats[hazard][level] = hazard_stats[hazard].get(level, 0) + 1
                hazard_stats[hazard]['총계'] += 1
            
            # 위험요인별 테이블 작성
            hazard_headers = ['위험요인', '높음', '보통', '낮음', '총계']
            for col_num, header in enumerate(hazard_headers):
                worksheet.write(10, col_num, header, header_format)
            
            row_num = 11
            for hazard, stats in hazard_stats.items():
                worksheet.write(row_num, 0, hazard, cell_format)
                worksheet.write(row_num, 1, stats.get('높음', 0), cell_format)
                worksheet.write(row_num, 2, stats.get('보통', 0), cell_format)
                worksheet.write(row_num, 3, stats.get('낮음', 0), cell_format)
                worksheet.write(row_num, 4, stats['총계'], cell_format)
                row_num += 1
                    
        except Exception as e:
            logger.error(f"통계 시트 생성 오류: {e}")
    
    def _calculate_risk_statistics(self, assessment_data: List[Dict]) -> Dict[str, Any]:
        """위험성 통계 계산"""
        try:
            total = len(assessment_data)
            if total == 0:
                return {
                    'total': 0, 'high': 0, 'medium': 0, 'low': 0,
                    'high_percent': 0, 'medium_percent': 0, 'low_percent': 0
                }
            
            high = 0
            medium = 0
            low = 0
            
            for item in assessment_data:
                risk_score = item.get('risk_score', 0)
                risk_level = item.get('risk_level', '')
                
                # 점수 또는 레벨로 판단
                if isinstance(risk_score, (int, float)) and risk_score >= 12:
                    high += 1
                elif isinstance(risk_score, (int, float)) and risk_score >= 6:
                    medium += 1
                elif isinstance(risk_score, (int, float)) and risk_score > 0:
                    low += 1
                elif risk_level == '높음':
                    high += 1
                elif risk_level == '보통':
                    medium += 1
                elif risk_level == '낮음':
                    low += 1
                else:
                    low += 1  # 기본값
            
            return {
                'total': total,
                'high': high,
                'medium': medium,
                'low': low,
                'high_percent': (high / total * 100) if total > 0 else 0,
                'medium_percent': (medium / total * 100) if total > 0 else 0,
                'low_percent': (low / total * 100) if total > 0 else 0
            }
        except Exception as e:
            logger.error(f"위험성 통계 계산 오류: {e}")
            return {
                'total': 0, 'high': 0, 'medium': 0, 'low': 0,
                'high_percent': 0, 'medium_percent': 0, 'low_percent': 0
            }
    
    def generate_pdf_report(self, data: Dict[str, Any]) -> bytes:
        """PDF 보고서 생성"""
        if not REPORTLAB_AVAILABLE:
            raise ImportError("ReportLab 라이브러리가 설치되지 않았습니다.")
        
        try:
            buffer = BytesIO()
            doc = SimpleDocTemplate(
                buffer,
                pagesize=A4,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18
            )
            
            # 스타일 설정
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'CustomTitle',
                parent=styles['Heading1'],
                fontSize=18,
                spaceAfter=30,
                alignment=1  # 중앙 정렬
            )
            
            story = []
            
            # 제목 페이지
            title = Paragraph(f"{self.report_date.year}년도 위험성평가 결과서", title_style)
            story.append(title)
            story.append(Spacer(1, 12))
            
            # 회사 정보
            title_data = data.get('title_data', data.get('표지', {}))
            company_info = [
                ['항목', '내용'],
                ['회사명', title_data.get('company_name', '')],
                ['주소', title_data.get('address', title_data.get('company_address', ''))],
                ['전화번호', title_data.get('telephone', title_data.get('company_tel', ''))],
                ['대표자', title_data.get('ceo_name', '')],
                ['작성일', self.report_date.strftime('%Y-%m-%d')]
            ]
            
            company_table = Table(company_info)
            company_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 14),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            story.append(company_table)
            story.append(PageBreak())
            
            # 위험성평가 결과
            assessment_data = []
            if hasattr(st.session_state, 'assessment_data') and st.session_state.assessment_data:
                assessment_data = st.session_state.assessment_data
            elif 'assessment_data' in data:
                assessment_data = data['assessment_data']
                
            if assessment_data:
                story.append(Paragraph("위험성평가 결과", styles['Heading2']))
                story.append(Spacer(1, 12))
                
                # 통계
                risk_stats = self._calculate_risk_statistics(assessment_data)
                stats_text = f"""
                전체 평가 항목: {risk_stats['total']}개<br/>
                - 높음 위험: {risk_stats['high']}개 ({risk_stats['high_percent']:.1f}%)<br/>
                - 보통 위험: {risk_stats['medium']}개 ({risk_stats['medium_percent']:.1f}%)<br/>
                - 낮음 위험: {risk_stats['low']}개 ({risk_stats['low_percent']:.1f}%)
                """
                story.append(Paragraph(stats_text, styles['Normal']))
            
            doc.build(story)
            buffer.seek(0)
            return buffer.getvalue()
            
        except Exception as e:
            logger.error(f"PDF 보고서 생성 오류: {e}")
            raise
    
    def create_risk_chart(self, assessment_data: List[Dict]) -> bytes:
        """위험성 평가 차트 생성"""
        if not PLOTLY_AVAILABLE:
            raise ImportError("Plotly 라이브러리가 설치되지 않았습니다.")
            
        try:
            if not assessment_data:
                return None
                
            risk_stats = self._calculate_risk_statistics(assessment_data)
            
            # 파이 차트 생성
            labels = ['높음', '보통', '낮음']
            values = [risk_stats['high'], risk_stats['medium'], risk_stats['low']]
            colors_list = ['#ff6b6b', '#ffd43b', '#51cf66']
            
            fig = go.Figure(data=[go.Pie(
                labels=labels, 
                values=values,
                marker_colors=colors_list,
                textinfo='label+percent',
                textfont_size=12
            )])
            
            fig.update_layout(
                title="위험성 평가 결과 분포",
                font=dict(size=14),
                showlegend=True
            )
            
            # 이미지로 변환
            img_bytes = fig.to_image(format="png", width=800, height=600)
            return img_bytes
            
        except Exception as e:
            logger.error(f"차트 생성 오류: {e}")
            return None

# Streamlit 사용을 위한 헬퍼 함수들
def display_report_options(data: Dict[str, Any]):
    """Streamlit에서 보고서 옵션 표시"""
    st.subheader("보고서 생성")
    
    col1, col2, col3 = st.columns(3)
    
    generator = RiskAssessmentReportGenerator()
    
    with col1:
        if st.button("엑셀 보고서 생성"):
            try:
                excel_data = generator.generate_excel_report(data)
                filename = f"위험성평가_보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                
                st.download_button(
                    label="엑셀 파일 다운로드",
                    data=excel_data,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("엑셀 보고서가 생성되었습니다!")
            except Exception as e:
                st.error(f"엑셀 보고서 생성 중 오류: {e}")
    
    with col2:
        if st.button("PDF 보고서 생성"):
            if REPORTLAB_AVAILABLE:
                try:
                    pdf_data = generator.generate_pdf_report(data)
                    filename = f"위험성평가_보고서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                    
                    st.download_button(
                        label="PDF 파일 다운로드",
                        data=pdf_data,
                        file_name=filename,
                        mime="application/pdf"
                    )
                    st.success("PDF 보고서가 생성되었습니다!")
                except Exception as e:
                    st.error(f"PDF 보고서 생성 중 오류: {e}")
            else:
                st.error("ReportLab 라이브러리가 필요합니다.")
    
    with col3:
        if st.button("차트 생성"):
            if PLOTLY_AVAILABLE:
                try:
                    # 평가 데이터 가져오기
                    assessment_data = []
                    if hasattr(st.session_state, 'assessment_data') and st.session_state.assessment_data:
                        assessment_data = st.session_state.assessment_data
                    elif 'assessment_data' in data:
                        assessment_data = data['assessment_data']
                    
                    if assessment_data:
                        risk_stats = generator._calculate_risk_statistics(assessment_data)
                        
                        # Plotly 차트 생성
                        labels = ['높음', '보통', '낮음']
                        values = [risk_stats['high'], risk_stats['medium'], risk_stats['low']]
                        
                        fig = go.Figure(data=[go.Pie(
                            labels=labels, 
                            values=values,
                            marker_colors=['#ff6b6b', '#ffd43b', '#51cf66'],
                            textinfo='label+percent'
                        )])
                        
                        fig.update_layout(title="위험성 평가 결과 분포")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning("차트를 생성할 평가 데이터가 없습니다.")
                        
                except Exception as e:
                    st.error(f"차트 생성 중 오류: {e}")
            else:
                st.error("Plotly 라이브러리가 설치되지 않았습니다.")

def create_comprehensive_report(data: Dict[str, Any], report_type: str = 'excel') -> bytes:
    """통합 보고서 생성 함수"""
    generator = RiskAssessmentReportGenerator()
    
    if report_type == 'excel':
        return generator.generate_excel_report(data)
    elif report_type == 'pdf':
        return generator.generate_pdf_report(data)
    else:
        raise ValueError(f"지원하지 않는 보고서 타입: {report_type}")
