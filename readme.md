# 위험성평가 시스템 (Risk Assessment System)

## 📋 프로젝트 개요

이 프로젝트는 산업안전보건법에 따른 위험성평가를 체계적으로 수행할 수 있도록 지원하는 웹 기반 시스템입니다. 안전보건컨설팅 전문가들이 현장에서 쉽게 사용할 수 있도록 설계되었습니다.

### 주요 특징
- 📊 단계별 위험성평가 프로세스 지원
- 📈 자동 위험성 계산 (5×4 매트릭스)
- 📄 Excel/PDF 보고서 자동 생성
- ⚖️ 법적근거 데이터베이스 연동
- 🎨 직관적인 웹 인터페이스
- 💾 데이터 자동 저장 및 복원

## 🛠️ 기술 스택

- **Backend**: Python 3.8+
- **Frontend**: Streamlit + HTML/CSS/JavaScript
- **데이터 처리**: Pandas, OpenPyXL
- **보고서 생성**: XlsxWriter, ReportLab
- **시각화**: Plotly
- **문서 처리**: Python-docx

## 📦 설치 방법

### 1. 저장소 클론
```bash
git clone https://github.com/your-username/risk-assessment-system.git
cd risk-assessment-system
```

### 2. 가상환경 생성 (권장)
```bash
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate
```

### 3. 의존성 설치
```bash
pip install -r requirements.txt
```

### 4. 애플리케이션 실행
```bash
streamlit run streamlit_app.py
```

브라우저에서 `http://localhost:8501`로 접속하여 시스템을 사용할 수 있습니다.

## 📁 프로젝트 구조

```
risk-assessment-system/
│
├── streamlit_app.py              # 메인 Streamlit 애플리케이션
├── data_handler.py               # 데이터 처리 및 관리 모듈
├── report_generator.py           # 보고서 생성 모듈
├── utils_data_handler.py         # 유틸리티 데이터 핸들러
├── requirements.txt              # Python 의존성 목록
├── README.md                     # 프로젝트 문서 (이 파일)
├── .gitignore                   # Git 무시 파일 목록
│
├── html_files/                   # HTML 템플릿 파일들
│   ├── integrated-risk-assessment.html    # 통합 시스템
│   ├── RA(1)-title.html                  # 표지
│   ├── RA(2)-WorkProcess.html            # 작업공정
│   ├── RA(3)-RiskInfo.html               # 위험정보
│   └── RA(4)_riskAssessment.html         # 위험성평가표
│
└── data/                         # 데이터 파일들
    └── DBRAclassOSHregulatory.xlsx       # 법적근거 데이터베이스
```

## 🚀 사용 방법

### 1. 시스템 시작
```bash
streamlit run streamlit_app.py
```

### 2. 위험성평가 단계별 진행

#### 단계 1: 표지 작성
- 회사 기본 정보 입력
- 결재란 작성

#### 단계 2: 작업공정 분석
- 사업장 개요 입력
- 공정도 작성 (공정별 상세 정보)

#### 단계 3: 위험정보 수집
- 업종명 입력
- 원재료, 기계기구, 유해화학물질 정보

#### 단계 4: 유해위험요인 분류
- 기계적, 전기적, 화학적, 생물학적, 작업특성, 작업환경 요인별 분류

#### 단계 5: 위험성평가표 작성
- 위험요인별 가능성과 중대성 평가
- 자동 위험성 계산 (5×4 매트릭스)
- 법적근거 자동 매칭

#### 단계 6: 위험감소대책 수립
- 위험성 등급별 감소대책 계획

#### 단계 7: 개선활동 관리
- 개선 전후 위험성 비교
- 개선활동 추적

### 3. 보고서 생성
- Excel 보고서: 전체 평가 결과를 체계적으로 정리
- PDF 보고서: 요약 보고서 형태로 출력

## ⚙️ 설정 및 커스터마이징

### 법적근거 데이터베이스 업데이트
1. `data/DBRAclassOSHregulatory.xlsx` 파일 수정
2. 파일 업로드 기능을 통해 새로운 법적근거 데이터 추가

### 위험성 계산 매트릭스 변경
`data_handler.py`의 `calculate_risk_score` 메서드에서 위험성 등급 기준을 수정할 수 있습니다.

```python
def calculate_risk_score(self, possibility: int, severity: int) -> Dict[str, Any]:
    score = possibility * severity
    
    # 위험성 등급 기준 수정 가능
    if score >= 12:
        level = "높음"
        action = "즉시 개선/작업중지"
    elif score >= 6:
        level = "보통"
        action = "개선"
    else:
        level = "낮음"
        action = "모니터링"
```

## 📊 위험성 평가 매트릭스

본 시스템은 5×4 매트릭스를 사용합니다:

| 가능성/중대성 | 매우위험(4) | 위험(3) | 보통(2) | 양호(1) |
|---------------|-------------|---------|---------|---------|
| 매우자주(5)   | 20 (높음)   | 15 (높음) | 10 (보통) | 5 (낮음) |
| 자주(4)       | 16 (높음)   | 12 (높음) | 8 (보통)  | 4 (낮음) |
| 보통(3)       | 12 (높음)   | 9 (보통)  | 6 (보통)  | 3 (낮음) |
| 가끔(2)       | 8 (보통)    | 6 (보통)  | 4 (낮음)  | 2 (낮음) |
| 매우가끔(1)   | 4 (낮음)    | 3 (낮음)  | 2 (낮음)  | 1 (낮음) |

### 위험성 등급별 조치
- **높음 (12~20)**: 즉시 개선/작업중지
- **보통 (6~11)**: 개선
- **낮음 (1~5)**: 모니터링

## 🔧 문제 해결

### 자주 발생하는 문제들

#### 1. 모듈 import 오류
```bash
ModuleNotFoundError: No module named 'streamlit'
```
**해결방법**: 의존성을 다시 설치해주세요.
```bash
pip install -r requirements.txt
```

#### 2. 한글 폰트 문제 (PDF 생성 시)
**해결방법**: 시스템에 한글 폰트가 설치되어 있는지 확인하고, `report_generator.py`에서 폰트 경로를 수정해주세요.

#### 3. Excel 파일 열기 오류
**해결방법**: 생성된 Excel 파일이 다른 프로그램에서 열려있지 않은지 확인해주세요.

#### 4. 메모리 부족 오류
**해결방법**: 대용량 데이터 처리 시 발생할 수 있습니다. 평가 항목을 나누어 처리하거나 시스템 메모리를 확인해주세요.

### 로그 확인
시스템 로그는 콘솔에서 확인할 수 있습니다. 문제 발생 시 로그를 먼저 확인해주세요.

## 🤝 기여 방법

1. 이 저장소를 Fork합니다
2. 새로운 기능 브랜치를 생성합니다 (`git checkout -b feature/새기능`)
3. 변경사항을 커밋합니다 (`git commit -am '새 기능 추가'`)
4. 브랜치에 푸시합니다 (`git push origin feature/새기능`)
5. Pull Request를 생성합니다

### 코딩 규칙
- Python PEP 8 스타일 가이드를 따라주세요
- 함수와 클래스에는 docstring을 작성해주세요
- 새로운 기능 추가 시 테스트 코드를 포함해주세요

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 `LICENSE` 파일을 참조하세요.

## 📞 지원 및 연락처

- **이슈 신고**: [GitHub Issues](https://github.com/your-username/risk-assessment-system/issues)
- **기능 제안**: [GitHub Discussions](https://github.com/your-username/risk-assessment-system/discussions)
- **이메일**: your-email@example.com

## 📚 추가 자료

### 관련 법령 및 지침
- [산업안전보건법](https://www.law.go.kr/법령/산업안전보건법)
- [산업안전보건기준에 관한 규칙](https://www.law.go.kr/법령/산업안전보건기준에관한규칙)
- [위험성평가에 관한 지침](https://www.kosha.or.kr)

### 기술 문서
- [Streamlit 공식 문서](https://docs.streamlit.io)
- [Pandas 사용법](https://pandas.pydata.org/docs/)
- [Plotly 차트 생성](https://plotly.com/python/)

## 🔄 업데이트 내역

### v1.0.0 (2025-01-01)
- 초기 버전 릴리스
- 기본 위험성평가 기능 구현
- Excel/PDF 보고서 생성 기능
- 법적근거 데이터베이스 연동

### 향후 계획
- [ ] 모바일 반응형 인터페이스 개선
- [ ] 클라우드 데이터베이스 연동
- [ ] 다국어 지원 (영어, 중국어)
- [ ] API 서버 분리
- [ ] 사용자 권한 관리 시스템
- [ ] 대시보드 기능 강화

## 🙏 감사의 말

이 프로젝트는 현장에서 위험성평가를 수행하는 모든 안전보건 전문가들의 업무 효율성 향상을 위해 개발되었습니다. 여러분의 피드백과 기여가 이 프로젝트를 더욱 발전시킬 수 있습니다.

---

**⚠️ 주의사항**: 이 시스템은 위험성평가 업무를 지원하는 도구이며, 실제 현장에서는 전문가의 판단과 함께 사용되어야 합니다. 모든 법적 책임은 사용자에게 있습니다.