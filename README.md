# 🔍 AI 문서 점검기 Pro

**경원알미늄 - 탁월한 업무 시스템 구축 TFT**

AI가 문서를 제대로 인식할 수 있는지 사전 점검하는 도구

## 🚀 주요 기능

### Excel 분석
- ✅ 병합 셀 자동 검사 및 해제
- ✅ 줄바꿈 자동 제거
- ✅ 숨겨진 행/열 검사 및 표시
- ✅ 기호 자동 변환 (○→예, X→아니오)
- ✅ 표 서식 유지하며 최적화

### Word/PPT/PDF
- ✅ 표 구조 분석
- ✅ 슬라이드 개수 검사
- ✅ PDF 텍스트 추출 가능 여부 확인
- ✅ 스캔 PDF Claude OCR 처리

### 이미지 OCR
- ✅ JPG, PNG 텍스트 추출
- ✅ Claude AI 기반 고정확도 OCR
- ✅ 한글 최적화

### 공통 기능
- ✅ AI 가독성 점수 (0-100점)
- ✅ 등급 부여 (A~D)
- ✅ 문제점 상세 리포트
- ✅ 최적화 파일 자동 생성

## 📊 지원 형식

- Excel (.xlsx, .xls)
- Word (.docx, .doc)
- PowerPoint (.pptx, .ppt)
- PDF (.pdf)
- 이미지 (.jpg, .jpeg, .png)

## 💻 배포 방법

### 1. GitHub 레포지토리 생성

### 2. 파일 업로드
- app.py
- requirements.txt
- packages.txt
- README.md

### 3. Streamlit Cloud 배포
```
https://share.streamlit.io/
→ New app
→ Repository 선택
→ Deploy
```

### 4. API 키 설정
Settings → Secrets:
```
ANTHROPIC_API_KEY = "sk-ant-api03-xxxxx"
```

## 🎯 사용 방법

1. 파일 업로드
2. 모드 선택 (표준/분석)
3. 분석 시작
4. 최적화 파일 다운로드

## 📋 등급 기준

- **A등급 (80-100점)**: AI 처리 최적
- **B등급 (60-79점)**: AI 처리 가능
- **C등급 (40-59점)**: 개선 필요
- **D등급 (0-39점)**: 전면 개편 필요

## 🔧 모드 설명

### 표준 모드
- 병합 셀 해제
- 줄바꿈 제거
- 숨김 해제
- 서식 유지

### 분석 모드
- 표준 모드 전체 기능
- 기호 자동 변환 (○→예, X→아니오)
- "여부", "수령", "확인" 컬럼 자동 인식

---

© 2025 경원알미늄 TFT
