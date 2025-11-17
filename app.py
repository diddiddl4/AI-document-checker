import streamlit as st
import os
import openpyxl
from openpyxl.utils import get_column_letter
from docx import Document
from pptx import Presentation
import PyPDF2
from datetime import datetime
import io

# 페이지 설정
st.set_page_config(
    page_title="AI 문서 점검기 - 경원알미늄",
    page_icon="🔍",
    layout="wide"
)

# CSS 스타일
st.markdown("""
<style>
    .main {
        background: linear-gradient(135deg, #2193b0 0%, #6dd5ed 100%);
    }
    .stApp {
        background: linear-gradient(135deg, #2193b0 0%, #6dd5ed 100%);
    }
    h1 {
        color: white;
        text-align: center;
    }
    .footer {
        position: fixed;
        bottom: 0;
        left: 0;
        right: 0;
        background: linear-gradient(135deg, #2193b0 0%, #6dd5ed 100%);
        color: white;
        text-align: center;
        padding: 10px;
    }
</style>
""", unsafe_allow_html=True)

class DocumentAnalyzer:
    def __init__(self, filepath, mode='standard'):
        self.filepath = filepath
        self.file_ext = os.path.splitext(filepath)[1].lower()
        self.mode = mode
        self.issues = []
        self.warnings = []
        self.score = 100
        self.cell_issues = []
        
    def analyze(self):
        if self.file_ext in ['.xlsx', '.xls']:
            return self._analyze_excel()
        elif self.file_ext in ['.docx', '.doc']:
            return self._analyze_word()
        elif self.file_ext in ['.pptx', '.ppt']:
            return self._analyze_ppt()
        elif self.file_ext == '.pdf':
            return self._analyze_pdf()
        return self._get_result()
    
    def _analyze_excel(self):
        try:
            wb = openpyxl.load_workbook(self.filepath, data_only=False)
            
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                
                # 병합 셀 검사
                merged_cells = list(sheet.merged_cells.ranges)
                if merged_cells:
                    self.score -= len(merged_cells) * 3
                    for merged in merged_cells:
                        self.cell_issues.append({
                            'sheet': sheet_name,
                            'cell': str(merged),
                            'type': 'MERGED_CELL',
                            'severity': 'HIGH',
                            'message': f'병합된 셀: {merged}',
                            'recommendation': '병합 해제 후 데이터 정규화 필요'
                        })
                    self.issues.append({
                        'type': 'MERGED_CELLS',
                        'count': len(merged_cells),
                        'message': f'{len(merged_cells)}개의 병합 셀 발견'
                    })
                
                # 줄바꿈 검사
                newline_count = 0
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str) and '\n' in cell.value:
                            newline_count += 1
                
                if newline_count > 0:
                    self.warnings.append({
                        'type': 'NEWLINES',
                        'count': newline_count,
                        'message': f'{newline_count}개 셀에 줄바꿈 포함'
                    })
                
                # 숨겨진 행/열
                hidden_rows = [i for i in range(1, sheet.max_row + 1) if sheet.row_dimensions[i].hidden]
                hidden_cols = [i for i in range(1, sheet.max_column + 1) 
                              if sheet.column_dimensions[get_column_letter(i)].hidden]
                
                if hidden_rows or hidden_cols:
                    self.score -= 15
                    self.issues.append({
                        'type': 'HIDDEN_DATA',
                        'message': f'숨겨진 행 {len(hidden_rows)}개, 열 {len(hidden_cols)}개'
                    })
            
        except Exception as e:
            self.issues.append({'type': 'ERROR', 'message': str(e)})
        
        return self._get_result()
    
    def _analyze_word(self):
        try:
            doc = Document(self.filepath)
            table_count = len(doc.tables)
            if table_count > 0:
                self.warnings.append({
                    'type': 'TABLES',
                    'message': f'{table_count}개의 표 발견'
                })
        except Exception as e:
            self.issues.append({'type': 'ERROR', 'message': str(e)})
        return self._get_result()
    
    def _analyze_ppt(self):
        try:
            prs = Presentation(self.filepath)
            slide_count = len(prs.slides)
            if slide_count > 50:
                self.score -= 10
                self.warnings.append({
                    'type': 'MANY_SLIDES',
                    'message': f'{slide_count}개의 슬라이드'
                })
        except Exception as e:
            self.issues.append({'type': 'ERROR', 'message': str(e)})
        return self._get_result()
    
    def _analyze_pdf(self):
        try:
            pdf = PyPDF2.PdfReader(self.filepath)
            text_extractable = False
            for page in pdf.pages[:3]:
                if page.extract_text().strip():
                    text_extractable = True
                    break
            
            if not text_extractable:
                self.score -= 30
                self.issues.append({
                    'type': 'SCANNED_PDF',
                    'message': '스캔된 PDF - OCR 처리 필요'
                })
        except Exception as e:
            self.issues.append({'type': 'ERROR', 'message': str(e)})
        return self._get_result()
    
    def _get_result(self):
        self.score = max(0, self.score)
        
        if self.score >= 80:
            grade = 'A'
        elif self.score >= 60:
            grade = 'B'
        elif self.score >= 40:
            grade = 'C'
        else:
            grade = 'D'
        
        return {
            'score': self.score,
            'grade': grade,
            'issues': self.issues,
            'warnings': self.warnings,
            'cell_issues': self.cell_issues,
            'file_type': self.file_ext,
            'mode': self.mode
        }
    
    def generate_optimized_version(self):
        if self.file_ext not in ['.xlsx', '.xls']:
            return None
        
        try:
            wb = openpyxl.load_workbook(self.filepath)
            output = io.BytesIO()
            
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                
                # 병합 셀 해제 + 값 복사 + 서식 유지
                merged_ranges = list(sheet.merged_cells.ranges)
                for merged in merged_ranges:
                    min_col, min_row, max_col, max_row = merged.bounds
                    
                    source_cell = sheet.cell(min_row, min_col)
                    merged_value = source_cell.value
                    source_font = source_cell.font.copy() if source_cell.font else None
                    source_fill = source_cell.fill.copy() if source_cell.fill else None
                    source_border = source_cell.border.copy() if source_cell.border else None
                    source_alignment = source_cell.alignment.copy() if source_cell.alignment else None
                    
                    sheet.unmerge_cells(str(merged))
                    
                    for row in range(min_row, max_row + 1):
                        for col in range(min_col, max_col + 1):
                            cell = sheet.cell(row, col)
                            cell.value = merged_value
                            if source_font:
                                cell.font = source_font.copy()
                            if source_fill:
                                cell.fill = source_fill.copy()
                            if source_border:
                                cell.border = source_border.copy()
                            if source_alignment:
                                cell.alignment = source_alignment.copy()
                
                # 줄바꿈 제거
                for row in sheet.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            cell.value = cell.value.replace('\n', ' ')
                
                # 기호 변환 (분석 모드)
                if self.mode == 'analysis':
                    for row_idx, row in enumerate(sheet.iter_rows(), 1):
                        for col_idx, cell in enumerate(row, 1):
                            if row_idx > 1:
                                header_cell = sheet.cell(4, col_idx)
                                header = str(header_cell.value or '')
                                
                                if '여부' in header or '수령' in header:
                                    if cell.value in ['○', 'O', 'o', '●']:
                                        cell.value = '예'
                                    elif cell.value in ['', None, 'X', '×']:
                                        cell.value = '아니오'
                
                # 숨김 해제
                for i in range(1, sheet.max_row + 1):
                    sheet.row_dimensions[i].hidden = False
                for i in range(1, sheet.max_column + 1):
                    sheet.column_dimensions[get_column_letter(i)].hidden = False
            
            wb.save(output)
            output.seek(0)
            return output
            
        except Exception as e:
            st.error(f"최적화 오류: {e}")
            return None

# 메인 앱
st.title("🔍 AI 문서 점검기 Pro")
st.markdown("### 경원알미늄 - 탁월한 업무 시스템 구축 TFT")

# 모드 선택
col1, col2 = st.columns(2)
with col1:
    mode = st.radio(
        "최적화 모드",
        ["표준 모드", "분석 모드"],
        help="표준: 병합셀 해제 + 줄바꿈 제거 | 분석: 표준 + 기호변환"
    )

selected_mode = 'standard' if mode == "표준 모드" else 'analysis'

# 파일 업로드
uploaded_file = st.file_uploader(
    "파일을 선택하세요",
    type=['xlsx', 'xls', 'docx', 'doc', 'pptx', 'ppt', 'pdf'],
    help="Excel, Word, PowerPoint, PDF 지원"
)

if uploaded_file:
    # 임시 파일 저장
    with open(f"temp_{uploaded_file.name}", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    # 분석
    with st.spinner('분석 중...'):
        analyzer = DocumentAnalyzer(f"temp_{uploaded_file.name}", mode=selected_mode)
        result = analyzer.analyze()
    
    # 결과 표시
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("점수", f"{result['score']}점")
    with col2:
        st.metric("등급", result['grade'])
    with col3:
        st.metric("모드", "표준" if selected_mode == 'standard' else "분석")
    
    # 이슈 표시
    if result['issues']:
        st.subheader("🚨 주요 이슈")
        for issue in result['issues']:
            st.error(f"**{issue.get('type')}**: {issue.get('message')}")
    
    if result['warnings']:
        st.subheader("⚠️ 경고")
        for warning in result['warnings']:
            st.warning(f"**{warning.get('type')}**: {warning.get('message')}")
    
    if result['cell_issues']:
        st.subheader("📍 셀별 문제점")
        for cell_issue in result['cell_issues'][:10]:
            st.info(f"{cell_issue['sheet']} - {cell_issue['cell']}: {cell_issue['message']}")
    
    # 다운로드 버튼
    st.subheader("📥 다운로드")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if result['file_type'] in ['.xlsx', '.xls']:
            optimized = analyzer.generate_optimized_version()
            if optimized:
                st.download_button(
                    label="✨ AI 최적화 버전",
                    data=optimized,
                    file_name=f"AI최적화_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    
    # 임시 파일 삭제
    os.remove(f"temp_{uploaded_file.name}")

# 푸터
st.markdown("""
<div class="footer">
경원알미늄 - 탁월한 업무 시스템 구축 TFT
</div>
""", unsafe_allow_html=True)
