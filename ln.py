#!/usr/bin/env python3
"""
latex_windows.py - 고급 MinerU 딥러닝 문서 → LaTeX 변환 도구
사용법: python latex_windows.py 경로/파일명

🤖 완전한 딥러닝 파이프라인:
- 캐시 정리 🔄 [Deep Learning Pipeline Initialization]
- pywin32로 Word → PDF 변환 (PDF로 Word 문서 없이도 작업 가능)
- MinerU 딥러닝 도구로 📊 활성화된 딥러닝 모델:
   🧮 nougat-latex-ocr: 수식 이미지 → LaTeX 변환 (정확도 우선)
   🔤 PaddleOCR v5: 한글/영문 텍스트 인식 (자동 언어 감지)
   📊 rapid-table: 표 구조 분석
   🎯 Layout-YOLO: 문서 레이아웃 분석
   🔍 Object Detection: 객체/영역 탐지

🔥 고급 처리 과정:
1. MinerU로 수식, 문자, 표 영역 감지 및 분리
2. 감지된 수식 영역을 이미지로 추출
3. nougat-latex-ocr로 수식 이미지를 정확한 LaTeX로 변환
4. PDF 문서에서 한글과 영문을 읽고 문자로 변환
5. 문자와 LaTeX로 md 파일 생성, json 파일 생성
6. 3패널 뷰어: PDF이미지 | LaTeX소스 | 렌더링
7. 각 패널별 스크롤 (상하좌우) + 축소확대기능 + 페이지별 보기

지원 형식:
- Word 파일 (.docx, .doc) → PDF → LaTeX
- PDF 파일 (.pdf) → LaTeX

예시:
    python latex_windows.py C:/test/document.pdf
    python latex_windows.py paper.docx
"""

import sys
import os
import subprocess
import time
import json
import tempfile
import shutil
import webbrowser
from pathlib import Path
from datetime import datetime, timedelta
import re
import argparse

# 패치들은 필요 없음 - MinerU는 문서 구조 분석용으로만 사용

# 최적화된 3패널 뷰어
from ln_final_3panel_viewer import generate_optimized_3panel_viewer

# Windows 전용 라이브러리
try:
    import win32com.client
    import pythoncom
    HAS_PYWIN32 = True
except ImportError:
    HAS_PYWIN32 = False
    print("⚠️ pywin32가 설치되어 있지 않습니다. Word 변환 기능이 제한됩니다.")


class CacheManager:
    """캐시 정리 및 시스템 최적화"""
    
    def __init__(self):
        self.temp_dirs = [
            os.environ.get('TEMP', ''),
            os.environ.get('TMP', ''),
            './temp',
            './cache'
        ]
    
    def clear_cache(self):
        """🔄 Deep Learning Pipeline Initialization"""
        print("🔄 [Deep Learning Pipeline Initialization] 시작...")
        
        cleared_size = 0
        cleared_files = 0
        
        for temp_dir in self.temp_dirs:
            if temp_dir and Path(temp_dir).exists():
                try:
                    for item in Path(temp_dir).iterdir():
                        if item.name.startswith(('mineru_', 'latex_', 'temp_')):
                            if item.is_file():
                                size = item.stat().st_size
                                item.unlink()
                                cleared_size += size
                                cleared_files += 1
                            elif item.is_dir():
                                shutil.rmtree(item)
                                cleared_files += 1
                except PermissionError:
                    continue
        
        if cleared_files > 0:
            print(f"✅ 캐시 정리 완료: {cleared_files}개 파일/폴더, {cleared_size/1024/1024:.1f}MB")
        else:
            print("✅ 캐시 정리 완료: 정리할 항목 없음")

class WordToPDFConverter:
    """pywin32를 사용한 Word → PDF 변환"""
    
    def __init__(self):
        self.word_app = None
    
    def convert(self, word_path, pdf_path=None):
        """Word 문서를 PDF로 변환"""
        if not HAS_PYWIN32:
            print("❌ pywin32가 필요합니다: pip install pywin32")
            return None
        
        word_path = Path(word_path).resolve()
        if not word_path.exists():
            print(f"❌ Word 파일을 찾을 수 없습니다: {word_path}")
            return None
        
        if pdf_path is None:
            pdf_path = word_path.with_suffix('.pdf')
        else:
            pdf_path = Path(pdf_path).resolve()
        
        print(f"📄 Word → PDF 변환: {word_path.name}")
        
        try:
            pythoncom.CoInitialize()
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
            
            doc = self.word_app.Documents.Open(str(word_path))
            doc.SaveAs(str(pdf_path), FileFormat=17)  # PDF 형식
            doc.Close()
            
            print(f"✅ PDF 변환 완료: {pdf_path}")
            return pdf_path
            
        except Exception as e:
            print(f"❌ Word 변환 실패: {e}")
            return None
        finally:
            if self.word_app:
                self.word_app.Quit()
                pythoncom.CoUninitialize()

class PipelineTimer:
    """파이프라인별 실행 시간 측정 클래스"""
    
    def __init__(self):
        self.stages = {}
        self.total_start_time = None
        self.current_stage = None
        self.current_start_time = None
    
    def start_total(self):
        self.total_start_time = time.time()
        start_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"⏱️ Pipeline started at: {start_time_str}")
    
    def start_stage(self, stage_name):
        if self.current_stage:
            self.end_stage()
        
        self.current_stage = stage_name
        self.current_start_time = time.time()
        print(f"🔄 [{stage_name}] Starting...")
    
    def end_stage(self):
        if self.current_stage and self.current_start_time:
            elapsed = time.time() - self.current_start_time
            self.stages[self.current_stage] = elapsed
            
            if elapsed < 60:
                time_str = f"{elapsed:.1f}s"
            elif elapsed < 3600:
                mins = int(elapsed // 60)
                secs = elapsed % 60
                time_str = f"{mins}m {secs:.1f}s"
            else:
                hours = int(elapsed // 3600)
                mins = int((elapsed % 3600) // 60)
                secs = elapsed % 60
                time_str = f"{hours}h {mins}m {secs:.1f}s"
            
            print(f"✅ [{self.current_stage}] Completed - Duration: {time_str}")
            
            self.current_stage = None
            self.current_start_time = None
    
    def end_total(self):
        if self.current_stage:
            self.end_stage()
        
        if self.total_start_time:
            total_elapsed = time.time() - self.total_start_time
            end_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            print("\n" + "="*70)
            print("📊 Pipeline Execution Time Summary")
            print("="*70)
            print(f"🏁 Finished at: {end_time_str}")
            
            if total_elapsed < 60:
                total_time_str = f"{total_elapsed:.1f}s"
            elif total_elapsed < 3600:
                mins = int(total_elapsed // 60)
                secs = total_elapsed % 60
                total_time_str = f"{mins}m {secs:.1f}s"
            else:
                hours = int(total_elapsed // 3600)
                mins = int((total_elapsed % 3600) // 60)
                secs = total_elapsed % 60
                total_time_str = f"{hours}h {mins}m {secs:.1f}s"
            
            print(f"⏰ Total Duration: {total_time_str}")
            
            if self.stages:
                print("\n📋 Stage Breakdown:")
                for stage, duration in self.stages.items():
                    if duration < 60:
                        stage_time_str = f"{duration:.1f}s"
                    elif duration < 3600:
                        mins = int(duration // 60)
                        secs = duration % 60
                        stage_time_str = f"{mins}m {secs:.1f}s"
                    else:
                        hours = int(duration // 3600)
                        mins = int((duration % 3600) // 60)
                        secs = duration % 60
                        stage_time_str = f"{hours}h {mins}m {secs:.1f}s"
                    
                    percentage = (duration / total_elapsed) * 100
                    print(f"   🔸 {stage:<25} {stage_time_str:>8} ({percentage:5.1f}%)")
            
            print("="*70)

class MinerUProcessor:
    """MinerU 딥러닝 파이프라인 프로세서"""
    
    def __init__(self):
        self.models_info = {
            "UniMERNet": "🧮 수식 → LaTeX 직접 변환 (810MB)",
            "PaddleOCR v5": "🔤 한글/영문 텍스트 인식 (자동 언어 감지)",
            "rapid-table": "📊 표 구조 분석",
            "Layout-YOLO": "🎯 문서 레이아웃 분석",
            "Object Detection": "🔍 객체/영역 탐지"
        }
    
    def show_models_status(self):
        """활성화된 딥러닝 모델 표시"""
        print("📊 활성화된 딥러닝 모델들:")
        for model, description in self.models_info.items():
            print(f"   {description}")
    
    def create_table_border_removed_pdf(self, pdf_path, output_dir):
        """표 테두리가 제거된 PDF 생성 (MinerU 처리 전)"""
        print("🔧 표 테두리 제거된 PDF 생성 중...")
        
        try:
            import pdfplumber
            from PIL import Image, ImageDraw
            import fitz  # PyMuPDF
            
            output_dir = Path(output_dir)
            processed_pdf_path = output_dir / f"{Path(pdf_path).stem}_table_borders_removed.pdf"
            
            # PyMuPDF로 새 PDF 문서 생성
            new_doc = fitz.open()
            
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages):
                    print(f"   📄 페이지 {page_num + 1} 테두리 제거 중...")
                    
                    # 페이지를 고해상도 이미지로 변환
                    page_image = page.to_image(resolution=300)
                    pil_image = page_image.original
                    
                    # 테이블 감지 및 테두리 제거
                    tables = page.extract_tables()
                    if tables:
                        print(f"      📊 {len(tables)}개 테이블 테두리 제거...")
                        draw = ImageDraw.Draw(pil_image)
                        
                        for table_idx, table in enumerate(tables):
                            table_finder = page.find_tables()
                            if table_idx < len(table_finder):
                                table_bbox = table_finder[table_idx].bbox
                                x0, y0, x1, y1 = table_bbox
                                
                                # 수평선 제거
                                for line in page.horizontal_edges:
                                    if (x0 <= line['x0'] <= x1 and x0 <= line['x1'] <= x1 and y0 <= line['y0'] <= y1):
                                        draw.line(
                                            [(line['x0'] * 300/72, line['y0'] * 300/72),
                                             (line['x1'] * 300/72, line['y1'] * 300/72)],
                                            fill='white', width=2  # 두께 증가
                                        )
                                
                                # 수직선 제거
                                for line in page.vertical_edges:
                                    if (x0 <= line['x0'] <= x1 and y0 <= line['y0'] <= y1 and y0 <= line['y1'] <= y1):
                                        draw.line(
                                            [(line['x0'] * 300/72, line['y0'] * 300/72),
                                             (line['x1'] * 300/72, line['y1'] * 300/72)],
                                            fill='white', width=2  # 두께 증가
                                        )
                    
                    # PIL 이미지를 바이트로 변환  
                    import io
                    img_bytes = io.BytesIO()
                    pil_image.save(img_bytes, format='PNG')
                    img_bytes.seek(0)
                    
                    # PyMuPDF에 이미지 페이지 추가
                    img_rect = fitz.Rect(0, 0, 595, 842)  # A4 크기
                    img_page = new_doc.new_page(width=595, height=842)
                    img_page.insert_image(img_rect, stream=img_bytes.getvalue())
            
            # 새 PDF 저장
            new_doc.save(str(processed_pdf_path))
            new_doc.close()
            
            print(f"✅ 표 테두리 제거된 PDF 생성: {processed_pdf_path}")
            return processed_pdf_path
            
        except Exception as e:
            print(f"⚠️ 표 테두리 제거 실패: {e}")
            print("📝 원본 PDF로 계속 진행")
            return pdf_path

    def extract_content_from_mineru_output(self, auto_dir):
        """MinerU 출력에서 content_list 추출"""
        try:
            content_list = []
            
            # 1. 먼저 *_uni_format.json 찾기 (MinerU 0.7.0b1)
            uni_files = list(auto_dir.glob("*_uni_format.json"))
            if uni_files:
                print(f"✅ uni_format.json 발견: {uni_files[0].name}")
                with open(uni_files[0], 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
            
            # 2. MD 파일에서 직접 추출
            md_files = list(auto_dir.glob("*.md"))
            if not md_files:
                print("⚠️ MD 파일을 찾을 수 없습니다.")
                return None
            
            with open(md_files[0], 'r', encoding='utf-8') as f:
                md_content = f.read()
            
            # 수식과 표 추출
            import re
            
            # 페이지 정보 파악 (pages_info.json이 있으면 사용)
            total_pages = 1
            pages_info_file = auto_dir.parent.parent / "pages_info.json"
            if pages_info_file.exists():
                with open(pages_info_file, 'r', encoding='utf-8') as f:
                    pages_data = json.load(f)
                    total_pages = len(pages_data)
            
            # 모든 컨텐츠 항목과 위치 정보 수집
            all_items = []
            
            # 블록 수식
            block_formulas = re.finditer(r'\$\$(.*?)\$\$', md_content, re.DOTALL)
            for i, match in enumerate(block_formulas):
                all_items.append({
                    "type": "equation",
                    "text": match.group(1).strip(),
                    "start": match.start(),
                    "end": match.end()
                })
            
            # 인라인 수식
            inline_formulas = re.finditer(r'(?<!\$)\$([^$\n]+)\$(?!\$)', md_content)
            for match in inline_formulas:
                # 블록 수식과 겹치지 않는지 확인
                is_in_block = any(
                    item['type'] == 'equation' and 
                    item['start'] <= match.start() < item['end'] 
                    for item in all_items
                )
                if not is_in_block:
                    all_items.append({
                        "type": "interline_equation", 
                        "text": match.group(1).strip(),
                        "start": match.start(),
                        "end": match.end()
                    })
            
            # 표 이미지
            table_imgs = re.finditer(r'!\[\]\((images/[^)]+)\)', md_content)
            for match in table_imgs:
                all_items.append({
                    "type": "table",
                    "img_path": match.group(1),
                    "text": "표 내용",
                    "start": match.start(),
                    "end": match.end()
                })
            
            # 위치 순으로 정렬
            all_items.sort(key=lambda x: x['start'])
            
            # 페이지별로 컨텐츠 분할
            if total_pages > 1:
                # MD 컨텐츠를 페이지 수로 균등 분할
                content_length = len(md_content)
                page_size = content_length // total_pages
                
                for item in all_items:
                    # 현재 아이템의 위치로 페이지 추정
                    estimated_page = min(item['start'] // page_size, total_pages - 1)
                    
                    content_item = {
                        "type": item["type"],
                        "page_idx": estimated_page,
                        "bbox": []  # bbox 정보는 model.json에서 가져와야 함
                    }
                    
                    if "text" in item:
                        content_item["text"] = item["text"]
                    if "img_path" in item:
                        content_item["img_path"] = item["img_path"]
                    
                    content_list.append(content_item)
            else:
                # 단일 페이지
                for item in all_items:
                    content_item = {
                        "type": item["type"],
                        "page_idx": 0,
                        "bbox": []
                    }
                    
                    if "text" in item:
                        content_item["text"] = item["text"]
                    if "img_path" in item:
                        content_item["img_path"] = item["img_path"]
                    
                    content_list.append(content_item)
            
            # 페이지별 통계 출력
            page_stats = {}
            for item in content_list:
                page = item.get('page_idx', 0)
                if page not in page_stats:
                    page_stats[page] = {'equation': 0, 'interline_equation': 0, 'table': 0}
                page_stats[page][item['type']] += 1
            
            print(f"📋 추출된 컨텐츠: {len(content_list)}개")
            for page, stats in sorted(page_stats.items()):
                equations = stats['equation'] + stats['interline_equation']
                tables = stats['table']
                print(f"   📄 페이지 {page + 1}: 수식 {equations}개, 표 {tables}개")
            
            return content_list
            
        except Exception as e:
            print(f"❌ content_list 추출 실패: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def extract_latex_from_md(self, md_content, page_idx):
        """MD 내용에서 LaTeX 수식 추출"""
        # 간단한 구현 - $$ ... $$ 패턴 찾기
        import re
        formulas = re.findall(r'\$\$(.*?)\$\$', md_content, re.DOTALL)
        if formulas and page_idx < len(formulas):
            return formulas[page_idx].strip()
        
        # 인라인 수식
        inline_formulas = re.findall(r'\$([^$]+)\$', md_content)
        if inline_formulas:
            return inline_formulas[0] if inline_formulas else ""
        
        return ""
    
    def enhance_formulas_with_nougat(self, output_dir, pdf_path=None):
        """nougat-latex-ocr로 수식 개선 - 모든 경우에 대응"""
        try:
            import subprocess
            import json
            from pathlib import Path
            
            print("\n🔄 UniMERNet → nougat-latex-ocr 수식 개선...")
            
            # nougat 설정
            nougat_python = "C:/git/nougat-latex-ocr/venv/Scripts/python.exe"
            nougat_path = "C:/git/nougat-latex-ocr/nougat-latex-ocr"
            
            if not Path(nougat_python).exists():
                print("❌ nougat 가상환경을 찾을 수 없습니다.")
                print("   💡 C:\\git\\nougat-latex-ocr\\venv\\ 경로를 확인하세요.")
                return
            
            # 전체 처리 시간 측정
            total_start_time = time.time()
            equation_count = 0
            processed_images = set()  # 중복 처리 방지
            
            # PDF에서 직접 수식 이미지 추출 (메인 방법)
            if pdf_path:
                print("\n   🔍 PDF에서 수식 이미지 직접 추출 시도...")
                
                # 방법 1: 기존 extract_formula_images 사용
                try:
                    from extract_formula_images import extract_formula_images_from_pdf, extract_from_content_list
                    
                    # 수식 이미지 저장 디렉토리
                    formula_images_dir = Path(output_dir) / "formula_images"
                    formula_images_dir.mkdir(parents=True, exist_ok=True)
                    
                    # model.json 또는 content_list.json 찾기
                    json_candidates = [
                        (Path(output_dir) / "model.json", "model"),
                        *[(f, "content") for f in Path(output_dir).glob("**/*content_list*.json")]
                    ]
                    
                    for json_path, json_type in json_candidates:
                        if json_path.exists():
                            print(f"   📄 {json_path.name} 사용하여 수식 추출")
                            try:
                                if json_type == "content":
                                    formula_info = extract_from_content_list(pdf_path, json_path, formula_images_dir)
                                else:
                                    formula_info = extract_formula_images_from_pdf(pdf_path, json_path, formula_images_dir)
                                
                                # 추출된 이미지 처리
                                for page_idx, formulas in formula_info.items():
                                    for formula in formulas:
                                        img_path = Path(formula['path'])
                                        if img_path.exists() and str(img_path) not in processed_images:
                                            processed_images.add(str(img_path))
                                            equation_count += 1
                                            latex_result = self._process_single_equation(img_path, equation_count, nougat_python, nougat_path, output_dir)
                                            if latex_result:
                                                # 결과 저장
                                                formula['nougat_latex'] = latex_result
                                
                                # 결과 저장
                                if equation_count > 0:
                                    result_path = formula_images_dir / "nougat_results.json"
                                    with open(result_path, 'w', encoding='utf-8') as f:
                                        json.dump(formula_info, f, ensure_ascii=False, indent=2)
                                    print(f"   ✅ nougat 결과 저장: {result_path}")
                                
                                break
                            except Exception as e:
                                print(f"   ⚠️ 수식 추출 실패: {e}")
                                import traceback
                                traceback.print_exc()
                except ImportError:
                    print("   ⚠️ extract_formula_images 모듈을 찾을 수 없습니다.")
                
                # 방법 2: 강제 추출 (방법 1이 실패했거나 수식을 못 찾은 경우)
                if equation_count == 0:
                    print("\n   🔧 강제 수식 추출 시도...")
                    try:
                        from force_extract_formulas import force_extract_formula_images, extract_from_model_json
                        
                        # content_list.json 사용한 강제 추출
                        forced_dir, forced_count = force_extract_formula_images(pdf_path, output_dir)
                        
                        if forced_count == 0:
                            # model.json 사용한 추출 시도
                            model_dir, model_count = extract_from_model_json(pdf_path, output_dir)
                            if model_count > 0:
                                forced_dir = model_dir
                                forced_count = model_count
                        
                        if forced_count > 0 and forced_dir:
                            # 추출된 이미지 처리
                            info_files = list(Path(forced_dir).glob("*_info.json"))
                            if info_files:
                                with open(info_files[0], 'r', encoding='utf-8') as f:
                                    forced_info = json.load(f)
                                
                                for page_idx, formulas in forced_info.items():
                                    for formula in formulas:
                                        img_path = Path(formula['path'])
                                        if img_path.exists():
                                            equation_count += 1
                                            latex_result = self._process_single_equation(img_path, equation_count, nougat_python, nougat_path, output_dir)
                                            if latex_result:
                                                formula['nougat_latex'] = latex_result
                                
                                # 결과 저장
                                result_path = Path(forced_dir) / "nougat_results.json"
                                with open(result_path, 'w', encoding='utf-8') as f:
                                    json.dump(forced_info, f, ensure_ascii=False, indent=2)
                                print(f"   ✅ 강제 추출 nougat 결과 저장: {result_path}")
                    
                    except Exception as e:
                        print(f"   ⚠️ 강제 추출 실패: {e}")
                        import traceback
                        traceback.print_exc()
                
                # 방법 3: 사전 추출 (아직도 수식이 없다면)
                if equation_count == 0:
                    print("\n   🔍 사전 수식 감지 시도...")
                    try:
                        from pre_extract_formulas import detect_formula_regions, extract_formulas_using_text_blocks
                        
                        # 이미지 기반 감지
                        pre_dir1, pre_count1 = detect_formula_regions(pdf_path, output_dir)
                        
                        # 텍스트 기반 감지
                        pre_dir2, pre_count2 = extract_formulas_using_text_blocks(pdf_path, output_dir)
                        
                        # 두 방법 중 더 많이 찾은 것 사용
                        if pre_count1 > 0 or pre_count2 > 0:
                            use_dir = pre_dir1 if pre_count1 > pre_count2 else pre_dir2
                            info_files = list(Path(use_dir).glob("*_info.json"))
                            
                            if info_files:
                                with open(info_files[0], 'r', encoding='utf-8') as f:
                                    pre_info = json.load(f)
                                
                                for page_idx, formulas in pre_info.items():
                                    for formula in formulas:
                                        img_path = Path(formula['path'])
                                        if img_path.exists():
                                            equation_count += 1
                                            latex_result = self._process_single_equation(img_path, equation_count, nougat_python, nougat_path, output_dir)
                                            if latex_result:
                                                formula['nougat_latex'] = latex_result
                    
                    except Exception as e:
                        print(f"   ⚠️ 사전 추출 실패: {e}")
                        import traceback
                        traceback.print_exc()
            
            # 총 통계 및 폴더 열기
            total_time = time.time() - total_start_time
            if equation_count > 0:
                print(f"\n   ✅ 총 {equation_count}개 이미지 nougat-latex-ocr 처리 완료")
                print(f"   ⏱️  총 소요시간: {total_time:.1f}초 (평균 {total_time/equation_count:.1f}초/이미지)")
                
                # 향상된 위치 기반 MD 파일 생성
                print("\n   🔄 위치 정보 기반 MD 파일 생성 중...")
                try:
                    from enhanced_md_generator import create_final_md
                    
                    # nougat 결과가 있는지 확인
                    nougat_result_files = list(Path(output_dir).glob("**/nougat_results.json"))
                    
                    if nougat_result_files:
                        # 향상된 MD 생성
                        new_md_path = create_final_md(pdf_path, output_dir)
                        
                        print(f"\n   ✅ 위치 기반 MD 파일 생성 완료!")
                        print(f"   📄 파일: {new_md_path}")
                        
                        # 기존 MD 백업
                        original_md_files = list(Path(output_dir).glob("*.md"))
                        for md_file in original_md_files:
                            if 'enhanced' not in md_file.name and 'nougat' not in md_file.name:
                                backup_path = md_file.with_suffix('.md.backup')
                                import shutil
                                shutil.copy2(md_file, backup_path)
                                print(f"   💾 원본 백업: {backup_path}")
                        
                except Exception as e:
                    print(f"   ⚠️ MD 생성 중 오류: {e}")
                    import traceback
                    traceback.print_exc()
                    
                    # 대체 방법: 기존 교체 방식 사용
                    try:
                        from replace_mineru_with_nougat import replace_mineru_formulas_with_nougat, update_md_with_nougat_results
                        
                        content_list_files = list(Path(output_dir).glob("**/content_list*.json"))
                        if nougat_result_files and content_list_files:
                            content_list_path = content_list_files[0]
                            replaced_count, new_content_path = replace_mineru_formulas_with_nougat(
                                content_list_path,
                                nougat_result_files[0]
                            )
                            
                            if replaced_count > 0:
                                md_files = list(Path(output_dir).glob("*.md"))
                                if md_files:
                                    md_path = update_md_with_nougat_results(md_files[0], new_content_path)
                                    print(f"   ✅ 대체 방법으로 MD 업데이트: {md_path}")
                    except:
                        pass
                
                # 수식 이미지 폴더 자동 열기
                self._open_images_folder(output_dir)
            else:
                print("⚠️ 처리할 수식 이미지를 찾을 수 없습니다.")
                print("   💡 MinerU가 수식을 탐지하지 못했거나, 이미 LaTeX 변환이 완료된 상태일 수 있습니다.")
                
                # images 폴더가 있으면 열기 (표 이미지 폴더)
                for img_dir in Path(output_dir).glob("**/images"):
                    if img_dir.exists() and any(img_dir.iterdir()):
                        print(f"\n📁 표(table) 이미지 폴더 열기: {img_dir}")
                        try:
                            if sys.platform == 'win32':
                                os.startfile(str(img_dir))
                                print("   ✅ 폴더가 열렸습니다.")
                            elif sys.platform == 'darwin':
                                subprocess.run(['open', str(img_dir)])
                            else:
                                subprocess.run(['xdg-open', str(img_dir)])
                        except Exception as e:
                            print(f"   ⚠️ 폴더 열기 실패: {e}")
                        break
                
        except Exception as e:
            print(f"❌ nougat-latex-ocr 처리 중 오류: {e}")
            import traceback
            traceback.print_exc()
    
    def _process_single_equation(self, img_path, equation_count, nougat_python, nougat_path, output_dir=None):
        """단일 수식 이미지 처리"""
        try:
            print(f"\n   🧮 수식 {equation_count} 처리 시작: {img_path.name}")
            print(f"      📍 이미지 크기: {img_path.stat().st_size / 1024:.1f}KB")
            print(f"      🚀 nougat-latex-ocr 실행중...")
            
            # 시간 측정 시작
            start_time = time.time()
            
            # output_dir 기본값 설정
            if output_dir is None:
                output_dir = img_path.parent
            
            # 임시 출력 파일
            temp_output = Path(output_dir) / f"temp_nougat_{equation_count}.txt"
            
            # nougat-latex-ocr 실행 - 올바른 명령
            # GitHub 저장소의 실제 실행 방법
            cmd = f'cd /d "C:\\git\\nougat-latex-ocr" && venv\\Scripts\\python predict.py -i "{str(img_path)}" -o "{str(temp_output.parent)}"'
            
            # Windows에서 shell=True로 실행
            process = subprocess.Popen(
                cmd,
                shell=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding='utf-8',
                errors='ignore',
                bufsize=1
            )
            
            # 진행 상황 표시
            output_lines = []
            while True:
                line = process.stdout.readline()
                if line == '' and process.poll() is not None:
                    break
                if line:
                    line = line.strip()
                    if line:
                        # 주요 단계만 표시
                        if any(keyword in line.lower() for keyword in ['loading', 'model', 'processing', 'generating', 'downloading']):
                            print(f"         → {line}")
                        output_lines.append(line)
            
            # 처리 시간 계산
            elapsed_time = time.time() - start_time
            
            # 결과 추출
            return_code = process.poll()
            
            # 방법 2: 출력 파일에서 결과 읽기
            if return_code == 0:
                # 생성된 파일 찾기
                output_files = list(Path(output_dir).glob(f"temp_nougat_{equation_count}*.txt"))
                if not output_files:
                    output_files = list(Path(output_dir).glob(f"{img_path.stem}*.txt"))
                
                if output_files:
                    # 첫 번째 파일 읽기
                    with open(output_files[0], 'r', encoding='utf-8') as f:
                        latex_code = f.read().strip()
                    
                    # 임시 파일 삭제
                    for f in output_files:
                        try:
                            f.unlink()
                        except:
                            pass
                    
                    if latex_code:
                        print(f"      ✅ 변환 성공: {latex_code[:60]}...")
                        print(f"      ⏱️  소요시간: {elapsed_time:.2f}초")
                        return latex_code
                
                # 출력에서 직접 찾기
                if output_lines:
                    # 마지막 줄이 보통 LaTeX 결과
                    for line in reversed(output_lines):
                        if line and not any(keyword in line.lower() for keyword in ['loading', 'model', 'error', 'warning', 'downloading']):
                            print(f"      ✅ 변환 성공: {line[:60]}...")
                            print(f"      ⏱️  소요시간: {elapsed_time:.2f}초")
                            return line
                
                print(f"      ⚠️  LaTeX 추출 실패")
                print(f"      ⏱️  소요시간: {elapsed_time:.2f}초")
            else:
                print(f"      ❌ 처리 실패 (코드: {return_code})")
                print(f"      ⏱️  소요시간: {elapsed_time:.2f}초")
                
                # 에러 메시지 출력
                if output_lines:
                    print("      📋 에러 내용:")
                    for line in output_lines[-5:]:  # 마지막 5줄만
                        print(f"         {line}")
            
            return None
            
        except Exception as e:
            print(f"      ❌ 오류: {e}")
            return None
    
    def _open_images_folder(self, output_dir):
        """수식 이미지 폴더 자동 열기"""
        try:
            for img_dir in Path(output_dir).glob("**/images"):
                if any(img_dir.glob("*equation*.png")) or any(img_dir.glob("*formula*.png")):
                    print(f"\n📁 수식 이미지 폴더 열기: {img_dir}")
                    if sys.platform == 'win32':
                        os.startfile(str(img_dir))
                    elif sys.platform == 'darwin':
                        subprocess.run(['open', str(img_dir)])
                    else:
                        subprocess.run(['xdg-open', str(img_dir)])
                    break
        except Exception as e:
            print(f"   ⚠️ 폴더 열기 실패: {e}")
            print(f"   💡 수동으로 열어주세요.")
            
            # 중복 코드 제거됨
    
    def update_md_with_latex(self, output_dir, json_data):
        """MD 파일의 수식 이미지를 LaTeX로 교체"""
        try:
            md_files = list(Path(output_dir).glob("*.md"))
            
            for md_file in md_files:
                content = md_file.read_text(encoding='utf-8')
                updated = False
                
                for item in json_data.get('content_list', []):
                    if item.get('type') in ['equation', 'interline_equation'] and 'latex' in item:
                        img_name = Path(item.get('img_path', '')).name
                        latex_code = item['latex']
                        
                        # 이미지 참조를 LaTeX로 교체
                        old_pattern = f"![](images/{img_name})"
                        new_pattern = f"$${latex_code}$$"
                        
                        if old_pattern in content:
                            content = content.replace(old_pattern, new_pattern)
                            updated = True
                
                if updated:
                    md_file.write_text(content, encoding='utf-8')
                    print(f"   ✅ {md_file.name} 수식 LaTeX 변환 완료")
                    
        except Exception as e:
            print(f"❌ MD 파일 업데이트 오류: {e}")
    
    def process_with_mineru(self, pdf_path, output_dir):
        """MinerU로 문서 처리 (수식/문자/표 분리 포함)"""
        print("🚀 MinerU 딥러닝 파이프라인 시작")
        self.show_models_status()
        
        # 수식 이미지는 nougat이 처리할 때 PDF에서 직접 추출
        
        # 표 테두리 제거는 선택사항 - 원본 PDF 사용
        # processed_pdf_path = self.create_table_border_removed_pdf(pdf_path, output_dir)
        # pdf_path = processed_pdf_path
        
        timer = PipelineTimer()
        timer.start_total()
        
        try:
            timer.start_stage("MinerU Deep Learning Processing")
            
            # MinerU 실행 (수식 인식 활성화)
            cmd = [
                "magic-pdf",     # 정확한 명령어: magic-pdf
                "-p", str(pdf_path),
                "-o", str(output_dir),
                "-m", "auto",  # 파싱 방법: auto (기본값)
                # 추가 옵션 제거 - 기본 설정만 사용
            ]
            
            print(f"🔥 실행 명령어: {' '.join(cmd)}")
            print("🚀 MinerU 딥러닝 처리 시작...")
            print("="*60)
            
            # 실시간 출력으로 처리 과정 표시
            process = subprocess.Popen(
                cmd, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.STDOUT, 
                text=True, 
                encoding='utf-8',
                errors='ignore',
                bufsize=1, 
                universal_newlines=True
            )
            
            # 실시간 출력 표시
            current_step = 0
            total_steps = 0
            
            while True:
                output = process.stdout.readline()
                if output == '' and process.poll() is not None:
                    break
                if output:
                    # MinerU 출력을 실시간으로 표시
                    clean_output = output.strip()
                    if clean_output:
                        # 진행률 추출 시도
                        progress_info = ""
                        if "%" in clean_output:
                            try:
                                import re
                                progress_match = re.search(r'(\d+)%', clean_output)
                                if progress_match:
                                    progress = progress_match.group(1)
                                    progress_info = f" [{progress}%]"
                            except:
                                pass
                        
                        # 페이지 진행률 추출
                        if "/" in clean_output and "page" in clean_output.lower():
                            try:
                                page_match = re.search(r'(\d+)/(\d+)', clean_output)
                                if page_match:
                                    current = page_match.group(1)
                                    total = page_match.group(2)
                                    progress_info = f" [{current}/{total}]"
                            except:
                                pass
                        
                        # 주요 단계별로 이모지 추가
                        if "Loading" in clean_output or "loading" in clean_output:
                            print(f"🔄 {clean_output}{progress_info}")
                        elif "Processing" in clean_output or "processing" in clean_output:
                            print(f"⚙️ {clean_output}{progress_info}")
                        elif "Extracting" in clean_output or "extracting" in clean_output:
                            print(f"🔍 {clean_output}{progress_info}")
                        elif "Detecting" in clean_output or "detecting" in clean_output:
                            print(f"🎯 {clean_output}{progress_info}")
                        elif "Formula" in clean_output or "formula" in clean_output:
                            print(f"🧮 {clean_output}{progress_info}")
                        elif "Table" in clean_output or "table" in clean_output:
                            print(f"📊 {clean_output}{progress_info}")
                        elif "OCR" in clean_output or "ocr" in clean_output:
                            print(f"🔤 {clean_output}{progress_info}")
                        elif "Saving" in clean_output or "saving" in clean_output:
                            print(f"💾 {clean_output}{progress_info}")
                        elif "Complete" in clean_output or "complete" in clean_output:
                            print(f"✅ {clean_output}{progress_info}")
                        elif "Error" in clean_output or "error" in clean_output:
                            print(f"❌ {clean_output}")
                        elif "Warning" in clean_output or "warning" in clean_output:
                            print(f"⚠️ {clean_output}")
                        elif "Model" in clean_output or "model" in clean_output:
                            print(f"🤖 {clean_output}{progress_info}")
                        elif "Page" in clean_output or "page" in clean_output:
                            print(f"📄 {clean_output}{progress_info}")
                        elif clean_output.strip():  # 빈 줄이 아닌 경우만
                            print(f"📝 {clean_output}{progress_info}")
                        
                        # 타임스탬프 추가 (선택적)
                        import time
                        if any(keyword in clean_output.lower() for keyword in ['loading', 'processing', 'complete']):
                            timestamp = time.strftime("%H:%M:%S")
                            print(f"   ⏰ {timestamp}")
                        
                        # 버퍼 플러시로 즉시 출력
                        sys.stdout.flush()
            
            # 프로세스 종료 대기
            return_code = process.wait()
            
            print("="*60)
            
            if return_code != 0:
                print(f"❌ MinerU 실행 중 오류 발생 (종료 코드: {return_code})")
                return False
            else:
                print("✅ MinerU 딥러닝 처리 완료!")
                
                # nougat-latex-ocr로 수식 개선 (UniMERNet 대체)
                print("\n🔄 UniMERNet → nougat-latex-ocr 수식 개선...")
                self.enhance_formulas_with_nougat(output_dir, pdf_path)
            
            timer.end_stage()
            
            # 결과 분석
            timer.start_stage("Content Analysis & Separation")
            
            output_path = Path(output_dir)
            pdf_name = Path(pdf_path).stem
            
            # auto 디렉토리 찾기
            auto_dirs = list(output_path.glob(f"**/*auto"))
            if not auto_dirs:
                print("❌ auto 디렉토리를 찾을 수 없습니다.")
                return False
            
            auto_dir = auto_dirs[0]
            print(f"📁 결과 디렉토리: {auto_dir}")
            
            # JSON 파일 분석 (MinerU 0.7.0b1은 파일명이 다름)
            json_files = list(auto_dir.glob("*_content_list.json"))
            if not json_files:
                # content_list.json 없이 다른 이름으로 저장될 수 있음
                json_files = list(auto_dir.glob("*.json"))
                print(f"📋 발견된 JSON 파일들: {[f.name for f in json_files]}")
                
                # content_list 생성을 위해 middle.json과 model.json에서 추출
                content_list_data = self.extract_content_from_mineru_output(auto_dir)
                if content_list_data:
                    # content_list.json 생성
                    content_list_file = auto_dir / f"{pdf_name}_content_list.json"
                    with open(content_list_file, 'w', encoding='utf-8') as f:
                        json.dump(content_list_data, f, ensure_ascii=False, indent=2)
                    print(f"✅ content_list.json 생성: {content_list_file}")
                    json_files = [content_list_file]
                
            if json_files:
                json_file = json_files[0]
                print(f"📄 JSON 파일: {json_file}")
                
                with open(json_file, 'r', encoding='utf-8', errors='ignore') as f:
                    content = json.load(f)
                
                # 컨텐츠 분류
                equations = [item for item in content if item.get('type') == 'equation']
                texts = [item for item in content if item.get('type') == 'text']
                tables = [item for item in content if item.get('type') == 'table']
                
                print(f"🔢 탐지된 수식: {len(equations)}개")
                print(f"📝 탐지된 텍스트: {len(texts)}개")
                print(f"📊 탐지된 표: {len(tables)}개")
                
                # 빈 결과 처리
                if len(content) == 0:
                    print("⚠️ 콘텐츠가 비어있습니다. 다음을 확인하세요:")
                    print("   1. PDF 파일이 손상되지 않았는지")
                    print("   2. PDF에 텍스트/수식이 포함되어 있는지")
                    print("   3. MinerU 모델들이 정상 로딩되었는지")
                    return False
                
                # 수식 LaTeX 변환 확인 (0으로 나누기 방지)
                if len(equations) > 0:
                    latex_count = sum(1 for eq in equations if eq.get('text'))
                    print(f"✅ LaTeX 변환 완료: {latex_count}/{len(equations)}개 ({latex_count/len(equations)*100:.1f}%)")
                else:
                    print("⚠️ 수식이 탐지되지 않았습니다.")
            
            timer.end_stage()
            timer.end_total()
            
            return auto_dir
            
        except Exception as e:
            timer.end_total()
            import traceback
            print(f"❌ 처리 중 오류 발생: {e}")
            print(f"📍 상세 오류 위치:")
            traceback.print_exc()
            return False

class PDFPageSeparator:
    """pdfplumber를 사용한 PDF 페이지 분리 및 테이블 테두리 제거"""
    
    def __init__(self):
        self.pages_data = []
    
    def separate_pages(self, pdf_path, output_dir):
        """PDF를 페이지별로 분리하고 테이블 테두리 제거"""
        try:
            # pdfplumber 대신 PyMuPDF 사용
            import fitz  # PyMuPDF
            from PIL import Image, ImageDraw, ImageFont
            import numpy as np
            import matplotlib.pyplot as plt
            import matplotlib.font_manager as fm
            import io
            
            # 한글 폰트 설정
            plt.rcParams['font.family'] = ['Malgun Gothic', 'DejaVu Sans']
            plt.rcParams['axes.unicode_minus'] = False
        except ImportError:
            print("❌ PyMuPDF와 PIL이 필요합니다: pip install PyMuPDF pillow matplotlib")
            return False
        
        print("📄 PyMuPDF로 PDF 페이지 분리 중...")
        
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
        pages_dir = output_dir / "pages"
        pages_dir.mkdir(parents=True, exist_ok=True)
        
        try:
            # PyMuPDF로 PDF 열기
            pdf_document = fitz.open(pdf_path)
            total_pages = len(pdf_document)
            print(f"📊 총 {total_pages} 페이지 발견")
            
            for page_num in range(total_pages):
                print(f"\n🔄 페이지 {page_num + 1}/{total_pages} 처리 중...")
                
                # 페이지 가져오기
                page = pdf_document[page_num]
                
                # 페이지를 이미지로 변환 (300 DPI)
                mat = fitz.Matrix(300/72, 300/72)
                pix = page.get_pixmap(matrix=mat)
                
                # PIL 이미지로 변환
                img_data = pix.pil_tobytes(format="PNG")
                pil_image = Image.open(io.BytesIO(img_data))
                
                # PyMuPDF는 테이블 감지 기능이 없으므로 테이블 처리 건너뛰기
                # (pdfplumber 대신 PyMuPDF 사용)
                
                # 이미지 저장
                img_path = pages_dir / f"page_{page_num + 1}.png"
                pil_image.save(str(img_path), 'PNG')
                
                # 페이지 정보 저장
                page_info = {
                    "page_num": page_num + 1,
                    "image_path": str(img_path.relative_to(output_dir)),
                    "width": pil_image.width,
                    "height": pil_image.height,
                    "tables_count": 0,  # PyMuPDF는 테이블 감지 미지원
                    "table_borders_removed": False
                }
                self.pages_data.append(page_info)
                
                print(f"✅ 페이지 {page_num + 1} 분리 완료")
            
            # PDF 문서 닫기
            pdf_document.close()
            
            # 페이지 정보 JSON 저장
            pages_json = output_dir / "pages_info.json"
            with open(pages_json, 'w', encoding='utf-8') as f:
                json.dump(self.pages_data, f, ensure_ascii=False, indent=2)
            
            print(f"\n📊 총 {len(self.pages_data)}페이지 분리 완료")
            print(f"✅ 페이지 이미지 생성 완료")
            return True
            
        except Exception as e:
            print(f"❌ PDF 페이지 분리 실패: {e}")
            import traceback
            traceback.print_exc()
            return False


def create_word_conversion_json(auto_dir):
    """Word 변환용 최적화된 JSON 생성"""
    try:
        auto_dir = Path(auto_dir)
        
        # Enhanced JSON 파일 찾기
        enhanced_files = list(auto_dir.glob("*content_list_enhanced.json"))
        if not enhanced_files:
            print("⚠️ Enhanced JSON 파일을 찾을 수 없습니다.")
            return None
        
        # Enhanced JSON 로드
        with open(enhanced_files[0], 'r', encoding='utf-8') as f:
            enhanced_data = json.load(f)
        
        # 페이지 정보 로드
        pages_info_file = auto_dir.parent.parent / "pages_info.json"
        pages_info = []
        if pages_info_file.exists():
            with open(pages_info_file, 'r', encoding='utf-8') as f:
                pages_info = json.load(f)
        
        # Word 변환용 데이터 구조 생성
        word_data = {
            "document_info": {
                "total_pages": len(pages_info),
                "page_size": {"width": 595, "height": 842},  # A4 기본값
                "margins": {"top": 72, "bottom": 72, "left": 72, "right": 72},
                "source_file": str(auto_dir.parent.parent),
                "creation_timestamp": datetime.now().isoformat()
            },
            "pages": [],
            "content_elements": []
        }
        
        # 페이지별 정보 추가
        for page_info in pages_info:
            word_data["pages"].append({
                "page_num": page_info["page_num"],
                "width": page_info.get("width", 595),
                "height": page_info.get("height", 842),
                "image_path": page_info["image_path"],
                "tables_processed": page_info.get("table_borders_removed", False)
            })
        
        # 컨텐츠 요소 변환
        for idx, item in enumerate(enhanced_data):
            element = {
                "id": idx + 1,
                "page_num": item.get("page_idx", 0) + 1,
                "type": item.get("type", "unknown"),
                "content": item.get("text", ""),
                "position": None,
                "word_formatting": {}
            }
            
            # 위치 정보 처리 (픽셀 → pt 변환)
            if item.get("bbox"):
                bbox = item["bbox"]
                # PDF 좌표를 Word pt 좌표로 변환 (72pt = 1inch)
                element["position"] = {
                    "x": bbox[0] * 72 / 96,  # 96 DPI 기준
                    "y": bbox[1] * 72 / 96,
                    "width": item.get("position_info", {}).get("width", 0) * 72 / 96,
                    "height": item.get("position_info", {}).get("height", 0) * 72 / 96
                }
            
            # 타입별 특수 처리
            if element["type"] == "equation" or element["type"] == "interline_equation":
                element["content_type"] = "latex_math"
                element["latex_original"] = item.get("text", "")
                element["word_math_required"] = True
                element["math_type"] = "block" if element["type"] == "interline_equation" else "inline"
                
            elif element["type"] == "table":
                element["content_type"] = "table"
                element["table_html"] = item.get("table_body", "")
                element["image_path"] = item.get("img_path", "")
                element["word_table_required"] = True
                
                # 테이블 내 수식 감지
                if element["table_html"]:
                    # 간단한 LaTeX 패턴 감지
                    import re
                    latex_patterns = [r'\$[^$]+\$', r'\\\([^)]+\\\)', r'\\begin\{[^}]+\}.*?\\end\{[^}]+\}']
                    has_math = any(re.search(pattern, element["table_html"]) for pattern in latex_patterns)
                    element["contains_math"] = has_math
                    
            elif element["type"] == "text":
                element["content_type"] = "text"
                
                # 인라인 수식 감지
                import re
                if re.search(r'\$[^$]+\$', element["content"]):
                    element["contains_inline_math"] = True
                    element["word_math_required"] = True
            
            # 스타일 정보 추가
            if item.get("text_level"):
                element["word_formatting"]["outline_level"] = item["text_level"]
            
            if item.get("score"):
                element["confidence_score"] = item["score"]
            
            word_data["content_elements"].append(element)
        
        # Word 변환용 JSON 저장
        output_file = auto_dir / "word_conversion_data.json"
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(word_data, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Word 변환용 JSON 생성 완료")
        print(f"   📊 총 {len(word_data['content_elements'])}개 요소 변환")
        print(f"   📄 {word_data['document_info']['total_pages']} 페이지 정보 포함")
        
        # 수식 통계
        math_elements = [e for e in word_data["content_elements"] if e.get("word_math_required")]
        table_elements = [e for e in word_data["content_elements"] if e.get("word_table_required")]
        
        if math_elements:
            print(f"   🔢 수식 요소: {len(math_elements)}개 (Word 수식 변환 필요)")
        if table_elements:
            print(f"   📊 표 요소: {len(table_elements)}개 (Word 테이블 생성 필요)")
        
        return output_file
        
    except Exception as e:
        print(f"❌ Word 변환용 JSON 생성 실패: {e}")
        import traceback
        traceback.print_exc()
        return None

class AdvancedHTMLViewer:
    """3등분 위치정보 기반 HTML 뷰어 생성기"""
    
    def __init__(self, auto_dir, pages_data):
        self.auto_dir = Path(auto_dir)
        self.pages_data = pages_data
        self.content_data = []
        self.md_content = ""
    
    def create_viewer(self):
        """3등분 위치정보 기반 HTML 뷰어 생성"""
        print("🎨 3등분 위치정보 기반 HTML 뷰어 생성 중...")
        
        # 데이터 로드
        if not self._load_data():
            return None
        
        # HTML 생성
        html_content = self._create_html()
        
        # HTML 파일 저장
        viewer_path = self.auto_dir / "position_based_viewer.html"
        with open(viewer_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ 3등분 위치정보 기반 HTML 뷰어 생성 완료: {viewer_path}")
        return viewer_path
    
    def _load_data(self):
        """필요한 데이터 로드"""
        try:
            # auto/pages 폴더 생성 및 이미지 복사
            pages_dir = self.auto_dir / "pages"
            pages_dir.mkdir(exist_ok=True)
            
            print("📄 PDF 페이지 이미지 복사 중...")
            for page_data in self.pages_data:
                src_path = self.auto_dir.parent.parent / page_data['image_path']
                if src_path.exists():
                    dst_path = pages_dir / f"page_{page_data['page_num']}.png"
                    import shutil
                    shutil.copy2(src_path, dst_path)
                    print(f"   ✅ 페이지 {page_data['page_num']} 복사 완료")
                    # 페이지 데이터 경로 업데이트 (상대 경로)
                    page_data['image_path'] = f"pages/page_{page_data['page_num']}.png"
            
            # JSON 파일 로드 (다양한 패턴 시도)
            json_file = None
            patterns = [
                "*content_list_regenerated.json",  # 재생성된 파일 우선
                "*content_list_enhanced.json",
                "*_content_list.json",
                "content_list.json"
            ]
            
            for pattern in patterns:
                files = list(self.auto_dir.glob(pattern))
                if files:
                    json_file = files[0]
                    print(f"📊 {json_file.name} 파일 사용")
                    break
            
            if not json_file:
                print("❌ content_list.json 파일을 찾을 수 없습니다.")
                return False
            else:
                # JSON 데이터 로드
                import json
                with open(json_file, 'r', encoding='utf-8') as f:
                    self.content_data = json.load(f)
            
            # MD 파일 로드
            md_files = list(self.auto_dir.glob("*.md"))
            if md_files:
                with open(md_files[0], 'r', encoding='utf-8') as f:
                    self.md_content = f.read()
                print(f"📝 MD 파일 로드: {md_files[0].name}")
            else:
                print("⚠️ MD 파일을 찾을 수 없습니다.")
                self.md_content = "MD 파일을 찾을 수 없습니다."
            
            return True
            
        except Exception as e:
            print(f"❌ 데이터 로드 실패: {e}")
            return False
    
    def _create_html(self):
        """3등분 HTML 뷰어 생성"""
        # JavaScript 데이터 준비
        content_data_js = json.dumps(self.content_data, ensure_ascii=False, indent=2)
        pages_data_js = json.dumps(self.pages_data, ensure_ascii=False, indent=2)
        md_content_js = json.dumps(self.md_content, ensure_ascii=False)
        
        return f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>3등분 위치정보 기반 뷰어 - PDF | MD | JSON</title>
    
    <!-- MathJax 설정 먼저 -->
    <script>
        window.MathJax = {{
            tex: {{ 
                inlineMath: [['$', '$'], ['\\\\(', '\\\\)']], 
                displayMath: [['$$', '$$'], ['\\\\[', '\\\\]']],
                packages: {{'[+]': ['ams', 'base']}}
            }},
            svg: {{ fontCache: 'global' }},
            startup: {{
                ready: () => {{
                    MathJax.startup.defaultReady();
                    console.log('✅ MathJax 로드 완료');
                }}
            }}
        }};
    </script>
    <script src="https://polyfill.io/v3/polyfill.min.js?features=es6"></script>
    <script id="MathJax-script" async src="https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-mml-chtml.js"></script>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        
        body {{ 
            font-family: 'Malgun Gothic', '맑은 고딕', 'Noto Sans KR', 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #f5f5f5;
            overflow: hidden;
            height: 100vh;
        }}
        
        /* 메인 컨테이너 - 세로 배치 */
        .main-container {{
            display: flex;
            flex-direction: column;
            height: 100vh;
            width: 100vw;
        }}
        
        /* 컨트롤 패널 */
        .control-panel {{
            background: #2c3e50;
            color: white;
            padding: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            height: 60px;
            flex-shrink: 0;
        }}
        
        .control-group {{
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        
        .btn {{
            background: #3498db;
            color: white;
            border: none;
            padding: 8px 15px;
            border-radius: 5px;
            cursor: pointer;
            font-size: 14px;
        }}
        
        .btn:hover {{ background: #2980b9; }}
        .btn:disabled {{ background: #7f8c8d; cursor: not-allowed; }}
        
        /* 3등분 패널 컨테이너 - 가로 배치 */
        .panels-container {{
            flex: 1;
            display: flex;
            flex-direction: row;
            height: calc(100vh - 60px);
        }}
        
        /* 개별 패널 스타일 - 가로 33.333% 너비 */
        .panel {{
            flex: 1;
            border-right: 2px solid #34495e;
            display: flex;
            flex-direction: column;
            width: 33.333%;
        }}
        
        .panel:last-child {{ border-right: none; }}
        
        .panel-header {{
            background: #34495e;
            color: white;
            padding: 10px 15px;
            font-weight: bold;
            font-size: 16px;
            flex-shrink: 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        
        .panel-content {{
            flex: 1;
            overflow: auto;
            background: white;
            position: relative;
            cursor: grab;
        }}
        
        .panel-content:active {{
            cursor: grabbing;
        }}
        
        /* PDF 패널 */
        #pdf-panel .panel-content {{
            display: flex;
            justify-content: flex-start;
            align-items: flex-start;
            background: #ecf0f1;
            overflow: auto;
            position: relative;
        }}
        
        .pdf-image {{
            cursor: grab;
            transition: transform 0.1s ease;
            transform-origin: top left;
        }}
        
        .pdf-image:active {{ cursor: grabbing; }}
        
        /* 드래그 가능한 컨테이너 */
        .draggable-container {{
            cursor: grab;
            overflow: auto;
            user-select: none;
        }}
        
        .draggable-container:active {{ cursor: grabbing; }}
        
        /* 패널별 줌 컨트롤 - 패널 상단에 고정 */
        .panel-zoom-controls {{
            position: absolute;
            top: 10px;
            right: 10px;
            display: flex;
            gap: 3px;
            z-index: 1000;
            background: rgba(0,0,0,0.8);
            padding: 5px;
            border-radius: 5px;
            pointer-events: auto;
        }}
        
        .panel-zoom-btn {{
            background: rgba(255,255,255,0.2);
            border: none;
            color: white;
            padding: 4px 8px;
            border-radius: 3px;
            cursor: pointer;
            font-size: 12px;
            font-weight: bold;
            transition: background 0.2s;
        }}
        
        .panel-zoom-btn:hover {{
            background: rgba(255,255,255,0.4);
        }}
        
        .panel-zoom-display {{
            background: rgba(255,255,255,0.1);
            color: white;
            padding: 4px 8px;
            border-radius: 3px;
            font-size: 11px;
            min-width: 45px;
            text-align: center;
            pointer-events: none;
        }}
        
        /* MD 패널 */
        #md-panel .panel-content {{
            padding: 15px;
            background: #2c3e50;
            color: #ecf0f1;
            overflow: auto;
            position: relative;
        }}
        
        .md-content {{
            font-family: 'Malgun Gothic', '맑은 고딕', 'Noto Sans KR', 'Consolas', 'Monaco', monospace;
            font-size: 13px;
            line-height: 1.5;
            white-space: pre-wrap;
            word-wrap: break-word;
            transform-origin: top left;
            transition: transform 0.1s ease;
        }}
        
        /* JSON 렌더링 패널 */
        #json-panel .panel-content {{
            padding: 15px;
            background: white;
            overflow: auto;
            position: relative;
        }}
        
        .json-content {{
            font-family: 'Malgun Gothic', '맑은 고딕', 'Noto Sans KR', 'Segoe UI', sans-serif;
            font-size: 14px;
            line-height: 1.6;
            transform-origin: top left;
            transition: transform 0.1s ease;
        }}
        
        .json-content .equation {{
            margin: 10px 0;
            padding: 10px;
            background: #f8f9fa;
            border-left: 4px solid #3498db;
            border-radius: 3px;
        }}
        
        .json-content .text-block {{
            margin: 8px 0;
            padding: 8px;
            background: #ffffff;
            border-left: 3px solid #27ae60;
        }}
        
        .json-content .table-block {{
            margin: 10px 0;
            padding: 10px;
            background: #fff3cd;
            border-left: 4px solid #ffc107;
            border-radius: 3px;
        }}
        
        /* 줌 컨트롤 */
        .zoom-info {{
            background: rgba(0,0,0,0.7);
            color: white;
            padding: 5px 10px;
            border-radius: 3px;
            font-size: 12px;
        }}
        
        /* 페이지 정보 */
        .page-info {{
            background: #27ae60;
            color: white;
            padding: 5px 15px;
            border-radius: 15px;
            font-weight: bold;
        }}
        
        /* 위치 정보 표시 */
        .position-info {{
            font-size: 11px;
            color: #7f8c8d;
            margin-bottom: 5px;
            font-family: monospace;
        }}
        
        /* 스크롤바 스타일 */
        .panel-content::-webkit-scrollbar {{
            width: 8px;
            height: 8px;
        }}
        
        .panel-content::-webkit-scrollbar-track {{
            background: #f1f1f1;
        }}
        
        .panel-content::-webkit-scrollbar-thumb {{
            background: #888;
            border-radius: 4px;
        }}
        
        .panel-content::-webkit-scrollbar-thumb:hover {{
            background: #555;
        }}
    </style>
</head>
<body>
    <!-- 컨트롤 패널 -->
    <div class="control-panel">
        <div class="control-group">
            <h2>🔍 3등분 위치정보 기반 뷰어</h2>
        </div>
        <div class="control-group">
            <button class="btn" onclick="changePage(-1)" id="prev-btn">◀ 이전</button>
            <span class="page-info" id="page-info">페이지 1 / 1</span>
            <button class="btn" onclick="changePage(1)" id="next-btn">다음 ▶</button>
        </div>
        <div class="control-group">
            <select id="panel-selector" onchange="changeActivePanel()">
                <option value="pdf">📄 PDF 패널</option>
                <option value="md">📝 MD 패널</option>  
                <option value="json">🎨 JSON 패널</option>
            </select>
            <button class="btn" onclick="zoomActivePanel(0.8)">🔍-</button>
            <span class="zoom-info" id="zoom-info">100%</span>
            <button class="btn" onclick="zoomActivePanel(1.5)">🔍+</button>
            <button class="btn" onclick="resetActivePanel()">원본</button>
        </div>
    </div>
    
    <!-- 3등분 패널 컨테이너 - 가로 배치 -->
    <div class="panels-container">
        <!-- PDF 이미지 패널 (왼쪽 33%) -->
        <div class="panel" id="pdf-panel">
            <div class="panel-header">
                📄 PDF 원본 이미지
                <span id="pdf-info">페이지 이미지 로딩 중...</span>
            </div>
            <div class="panel-content">
                <img class="pdf-image" id="pdf-image" src="" alt="PDF Page" />
            </div>
        </div>
        
        <!-- MD 원시 텍스트 패널 (중간 33%) -->
        <div class="panel" id="md-panel">
            <div class="panel-header">
                📝 MD 원시 텍스트 (LaTeX)
                <span id="md-info">원시 마크다운 표시</span>
            </div>
            <div class="panel-content">
                <div class="md-content" id="md-content">MD 내용 로딩 중...</div>
            </div>
        </div>
        
        <!-- JSON 렌더링 패널 (오른쪽 33%) -->
        <div class="panel" id="json-panel">
            <div class="panel-header">
                🎨 JSON 렌더링 (수식 포함)
                <span id="json-info">수식 렌더링 결과</span>
            </div>
            <div class="panel-content">
                <div class="json-content" id="json-content">JSON 렌더링 중...</div>
            </div>
        </div>
    </div>
    
    <script>
        // 전역 변수
        let currentPage = 1;
        let totalPages = 1;
        let zoomLevel = 1.0;
        
        // 패널별 줌 레벨
        let panelZooms = {{
            pdf: 1.0,
            md: 1.0,
            json: 1.0
        }};
        
        // 드래그 상태
        let dragState = {{
            isDragging: false,
            startX: 0,
            startY: 0,
            startScrollLeft: 0,
            startScrollTop: 0,
            currentPanel: null
        }};
        
        // 데이터 (하드코딩)
        const contentData = {content_data_js};
        const pagesData = {pages_data_js};
        const mdContent = {md_content_js};
        
        // MathJax는 이미 head에서 설정됨
        
        // 페이지 변경
        function changePage(direction) {{
            const newPage = currentPage + direction;
            if (newPage >= 1 && newPage <= totalPages) {{
                currentPage = newPage;
                updatePage();
            }}
        }}
        
        // 활성 패널 기반 줌 기능
        let activePanel = 'pdf';  // 기본 활성 패널
        
        function changeActivePanel() {{
            const selector = document.getElementById('panel-selector');
            activePanel = selector.value;
            updateActiveZoomDisplay();
        }}
        
        function zoomActivePanel(factor) {{
            zoomPanel(activePanel, factor);
        }}
        
        function resetActivePanel() {{
            resetPanelZoom(activePanel);
        }}
        
        function updateActiveZoomDisplay() {{
            const zoomPercentage = Math.round(panelZooms[activePanel] * 100);
            document.getElementById('zoom-info').textContent = `${{zoomPercentage}}%`;
        }}
        
        // 페이지 업데이트
        function updatePage() {{
            updatePDFPanel();
            updateMDPanel();
            updateJSONPanel();
            updatePageInfo();
        }}
        
        // PDF 패널 업데이트
        function updatePDFPanel() {{
            if (pagesData[currentPage - 1]) {{
                const pageData = pagesData[currentPage - 1];
                document.getElementById('pdf-image').src = pageData.image_path;
                document.getElementById('pdf-info').textContent = `페이지 ${{currentPage}} (${{pageData.width}}x${{pageData.height}})`;
            }}
        }}
        
        // MD 패널 업데이트 (페이지별로 분할)
        function updateMDPanel() {{
            const mdElement = document.getElementById('md-content');
            
            // 현재 페이지의 컨텐츠 필터링
            const pageItems = contentData.filter(item => 
                (item.page_idx || 0) === currentPage - 1
            );
            
            let pageMDContent = '';
            
            // 페이지별 MD 내용 생성
            pageItems.forEach((item, index) => {{
                const posInfo = item.bbox ? 
                    `<!-- 위치: [x:${{Math.round(item.bbox[0])}}, y:${{Math.round(item.bbox[1])}}, w:${{Math.round(item.bbox[2] - item.bbox[0])}}, h:${{Math.round(item.bbox[3] - item.bbox[1])}}] -->` : 
                    '<!-- 위치정보 없음 -->';
                
                if (item.type === 'equation' || item.type === 'interline_equation') {{
                    pageMDContent += `${{posInfo}}
# 수식 #${{index + 1}}
$$$${{item.text || ''}}$$

`;
                }} else if (item.type === 'text') {{
                    const text = item.text || '';
                    pageMDContent += `${{posInfo}}
${{text}}

`;
                }} else if (item.type === 'table') {{
                    // 테이블은 table_body(HTML) 우선, 없으면 text 사용
                    const tableContent = item.table_body || item.text || '표 내용';
                    pageMDContent += `${{posInfo}}
## 표 #${{index + 1}}
${{tableContent}}

`;
                }}
            }});
            
            if (pageMDContent === '') {{
                pageMDContent = `# 페이지 ${{currentPage}}

이 페이지에는 표시할 MD 컨텐츠가 없습니다.`;
            }} else {{
                pageMDContent = `# 페이지 ${{currentPage}} MD 원시 텍스트

${{pageMDContent}}`;
            }}
            
            mdElement.textContent = pageMDContent;
        }}
        
        // JSON 패널 업데이트 (위치정보 기반)
        function updateJSONPanel() {{
            const jsonElement = document.getElementById('json-content');
            
            // 현재 페이지의 컨텐츠 필터링
            const pageItems = contentData.filter(item => 
                (item.page_idx || 0) === currentPage - 1
            );
            
            let jsonHtml = '';
            
            pageItems.forEach((item, index) => {{
                const posInfo = item.bbox ? 
                    `위치: [x:${{Math.round(item.bbox[0])}}, y:${{Math.round(item.bbox[1])}}, w:${{Math.round(item.bbox[2] - item.bbox[0])}}, h:${{Math.round(item.bbox[3] - item.bbox[1])}}]` : 
                    '위치정보 없음';
                
                if (item.type === 'equation' || item.type === 'interline_equation') {{
                    jsonHtml += `
                        <div class="equation">
                            <div class="position-info">${{posInfo}}</div>
                            <strong>수식 #${{index + 1}}:</strong><br>
                            $${{item.text || ''}}$$
                        </div>
                    `;
                }} else if (item.type === 'text') {{
                    jsonHtml += `
                        <div class="text-block">
                            <div class="position-info">${{posInfo}}</div>
                            <strong>텍스트 #${{index + 1}}:</strong><br>
                            ${{(item.text || '').replace(/</g, '&lt;').replace(/>/g, '&gt;')}}
                        </div>
                    `;
                }} else if (item.type === 'table') {{
                    // 테이블은 table_body(HTML) 우선, 없으면 text 사용
                    const tableContent = item.table_body || item.text || '표 내용';
                    const imgPath = item.img_path ? `<img src="${{item.img_path}}" alt="Table Image" style="max-width:100%; margin:10px 0;"/>` : '';
                    jsonHtml += `
                        <div class="table-block">
                            <div class="position-info">${{posInfo}}</div>
                            <strong>표 #${{index + 1}}:</strong><br>
                            ${{imgPath}}
                            ${{tableContent}}
                        </div>
                    `;
                }}
            }});
            
            if (jsonHtml === '') {{
                jsonHtml = '<p>이 페이지에는 표시할 컨텐츠가 없습니다.</p>';
            }}
            
            jsonElement.innerHTML = jsonHtml;
            
            // MathJax 재렌더링
            if (window.MathJax && window.MathJax.typesetPromise) {{
                MathJax.typesetPromise([jsonElement]).then(() => {{
                    console.log('✅ MathJax 수식 렌더링 완료');
                }}).catch((err) => {{
                    console.error('❌ MathJax 렌더링 실패:', err);
                }});
            }} else {{
                console.warn('⚠️ MathJax가 아직 로드되지 않았습니다.');
                // 1초 후 재시도
                setTimeout(() => {{
                    if (window.MathJax && window.MathJax.typesetPromise) {{
                        MathJax.typesetPromise([jsonElement]);
                    }}
                }}, 1000);
            }}
        }}
        
        // 페이지 정보 업데이트
        function updatePageInfo() {{
            document.getElementById('page-info').textContent = `페이지 ${{currentPage}} / ${{totalPages}}`;
            document.getElementById('prev-btn').disabled = currentPage === 1;
            document.getElementById('next-btn').disabled = currentPage === totalPages;
            
            // MD 패널 헤더 정보 업데이트
            document.getElementById('md-info').textContent = `페이지 ${{currentPage}} 원시 마크다운`;
        }}
        
        // 패널별 줌 기능
        function zoomPanel(panelId, factor) {{
            panelZooms[panelId] = Math.max(0.2, Math.min(panelZooms[panelId] * factor, 5.0));
            applyPanelZoom(panelId);
            updateZoomIndicator(panelId);
        }}
        
        function resetPanelZoom(panelId) {{
            panelZooms[panelId] = 1.0;
            applyPanelZoom(panelId);
            updateZoomIndicator(panelId);
        }}
        
        function applyPanelZoom(panelId) {{
            const element = getPanelElement(panelId);
            if (element) {{
                element.style.transform = `scale(${{panelZooms[panelId]}})`;
            }}
        }}
        
        function getPanelElement(panelId) {{
            switch(panelId) {{
                case 'pdf': return document.getElementById('pdf-image');
                case 'md': return document.getElementById('md-content');
                case 'json': return document.getElementById('json-content');
                default: return null;
            }}
        }}
        
        function updateZoomIndicator(panelId) {{
            // 패널별 줌 인디케이터는 제거됨, 활성 패널의 줌 레벨만 상단에 업데이트
            if (panelId === activePanel) {{
                updateActiveZoomDisplay();
            }}
        }}
        
        // 드래그 기능 설정
        function setupPanelDrag() {{
            const panels = ['pdf-panel', 'md-panel', 'json-panel'];
            
            panels.forEach(panelId => {{
                const panel = document.getElementById(panelId);
                const content = panel.querySelector('.panel-content');
                
                // 마우스 드래그
                content.addEventListener('mousedown', function(e) {{
                    if (e.button === 0) {{ // 좌클릭만
                        dragState.isDragging = true;
                        dragState.startX = e.clientX;
                        dragState.startY = e.clientY;
                        dragState.startScrollLeft = content.scrollLeft;
                        dragState.startScrollTop = content.scrollTop;
                        dragState.currentPanel = panelId;
                        content.style.cursor = 'grabbing';
                        e.preventDefault();
                    }}
                }});
                
                // 마우스 휠 줌
                content.addEventListener('wheel', function(e) {{
                    if (e.ctrlKey) {{
                        e.preventDefault();
                        const panelType = panelId.replace('-panel', '');
                        const factor = e.deltaY > 0 ? 0.8 : 1.25;
                        zoomPanel(panelType, factor);
                    }}
                }});
                
                // 더블클릭으로 줌 리셋
                content.addEventListener('dblclick', function(e) {{
                    const panelType = panelId.replace('-panel', '');
                    resetPanelZoom(panelType);
                }});
            }});
            
            // 전역 마우스 이벤트
            document.addEventListener('mousemove', function(e) {{
                if (dragState.isDragging && dragState.currentPanel) {{
                    const panel = document.getElementById(dragState.currentPanel);
                    const content = panel.querySelector('.panel-content');
                    
                    const deltaX = e.clientX - dragState.startX;
                    const deltaY = e.clientY - dragState.startY;
                    
                    content.scrollLeft = dragState.startScrollLeft - deltaX;
                    content.scrollTop = dragState.startScrollTop - deltaY;
                }}
            }});
            
            document.addEventListener('mouseup', function() {{
                if (dragState.isDragging) {{
                    const panel = document.getElementById(dragState.currentPanel);
                    if (panel) {{
                        const content = panel.querySelector('.panel-content');
                        content.style.cursor = 'grab';
                    }}
                    dragState.isDragging = false;
                    dragState.currentPanel = null;
                }}
            }});
        }}
        
        // 키보드 단축키
        document.addEventListener('keydown', function(e) {{
            switch(e.key) {{
                case 'ArrowLeft':
                    changePage(-1);
                    break;
                case 'ArrowRight':
                    changePage(1);
                    break;
                case '=':
                case '+':
                    if (e.ctrlKey) {{
                        e.preventDefault();
                        zoomIn();
                    }}
                    break;
                case '-':
                    if (e.ctrlKey) {{
                        e.preventDefault();
                        zoomOut();
                    }}
                    break;
                case '0':
                    if (e.ctrlKey) {{
                        e.preventDefault();
                        resetZoom();
                    }}
                    break;
            }}
        }});
        
        // 초기화
        function init() {{
            try {{
                totalPages = pagesData.length;
                console.log(`총 ${{totalPages}} 페이지 로드됨`);
                console.log(`컨텐츠 항목 ${{contentData.length}}개 로드됨`);
                
                // 패널 드래그 기능 설정
                setupPanelDrag();
                
                // 초기 줌 인디케이터 표시
                updateZoomIndicator('pdf');
                updateZoomIndicator('md');
                updateZoomIndicator('json');
                
                // 초기 활성 패널 줌 표시
                updateActiveZoomDisplay();
                
                updatePage();
                console.log('✅ 3등분 위치정보 기반 뷰어 초기화 완료');
                console.log('🎮 사용법:');
                console.log('   📌 Ctrl + 마우스휠: 패널별 줌');
                console.log('   📌 마우스 드래그: 패널별 이동');
                console.log('   📌 더블클릭: 줌 리셋');
                
            }} catch (error) {{
                console.error('❌ 초기화 실패:', error);
            }}
        }}
        
        // 페이지 로드 후 초기화
        document.addEventListener('DOMContentLoaded', init);
    </script>
</body>
</html>"""
        
        with open(viewer_path, 'w', encoding='utf-8') as f:
            f.write(complete_html_content)
        
        # HTML 자동 실행 
        try:
            import webbrowser
            import subprocess
            import os
            
            # Windows에서 기본 브라우저로 HTML 파일 열기
            if os.name == 'nt':  # Windows
                subprocess.run(['start', '', str(viewer_path.absolute())], shell=True, check=False)
                print(f"🌐 HTML 뷰어 자동 실행: {viewer_path.absolute()}")
            else:  # macOS/Linux
                file_url = f"file:///{viewer_path.absolute().as_posix()}"
                webbrowser.open(file_url)
                print(f"🌐 HTML 뷰어 자동 실행: {file_url}")
            
        except Exception as e:
            print(f"⚠️ HTML 자동 실행 실패: {e}")
            # 백업: webbrowser 모듈 사용
            try:
                webbrowser.open(str(viewer_path.absolute()))
                print(f"🌐 HTML 백업 실행 성공")
            except:
                print(f"🌐 수동으로 열기: {viewer_path.absolute()}")
        
        print(f"✅ 고급 3등분 위치정보 기반 뷰어 생성 완료: {viewer_path}")
        
        # 최적화된 3패널 뷰어도 생성
        try:
            # model.json 읽기
            model_json_path = None
            for json_file in auto_dir.glob("**/model.json"):
                model_json_path = json_file
                break
            
            if model_json_path:
                with open(model_json_path, 'r', encoding='utf-8') as f:
                    model_json_data = json.load(f)
                
                # 최적화된 뷰어 생성
                optimized_viewer_path = generate_optimized_3panel_viewer(model_json_data, auto_dir.parent.parent)
                print(f"✅ 최적화된 3패널 뷰어 생성: {optimized_viewer_path}")
                
                # 브라우저에서 열기
                webbrowser.open(f'file:///{Path(optimized_viewer_path).absolute()}')
                
                # 수식 이미지 폴더 열기
                images_dir = None
                for img_dir in auto_dir.glob("**/images"):
                    if img_dir.is_dir():
                        images_dir = img_dir
                        break
                
                if images_dir and images_dir.exists():
                    # 수식 이미지가 있는지 확인
                    equation_images = list(images_dir.glob("equation_*.png"))
                    if equation_images:
                        print(f"\n📁 수식 이미지 폴더 열기: {images_dir}")
                        print(f"   🧮 수식 이미지 {len(equation_images)}개 발견")
                        
                        # Windows에서 폴더 열기
                        if sys.platform == "win32":
                            os.startfile(str(images_dir))
                        elif sys.platform == "darwin":  # macOS
                            subprocess.run(["open", str(images_dir)])
                        else:  # Linux
                            subprocess.run(["xdg-open", str(images_dir)])
                        
                        print("   ✅ 수식 이미지 폴더가 열렸습니다!")
                    else:
                        print(f"   ℹ️ 수식 이미지가 없습니다: {images_dir}")
                        
        except Exception as e:
            print(f"⚠️ 최적화된 뷰어 생성 실패: {e}")
        
        return viewer_path


def main():
    if len(sys.argv) < 2:
        print("사용법: python latex_windows.py <파일 경로> [옵션]")
        print("예시: python latex_windows.py C:/test/document.pdf")
        print("     python latex_windows.py paper.docx")
        print("옵션:")
        print("     --html-only <auto_dir>  : HTML 뷰어만 재생성")
        print("     예: python latex_windows.py --html-only output/20250703_222257_1/1/auto")
        sys.exit(1)
    
    # HTML만 재생성 모드
    if sys.argv[1] == "--html-only" and len(sys.argv) >= 3:
        auto_dir = Path(sys.argv[2])
        if not auto_dir.exists():
            print(f"❌ 디렉토리를 찾을 수 없습니다: {auto_dir}")
            sys.exit(1)
        
        # pages_info.json 찾기
        pages_info_path = auto_dir.parent.parent / "pages_info.json"
        if pages_info_path.exists():
            with open(pages_info_path, 'r', encoding='utf-8') as f:
                pages_data = json.load(f)
            
            viewer = AdvancedHTMLViewer(auto_dir, pages_data)
            viewer_path = viewer.create_viewer()
            
            if viewer_path:
                print(f"\n🎉 HTML 뷰어 재생성 완료!")
                print(f"📁 파일 위치: {viewer_path}")
                print(f"🌐 브라우저에서 열기: file:///{viewer_path.absolute()}")
            sys.exit(0)
        else:
            print("❌ pages_info.json을 찾을 수 없습니다.")
            sys.exit(1)
    
    input_file = sys.argv[1]
    input_path = Path(input_file)
    
    if not input_path.exists():
        print(f"❌ 파일을 찾을 수 없습니다: {input_file}")
        sys.exit(1)
    
    print("="*80)
    print("🔥 MinerU 고급 딥러닝 LaTeX 변환기")
    print("="*80)
    print(f"📄 입력 파일: {input_path}")
    
    # 작업 시간 추적 시작
    timer = PipelineTimer()
    timer.start_total()
    
    # 캐시 정리
    timer.start_stage("캐시 정리")
    cache_manager = CacheManager()
    cache_manager.clear_cache()
    timer.end_stage()
    
    # 출력 디렉토리 생성
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_base = Path("output")
    output_base.mkdir(parents=True, exist_ok=True)
    output_dir = output_base / f"{timestamp}_{input_path.stem}"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"📁 출력 디렉토리: {output_dir}")
    
    # Word → PDF 변환 (필요시)
    if input_path.suffix.lower() in ['.docx', '.doc']:
        if not HAS_PYWIN32:
            print("❌ pywin32가 필요합니다. 'pip install pywin32'로 설치하세요.")
            sys.exit(1)
        
        timer.start_stage("Word → PDF 변환")
        word_converter = WordToPDFConverter()
        try:
            pdf_path = word_converter.convert(input_path, output_dir)
            timer.end_stage()
            print(f"✅ PDF 변환 완료: {pdf_path}")
        except Exception as e:
            timer.end_stage()
            print(f"❌ Word → PDF 변환 실패: {e}")
            sys.exit(1)
        finally:
            pass
    else:
        pdf_path = input_path
    
    # PDF 페이지 분리 및 전처리
    timer.start_stage("PDF 페이지 분리")
    separator = PDFPageSeparator()
    try:
        success = separator.separate_pages(pdf_path, output_dir)
        timer.end_stage()
        
        if not success:
            print("❌ PDF 페이지 분리 실패")
            timer.end_total()
            sys.exit(1)
        
        pages_data = separator.pages_data
        print(f"✅ PDF 페이지 분리 완료: {len(pages_data)}페이지")
        
        # pages_info.json 저장
        pages_info_path = output_dir / "pages_info.json"
        with open(pages_info_path, 'w', encoding='utf-8') as f:
            json.dump(pages_data, f, ensure_ascii=False, indent=2)
        
    except Exception as e:
        timer.end_stage()
        print(f"❌ PDF 페이지 분리 실패: {e}")
        timer.end_total()
        sys.exit(1)
    
    # MinerU 딥러닝 처리
    timer.start_stage("MinerU 딥러닝 처리")
    mineru_processor = MinerUProcessor()
    try:
        auto_dir = mineru_processor.process_with_mineru(pdf_path, output_dir)
        timer.end_stage()
        
        if auto_dir:
            print(f"✅ MinerU 딥러닝 처리 완료")
            print(f"📁 결과 위치: {auto_dir}")
            
            # nougat-latex-ocr로 수식 개선
            timer.start_stage("nougat-latex-ocr 수식 개선")
            try:
                mineru_processor.enhance_formulas_with_nougat(auto_dir)
                timer.end_stage()
            except Exception as e:
                timer.end_stage()
                print(f"⚠️ nougat-latex-ocr 처리 실패: {e}")
                import traceback
                traceback.print_exc()
            
            # Word 변환 JSON 생성
            try:
                word_json_path = create_word_conversion_json(auto_dir)
                print(f"✅ Word 변환 JSON 생성: {word_json_path}")
            except Exception as e:
                print(f"⚠️ Word 변환 JSON 생성 실패: {e}")
            
            # 위치 정보 분석
            timer.start_stage("위치 정보 분석")
            try:
                from position_analyzer import analyze_position_info
                enhanced_json_path = analyze_position_info(auto_dir)
                timer.end_stage()
                print(f"✅ 위치 정보 분석 완료: {enhanced_json_path}")
            except Exception as e:
                timer.end_stage()
                print(f"⚠️ 위치 정보 분석 실패 - 기본 뷰어로 계속 진행: {e}")
            
            # MD 파일 위치 정보 추가
            timer.start_stage("MD 위치 정보 추가")
            try:
                from md_enhancer import enhance_markdown_with_positions
                enhanced_md_path = enhance_markdown_with_positions(auto_dir)
                timer.end_stage()
                print(f"✅ MD 위치 정보 추가 완료: {enhanced_md_path}")
            except Exception as e:
                timer.end_stage()
                print(f"⚠️ MD 위치 정보 추가 실패: {e}")
            
            # 고급 HTML 뷰어 생성
            timer.start_stage("고급 HTML 뷰어 생성")
            try:
                viewer = AdvancedHTMLViewer(auto_dir, pages_data)
                viewer_path = viewer.create_viewer()
                timer.end_stage()
                
                if viewer_path:
                    print(f"\n🎉 처리 완료!")
                    print(f"📁 결과 위치: {auto_dir.absolute()}")
                    print(f"🌐 고급 3등분 위치정보 기반 뷰어: {viewer_path.absolute()}")
                    
                    # 결과 폴더 자동 열기
                    try:
                        import subprocess
                        import os
                        if os.name == 'nt':  # Windows
                            subprocess.run(['explorer', str(auto_dir.absolute())], check=False)
                            print("📂 결과 폴더 자동 열림")
                        elif os.name == 'posix':  # macOS/Linux
                            subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', str(auto_dir.absolute())], check=False)
                    except Exception as e:
                        print(f"⚠️ 폴더 자동 열기 실패: {e}")
                    
                    # HTML 뷰어 자동 실행
                    try:
                        import time
                        print("\n🌐 HTML 뷰어를 자동으로 실행합니다...")
                        time.sleep(1)  # 1초 대기
                        
                        if os.name == 'nt':  # Windows
                            # Chrome으로 우선 시도
                            chrome_paths = [
                                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe")
                            ]
                            
                            chrome_found = False
                            for chrome_path in chrome_paths:
                                if os.path.exists(chrome_path):
                                    subprocess.Popen([chrome_path, str(viewer_path.absolute())])
                                    print("✅ Chrome에서 HTML 뷰어 자동 실행됨")
                                    chrome_found = True
                                    break
                            
                            if not chrome_found:
                                # 기본 브라우저로 실행
                                os.startfile(str(viewer_path.absolute()))
                                print("✅ 기본 브라우저에서 HTML 뷰어 자동 실행됨")
                        else:
                            # macOS/Linux
                            subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', str(viewer_path.absolute())], check=False)
                            print("✅ HTML 뷰어 자동 실행됨")
                        
                        print("\n💡 패널별 확대/축소 기능 추가됨:")
                        print("   - 각 패널 헤더의 +/- 버튼으로 개별 확대/축소")
                        print("   - 상단 🔍+/🔍- 버튼으로 전체 확대/축소")
                        print("   - Ctrl + 마우스휠로 세밀한 줌 조절")
                        print("   - 더블클릭으로 패널별 줌 리셋")
                        
                    except Exception as e:
                        print(f"⚠️ HTML 뷰어 자동 실행 실패: {e}")
                        print(f"   수동으로 파일을 열어주세요: {viewer_path.absolute()}")
                        print(f"   또는 배치파일 실행: {auto_dir.absolute()}/open_viewer.bat")
                    
                    # 전체 작업 시간 통계 표시
                    timer.end_total()
                else:
                    print("⚠️ HTML 뷰어 생성 실패")
                    timer.end_total()
            except Exception as e:
                timer.end_stage()
                print(f"❌ HTML 뷰어 생성 실패: {e}")
                # 전체 작업 시간 통계 표시
                timer.end_total()
        else:
            print("⚠️ HTML 뷰어 생성 실패")
            timer.end_total()
    except Exception as e:
        timer.end_stage()
        print(f"❌ MinerU 처리 실패: {e}")
        timer.end_total()
        sys.exit(1)

if __name__ == "__main__":
    main()
