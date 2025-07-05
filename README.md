# MinerU + Nougat LaTeX Converter

Word 문서의 한글 텍스트와 수식을 보존하면서 LaTeX로 변환하고 다시 Word(OMath)로 재구성하는 파이프라인입니다.

## 📋 프로젝트 개요

이 프로젝트는 두 가지 버전으로 구성되어 있습니다:

1. **MinerU 단독 버전**: MinerU의 자체 수식 감지 기능만 사용
2. **MinerU + Nougat 통합 버전** (최종 버전): MinerU로 레이아웃 분석 + Nougat으로 수식 인식

## 🔧 시스템 요구사항

- Python 3.8-3.10 (3.11 이상에서는 일부 의존성 문제 발생 가능)
- CUDA 지원 GPU (선택사항, CPU로도 실행 가능)
- 최소 8GB RAM
- Windows 10/11 또는 Linux (WSL2 지원)

## 📦 핵심 의존성

### 1. MinerU (magic-pdf) 관련
```
magic-pdf[full]==0.7.0b1
torch==2.0.1
torchvision==0.15.2
paddlepaddle==2.5.1
paddleocr==2.7.0.3
unimernet==0.1.0
```

### 2. Nougat 관련
```
nougat-ocr==0.1.17
transformers==4.36.2
timm==0.5.4
pytorch-lightning==2.1.3
```

### 3. 문서 처리 관련
```
python-docx==1.1.2
PyMuPDF==1.24.5
Pillow==10.3.0
opencv-python==4.9.0.80
numpy==1.24.4  # 1.24.x 버전 필수 (호환성)
```

### 4. 기타 필수 패키지
```
pydantic==2.7.4
pydantic-settings==2.3.4
click==8.1.7
requests>=2.31.0
beautifulsoup4>=4.12.2
lxml>=4.9.3
```

## 🚀 설치 방법 (단계별)

### Step 1: 가상환경 생성
```bash
# Python 3.10 권장
python -m venv mineru_env
source mineru_env/bin/activate  # Linux/Mac
# 또는
mineru_env\Scripts\activate  # Windows
```

### Step 2: 기본 패키지 설치
```bash
# pip 업그레이드
pip install --upgrade pip setuptools wheel

# numpy 먼저 설치 (버전 중요!)
pip install numpy==1.24.4
```

### Step 3: PyTorch 설치
```bash
# CUDA 11.8 버전 (GPU 사용 시)
pip install torch==2.0.1 torchvision==0.15.2 --index-url https://download.pytorch.org/whl/cu118

# CPU 전용
pip install torch==2.0.1 torchvision==0.15.2 --index-url https://download.pytorch.org/whl/cpu
```

### Step 4: MinerU 설치
```bash
# MinerU 설치
pip install magic-pdf[full]==0.7.0b1

# 모델 다운로드
magic-pdf-prepare
```

### Step 5: Nougat 설치
```bash
# Nougat 설치
pip install nougat-ocr==0.1.17

# transformers 버전 고정
pip install transformers==4.36.2
```

### Step 6: 추가 의존성 설치
```bash
# 문서 처리 패키지
pip install python-docx==1.1.2 PyMuPDF==1.24.5

# OCR 관련
pip install paddlepaddle==2.5.1 paddleocr==2.7.0.3

# 기타
pip install opencv-python==4.9.0.80 Pillow==10.3.0
```

## 📁 프로젝트 구조

```
mineru-latex-converter/
├── lw.py                    # MinerU + Nougat 통합 버전 (최종)
├── ln.py                    # MinerU 단독 버전
├── requirements.txt         # 전체 의존성 목록
├── setup_mineru.py         # MinerU 설정 스크립트
├── setup_nougat.py         # Nougat 설정 스크립트
└── word_output_*/          # 출력 디렉토리
    ├── 01_temp.pdf
    ├── 02_images/
    ├── 03_layout_results.json
    ├── 04_formulas/
    ├── 05_combined_results.json
    ├── 06_final_document.docx
    └── 07_final_result.html
```

## 💻 사용 방법

### 기본 사용법
```bash
# MinerU + Nougat 통합 버전 (권장)
python lw.py "입력문서.docx"

# MinerU 단독 버전
python ln.py "입력문서.docx"
```

### 출력 결과
- `06_final_document.docx`: 최종 Word 문서 (OMath 수식 포함)
- `07_final_result.html`: 3패널 뷰어 (원본/LaTeX/렌더링)

## ⚠️ 알려진 문제 및 해결방법

### 1. NumPy 버전 충돌
```bash
# 해결: numpy 1.24.x 버전 고정
pip uninstall numpy
pip install numpy==1.24.4
```

### 2. CUDA 메모리 부족
```python
# lw.py에서 batch_size 조정
batch_size = 1  # 기본값 5에서 1로 변경
```

### 3. MinerU 모델 다운로드 실패
```bash
# 수동 다운로드
wget https://huggingface.co/wanderkid/unimernet_clean/resolve/main/unimernet_base.pth
mkdir -p ~/.cache/magic-pdf/models
mv unimernet_base.pth ~/.cache/magic-pdf/models/
```

### 4. Nougat 초기화 오류
```bash
# transformers 캐시 초기화
rm -rf ~/.cache/huggingface/transformers/
```

## 🔍 코드 차이점

### MinerU 단독 버전 (ln.py)
- MinerU의 내장 수식 감지 기능 사용
- 간단한 수식에 적합
- 처리 속도 빠름

### MinerU + Nougat 통합 버전 (lw.py)
- MinerU로 레이아웃 분석
- Nougat으로 정밀한 수식 인식
- 복잡한 수학 기호 처리 가능
- 더 높은 정확도

## 📊 성능 비교

| 기능 | MinerU 단독 | MinerU + Nougat |
|------|------------|----------------|
| 한글 텍스트 보존 | ✅ | ✅ |
| 기본 수식 인식 | ✅ | ✅ |
| 복잡한 수식 | ⚠️ | ✅ |
| 처리 속도 | 빠름 | 보통 |
| 메모리 사용량 | 낮음 | 높음 |

## 🛠️ 개발 환경 설정

### VSCode 설정
```json
{
    "python.defaultInterpreterPath": "./mineru_env/bin/python",
    "python.linting.enabled": true,
    "python.linting.pylintEnabled": true
}
```

### 디버깅 설정
```json
{
    "version": "0.2.0",
    "configurations": [
        {
            "name": "Python: lw.py",
            "type": "python",
            "request": "launch",
            "program": "${workspaceFolder}/lw.py",
            "args": ["test.docx"],
            "console": "integratedTerminal"
        }
    ]
}
```

## 📝 라이선스

이 프로젝트는 다음 오픈소스 프로젝트들을 활용합니다:
- MinerU (magic-pdf): Apache License 2.0
- Nougat: MIT License
- PyMuPDF: AGPL v3.0

## 🤝 기여 방법

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 문의

문제가 발생하거나 질문이 있으시면 Issues 탭에서 문의해주세요.