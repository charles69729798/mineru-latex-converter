# 상세 설치 가이드 - MinerU + Nougat LaTeX Converter

## 🔍 설치 전 확인사항

### 1. Python 버전 확인
```bash
python --version
# Python 3.8, 3.9, 또는 3.10이어야 함
# 3.11 이상은 호환성 문제 발생 가능
```

### 2. GPU 확인 (선택사항)
```bash
# NVIDIA GPU 확인
nvidia-smi

# CUDA 버전 확인
nvcc --version
```

## 📋 단계별 상세 설치

### Step 1: 가상환경 생성 (필수)
```bash
# Windows
python -m venv mineru_env
mineru_env\Scripts\activate

# Linux/Mac/WSL
python3 -m venv mineru_env
source mineru_env/bin/activate
```

### Step 2: 기본 도구 업그레이드
```bash
# pip 최신 버전으로 업그레이드
python -m pip install --upgrade pip

# 필수 빌드 도구
pip install --upgrade setuptools wheel
```

### Step 3: NumPy 설치 (순서 중요!)
```bash
# NumPy를 반드시 먼저 설치
pip install numpy==1.24.4

# 설치 확인
python -c "import numpy; print(f'NumPy version: {numpy.__version__}')"
```

### Step 4: PyTorch 설치

#### GPU 버전 (CUDA 11.8)
```bash
pip install torch==2.0.1 torchvision==0.15.2 --index-url https://download.pytorch.org/whl/cu118
```

#### CPU 버전
```bash
pip install torch==2.0.1 torchvision==0.15.2 --index-url https://download.pytorch.org/whl/cpu
```

#### 설치 확인
```python
python -c "import torch; print(f'PyTorch: {torch.__version__}, CUDA: {torch.cuda.is_available()}')"
```

### Step 5: MinerU (magic-pdf) 설치

```bash
# MinerU 설치
pip install magic-pdf[full]==0.7.0b1

# 의존성 설치
pip install unimernet==0.1.0
pip install detectron2==0.6 -f https://dl.fbaipublicfiles.com/detectron2/wheels/cu118/torch2.0/index.html
```

#### MinerU 모델 다운로드
```bash
# 자동 다운로드
magic-pdf-prepare

# 다운로드 확인
ls ~/.cache/magic-pdf/
```

### Step 6: PaddleOCR 설치

#### GPU 버전
```bash
pip install paddlepaddle-gpu==2.5.1 -f https://www.paddlepaddle.org.cn/whl/linux/mkl/avx/stable.html
pip install paddleocr==2.7.0.3
```

#### CPU 버전
```bash
pip install paddlepaddle==2.5.1
pip install paddleocr==2.7.0.3
```

### Step 7: Nougat 설치
```bash
# Nougat OCR 설치
pip install nougat-ocr==0.1.17

# 필수 의존성
pip install transformers==4.36.2
pip install timm==0.5.4
pip install pytorch-lightning==2.1.3
```

### Step 8: 문서 처리 라이브러리
```bash
# Word 문서 처리
pip install python-docx==1.1.2

# PDF 처리
pip install PyMuPDF==1.24.5

# 이미지 처리
pip install Pillow==10.3.0
pip install opencv-python==4.9.0.80

# 기타 유틸리티
pip install beautifulsoup4 lxml requests tqdm loguru
```

## 🔧 문제 해결

### 1. ImportError: numpy.core.multiarray failed to import
```bash
# NumPy 재설치
pip uninstall numpy -y
pip install numpy==1.24.4
```

### 2. CUDA out of memory
```python
# lw.py에서 batch_size 수정
batch_size = 1  # 기본값 5에서 1로 변경
```

### 3. MinerU 모델 다운로드 실패
```bash
# 수동 다운로드
mkdir -p ~/.cache/magic-pdf/models

# Layout 모델
wget https://huggingface.co/wanderkid/unimernet_clean/resolve/main/unimernet_base.pth -O ~/.cache/magic-pdf/models/unimernet_base.pth

# OCR 모델
wget https://paddleocr.bj.bcebos.com/PP-OCRv4/chinese/ch_PP-OCRv4_rec_infer.tar
tar -xf ch_PP-OCRv4_rec_infer.tar -C ~/.cache/magic-pdf/models/
```

### 4. Nougat 모델 로드 실패
```bash
# 캐시 삭제 후 재시도
rm -rf ~/.cache/huggingface/
python -c "from nougat import NougatModel; model = NougatModel.from_pretrained('facebook/nougat-small')"
```

### 5. Windows에서 긴 경로 문제
```powershell
# 관리자 권한 PowerShell에서 실행
New-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem" -Name "LongPathsEnabled" -Value 1 -PropertyType DWORD -Force
```

## ✅ 설치 확인 스크립트

```python
# check_installation.py
import sys

def check_import(module_name, package_name=None):
    if package_name is None:
        package_name = module_name
    try:
        __import__(module_name)
        print(f"✅ {package_name} 설치됨")
        return True
    except ImportError:
        print(f"❌ {package_name} 설치 안됨")
        return False

print("=== 설치 확인 ===")
print(f"Python 버전: {sys.version}")

# 핵심 패키지 확인
packages = [
    ("torch", "PyTorch"),
    ("magic_pdf", "MinerU"),
    ("nougat", "Nougat"),
    ("docx", "python-docx"),
    ("fitz", "PyMuPDF"),
    ("paddle", "PaddlePaddle"),
    ("paddleocr", "PaddleOCR"),
    ("cv2", "OpenCV"),
    ("PIL", "Pillow"),
]

all_installed = True
for module, name in packages:
    if not check_import(module, name):
        all_installed = False

if all_installed:
    print("\n✅ 모든 패키지가 정상적으로 설치되었습니다!")
else:
    print("\n❌ 일부 패키지가 누락되었습니다. 위의 가이드를 참고하여 설치하세요.")

# GPU 확인
try:
    import torch
    if torch.cuda.is_available():
        print(f"\n✅ GPU 사용 가능: {torch.cuda.get_device_name(0)}")
    else:
        print("\n⚠️  GPU 사용 불가 (CPU 모드로 실행됩니다)")
except:
    pass
```

## 🚀 빠른 시작

설치가 완료되면:

```bash
# 가상환경 활성화
source mineru_env/bin/activate  # Linux/Mac
# 또는
mineru_env\Scripts\activate  # Windows

# 실행
python lw.py "문서.docx"
```

## 📝 추가 참고사항

1. **메모리 요구사항**: 최소 8GB RAM, 권장 16GB
2. **디스크 공간**: 모델 파일 포함 약 10GB 필요
3. **처리 시간**: 10페이지 문서 기준 약 2-5분 (GPU 사용 시)
4. **지원 형식**: .docx 입력, .docx 출력 (OMath 수식 포함)