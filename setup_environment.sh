#!/bin/bash
# MinerU + Nougat LaTeX Converter 환경 설정 스크립트
# 사용법: ./setup_environment.sh [gpu|cpu]

set -e

# 색상 정의
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# GPU/CPU 선택
if [ "$1" == "gpu" ]; then
    USE_GPU=true
    echo -e "${GREEN}GPU 버전으로 설치합니다.${NC}"
elif [ "$1" == "cpu" ]; then
    USE_GPU=false
    echo -e "${YELLOW}CPU 버전으로 설치합니다.${NC}"
else
    echo -e "${YELLOW}사용법: ./setup_environment.sh [gpu|cpu]${NC}"
    echo "기본값: CPU 버전"
    USE_GPU=false
fi

echo -e "${GREEN}=== MinerU + Nougat LaTeX Converter 설치 시작 ===${NC}"

# Python 버전 확인
python_version=$(python3 --version 2>&1 | awk '{print $2}')
echo "Python 버전: $python_version"

# 가상환경 생성
echo -e "${GREEN}1. 가상환경 생성 중...${NC}"
if [ ! -d "mineru_env" ]; then
    python3 -m venv mineru_env
fi

# 가상환경 활성화
source mineru_env/bin/activate

# pip 업그레이드
echo -e "${GREEN}2. pip 업그레이드 중...${NC}"
pip install --upgrade pip setuptools wheel

# NumPy 먼저 설치 (중요!)
echo -e "${GREEN}3. NumPy 설치 중...${NC}"
pip install numpy==1.24.4

# PyTorch 설치
echo -e "${GREEN}4. PyTorch 설치 중...${NC}"
if [ "$USE_GPU" = true ]; then
    pip install torch==2.0.1 torchvision==0.15.2 --index-url https://download.pytorch.org/whl/cu118
else
    pip install torch==2.0.1 torchvision==0.15.2 --index-url https://download.pytorch.org/whl/cpu
fi

# MinerU 설치
echo -e "${GREEN}5. MinerU (magic-pdf) 설치 중...${NC}"
pip install magic-pdf[full]==0.7.0b1

# PaddlePaddle 설치
echo -e "${GREEN}6. PaddleOCR 설치 중...${NC}"
if [ "$USE_GPU" = true ]; then
    pip install paddlepaddle-gpu==2.5.1
else
    pip install paddlepaddle==2.5.1
fi
pip install paddleocr==2.7.0.3

# Nougat 설치
echo -e "${GREEN}7. Nougat OCR 설치 중...${NC}"
pip install nougat-ocr==0.1.17
pip install transformers==4.36.2

# 추가 의존성 설치
echo -e "${GREEN}8. 추가 패키지 설치 중...${NC}"
pip install python-docx==1.1.2 PyMuPDF==1.24.5
pip install opencv-python==4.9.0.80 Pillow==10.3.0
pip install beautifulsoup4 lxml requests tqdm loguru

# MinerU 모델 다운로드
echo -e "${GREEN}9. MinerU 모델 다운로드 중...${NC}"
magic-pdf-prepare

# 설치 확인
echo -e "${GREEN}10. 설치 확인 중...${NC}"
python -c "import torch; print(f'PyTorch 버전: {torch.__version__}')"
python -c "import magic_pdf; print('MinerU 설치 완료')"
python -c "import nougat; print('Nougat 설치 완료')"
python -c "import docx; print('python-docx 설치 완료')"

echo -e "${GREEN}=== 설치 완료! ===${NC}"
echo -e "${YELLOW}사용 방법:${NC}"
echo "1. 가상환경 활성화: source mineru_env/bin/activate"
echo "2. 실행: python lw.py '입력문서.docx'"