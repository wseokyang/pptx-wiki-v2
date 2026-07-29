# Local OCR worker environments

네 모델은 서로 충돌할 수 있는 런타임을 사용하므로 각각 독립된 `.venv`에
설치됩니다. 메인 `pptx-wiki` 환경에는 PaddlePaddle, PyTorch, Transformers 또는
vLLM을 설치하지 않습니다. 선택된 worker 프로세스는 모델을 한 번만 메모리에
올린 뒤 여러 ROI 요청을 JSONL로 순차 처리합니다.

## 2026-07-29 고정 버전

| profile | Python | 고정 모델 revision | 실행환경의 핵심 버전 | Windows 실행 방식 |
|---|---:|---|---|---|
| `paddleocr_vl_16` | 3.11 기본 | `66317acc4c9fc17bd154591ce650735cd2855f3e` | PaddleOCR 3.7.0, PaddleX 3.7.2, PaddlePaddle 3.3.1 | 네이티브 전체 document parser |
| `paddleocr_vl_15` | 3.11 기본 | `426bf5b6c89670e370e71ce0c51cf2bb458b7db9` | PaddleOCR 3.7.0, PaddleX 3.7.2, PaddlePaddle 3.3.1 | 네이티브 전체 document parser |
| `monkeyocr_v2_b` | 3.10 | `de7a993bd0f39a97b122dac767e82ae04935bce4` | PyTorch 2.6.0, Transformers 4.57.6 | 네이티브 Transformers ROI 인식(실험적) |
| `ovisocr2` | 3.12 | `65c619d374b55d4152e85150fc1b003700bc1f0c` | PyTorch 2.13.0, Transformers 5.14.1 | 네이티브 Transformers page parser |

Paddle 두 환경은 공통 layout 모델
`PaddlePaddle/PP-DocLayoutV3@7b48a7566925fa464281f930c58eee04fe2c862a`
도 함께 받습니다. Monkey만 Hugging Face repository의 고정된 Python 구현을
실행해야 하며, worker는 허용된 세 소스 파일의 SHA-256을 다운로드와 실행 시점에
모두 검사합니다. 다른 세 profile은 remote code를 실행하지 않습니다.

## 설치

PowerShell 실행 정책을 현재 프로세스에서 허용한 뒤 하나만 선택해 설치합니다.

```powershell
Set-ExecutionPolicy -Scope Process Bypass

# 권장 기본값: CUDA 12.6 PaddleOCR-VL 1.6
.\workers\paddleocr_vl_16\setup-windows.ps1 -Runtime cu126

# PaddleOCR-VL 1.5 비교 실행
.\workers\paddleocr_vl_15\setup-windows.ps1 -Runtime cu126

# MonkeyOCRv2-B (Windows native Transformers)
.\workers\monkeyocr_v2_b\setup-windows.ps1 -TorchBackend cu126

# OvisOCR2 (Windows native Transformers)
.\workers\ovisocr2\setup-windows.ps1 -TorchBackend cu130
```

Paddle runtime 선택지는 `cpu`, `cu118`, `cu126`, `cu129`입니다. Monkey/Ovis의
정확한 PyTorch index 선택지는 각 `setup-windows.ps1 -?`로 확인하십시오. GPU
driver와 wheel이 맞지 않으면 CPU를 선택할 수 있지만, 생성형 OCR은 CPU에서 매우
느릴 수 있습니다.

각 setup은 다음 작업을 수행합니다.

1. 해당 worker 아래에 전용 `.venv` 생성
2. 공식 CPU/CUDA wheel index에서 프레임워크 설치
3. `requirements.lock.txt`의 정확한 직접 의존성 설치
4. Hugging Face에서 고정 commit snapshot 다운로드 및 필수 파일 검증
5. 실제 전이 의존성 전체를 profile별 `*.freeze.txt`로 기록

설치 후 현재 환경을 직접 확인할 때는 profile의 Python을 정확히 지정합니다.

```powershell
& .\workers\paddleocr_vl_16\.venv\Scripts\python.exe -m pip list
& .\workers\paddleocr_vl_16\.venv\Scripts\python.exe -m pip check
```

다른 모델은 위 경로의 profile 이름만 바꾸면 됩니다. `requirements.lock.txt`는
의도한 직접 의존성이고, setup이 남긴 `*.freeze.txt`는 그 시점에 실제 설치된 모든
전이 의존성까지 포함하므로 재현 문제를 조사할 때는 둘을 함께 보십시오.

모델은 기본적으로 repository root의 `models/<profile>`에 저장됩니다. 공개
checkpoint이므로 token은 보통 필요하지 않습니다. 필요한 환경에서는 설치 전에만
`$env:HF_TOKEN`을 설정하십시오. 추론 worker에는 이 token이 전달되지 않습니다.

## 다운로드만 다시 실행

```powershell
.\workers\ovisocr2\.venv\Scripts\python.exe `
  .\workers\ovisocr2\download.py `
  --model-dir .\models\ovisocr2
```

네 downloader 모두 repo ID와 40자리 revision을 코드에 고정합니다. YAML에서 임의
repo나 `trust_remote_code` 값을 입력할 수 없습니다. 다운로드 후 추론은
`HF_HUB_OFFLINE=1`, `TRANSFORMERS_OFFLINE=1`, local model path만 사용합니다.

## profile 선택

root [config.yml](../config.yml)에서 이름만 바꿉니다.

```yaml
ocr:
  enabled: true
  backend: local_model
  local_model:
    model: ovisocr2
    workers_directory: ./workers
    models_directory: ./models
    device: auto
    dtype: auto
```

## 선택 기준

- 실제 작업 기본값은 `paddleocr_vl_16`입니다. 네이티브 표는 OOXML에서 직접
  추출하고, 이미지 문서에는 layout detector와 VLM을 함께 적용할 수 있습니다.
- `paddleocr_vl_15`는 동일한 pipeline 조건에서 1.6과 A/B 비교할 때 사용합니다.
- `monkeyocr_v2_b`의 공식 전체 parser와 고속 serving은 vLLM 중심이며 Windows
  네이티브 vLLM은 지원되지 않습니다. 포함된 worker는 이미 분리된 PPT 객체 ROI를
  직접 읽는 Transformers 경로라서 실험적으로 표시합니다.
- `ovisocr2`의 공식 최고속 경로도 `vllm==0.22.1`이지만 Linux/WSL2가 필요합니다.
  포함된 기본 worker는 Windows에서 직접 실행 가능한 Transformers 경로입니다.

공식 참고자료:

- [PaddleOCR-VL pipeline](https://www.paddleocr.ai/latest/en/version3.x/pipeline_usage/PaddleOCR-VL.html)
- [PaddleOCR-VL 1.6](https://huggingface.co/PaddlePaddle/PaddleOCR-VL-1.6)
- [PaddleOCR-VL 1.5](https://huggingface.co/PaddlePaddle/PaddleOCR-VL-1.5)
- [MonkeyOCRv2](https://github.com/Yuliang-Liu/MonkeyOCRv2)
- [OvisOCR2](https://huggingface.co/ATH-MaaS/OvisOCR2)
