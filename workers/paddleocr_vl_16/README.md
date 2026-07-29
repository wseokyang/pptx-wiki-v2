# PaddleOCR-VL 1.6 로컬 worker

이 디렉터리는 `PaddleOCR-VL-1.6`과 `PP-DocLayoutV3`를 Hugging Face에서
고정 리비전으로 내려받아 Windows에서 직접 실행합니다. 외부 OCR API와
`paddleocr` CLI를 호출하지 않으며, 추론 시에는 네트워크 접근을 차단합니다.
다운로드는 필요한 safetensors·설정·토크나이저·Paddle layout 파일만 허용하고
Hugging Face 저장소의 Python 코드나 pickle 계열 가중치는 받지 않습니다.

## 설치

권장 환경은 64-bit Windows, Python 3.11 또는 3.12, NVIDIA RTX 30 시리즈
이상과 VRAM 12GB 이상입니다. CPU도 지원하지만 VLM 추론은 매우 느릴 수
있습니다. PowerShell에서 다음을 실행합니다.

```powershell
Set-ExecutionPolicy -Scope Process Bypass
cd D:\path\to\pptx-wiki\workers\paddleocr_vl_16
.\setup-windows.ps1 -Runtime cu126
```

설치할 수 있는 런타임은 `cpu`, `cu118`, `cu126`, `cu129`입니다. Python
3.12를 지정하려면 `-PythonVersion 3.12`를 추가합니다. 공개 모델이라
Hugging Face 토큰은 필수가 아니지만, 필요하면 실행 전에 `$env:HF_TOKEN`을
설정할 수 있습니다.

설치 결과는 이 디렉터리의 `.venv`, `requirements.freeze.txt`와 저장소
루트의 `models\paddleocr_vl_16`에 보관됩니다. 모델 위치를 바꾸려면 설치 시
`-ModelDir "D:\models\paddleocr_vl_16"`을 지정하고 root `config.yml`의
`models_directory`를 `D:\models`로 맞춥니다. worker 단독 실행 시에도 같은
경로를 `--model-dir`로 전달해야 합니다. 다른 worker와 모델 디렉터리를 공유하지
않습니다.

## 단일 이미지 확인

```powershell
.\.venv\Scripts\python.exe .\worker.py `
  --image "D:\slides\roi.png" `
  --task table `
  --device gpu:0 `
  --output ".\smoke-result.json"
```

`task`는 `document`, `text`, `table`, `chart`, `formula` 중 하나입니다.
표 객체처럼 이미 분리된 ROI에는 `table`을 사용합니다. 이 경우 레이아웃
검출을 다시 적용하지 않아 이웃 표와 재결합되는 일을 피합니다.

## 장기 실행 JSONL 모드

모델을 ROI마다 다시 로드하면 매우 느리므로 실제 파이프라인은 worker를 한
번 시작한 뒤 다음 명령으로 유지해야 합니다.

```powershell
.\.venv\Scripts\python.exe .\worker.py --serve --device gpu:0
```

공통 설정의 `device: auto`는 Paddle의 자동 선택(GPU 0 우선, 없으면 CPU)에
맡기며, `cuda`와 `cuda:N`은 각각 `gpu:0`과 `gpu:N`으로 변환합니다.
`max_new_tokens`는 공식 `PaddleOCRVL.predict()` 인자로 전달됩니다.

stdin 요청 형식:

```json
{"id":"slide-3-image-2","image":"D:\\work\\roi.png","task":"table","language":"ko","context":"분기별 매출표"}
```

공통 orchestrator가 `image_path` 키를 사용해도 동일하게 동작합니다. `image`와
`image_path`를 함께 보낼 때는 값이 같아야 합니다.

종료 요청:

```json
{"type":"shutdown","id":"shutdown"}
```

Paddle 로그는 임의의 stdout 줄에 나타날 수 있습니다. 파이프라인은 오직
`@@PPTX_WIKI@@`로 시작하고 `protocol=pptx-wiki-ocr-worker/1`인 줄만 프로토콜
메시지로 읽어야 합니다. 모델이 준비되면 `type=ready`, 각 요청 뒤에는
`type=result` 메시지가 출력됩니다.

다운로드를 다시 검증하거나 이어받으려면 다음을 실행합니다.

```powershell
.\.venv\Scripts\python.exe .\download.py --model-dir "..\..\models\paddleocr_vl_16"
```

모델 리비전은 모델 디렉터리의 `manifest.json`에 기록되며 worker 시작 시
필수 파일과 함께 다시 검사됩니다.
