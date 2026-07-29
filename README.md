# pptx-wiki

한국어 텍스트박스와 표가 많은 PPTX를 구조 우선 방식으로 추출하고,
필요한 픽셀 영역만 OCR/VLM으로 보완한 뒤 LLM용 Wiki Markdown과
provenance JSONL을 만드는 파이프라인입니다.

## 설계 원칙

1. 네이티브 텍스트박스와 표는 OCR하지 않고 PPTX OOXML에서 직접 읽습니다.
2. 각 `<a:tbl>` 객체는 간격과 관계없이 독립된 표로 유지합니다.
3. 그림·차트·SmartArt 같은 시각 객체만 개별 ROI로 렌더링합니다.
4. ROI의 원본 픽셀은 이웃 객체와 겹치지 않게 중간 경계에서 자릅니다.
5. 모델에 필요한 여백은 crop 이후 흰색 픽셀로 추가합니다.
6. 원본에 충실한 slide corpus와 LLM이 재구성한 wiki를 분리합니다.
7. LLM 결과는 기존 block citation과 숫자만 사용할 수 있으며, 검증 실패 시
   원문 block을 그대로 쓰는 fail-closed fallback으로 전환합니다.

따라서 1pt 간격의 네이티브 표 두 개도 한 표로 OCR되지 않습니다. 단, 하나의
비트맵 이미지 자체에 여러 표가 들어 있는 경우에는 PPTX에 내부 경계가 없으므로
OCR/layout backend가 여러 block을 반환해야 합니다. 안전하게 분리할 수 없는
비트맵은 사람이 확인해야 합니다.

## Windows 권장 실행법

필요 조건은 64-bit Python 3.10 이상과 데스크톱 Microsoft PowerPoint입니다.
Office COM 자동화 특성상 Windows Service/SYSTEM 계정보다 로그인된 일반 사용자
세션에서 실행하는 것을 권장합니다.

최초 한 번 PowerShell에서 설치합니다.

```powershell
Set-ExecutionPolicy -Scope Process Bypass
.\setup-windows.ps1
```

OCR 모델은 메인 앱과 의존성을 섞지 않습니다. 아래 네 환경 중 사용할 것 하나를
별도로 설치하면 고정된 Hugging Face snapshot도 함께 내려받습니다. 기본값은
PaddleOCR-VL 1.6입니다.

```powershell
.\workers\paddleocr_vl_16\setup-windows.ps1 -Runtime cu126
```

CPU라면 `-Runtime cpu`를 사용합니다. 나머지 세 모델의 정확한 설치 명령과 버전은
[workers/README.md](./workers/README.md)에 정리되어 있습니다. 각 환경은 직접
의존성을 `requirements.lock.txt`로 고정하고 설치 완료 후 모든 전이 의존성을
`pip freeze` 파일로 기록합니다.

그다음 `setup-windows.ps1`이 [config.example.yml](./config.example.yml)을 복사해
만든 로컬 `config.yml`에서 OCR profile과 Wiki 합성용 endpoint를 설정합니다.
`config.yml`은 API key 보호를 위해 Git에서 제외됩니다. 로컬 OCR만 사용할 때
`vlm_api`는 비워 두어도 됩니다.

```yaml
ocr:
  enabled: true
  backend: local_model
  local_model:
    model: paddleocr_vl_16
    device: auto
    fallback_to_vlm: false

llm_api:
  base_url: http://127.0.0.1:8000/v1
  model: my-llm
  api_key: ""
  api_key_env: PPTX_WIKI_API_KEY
```

현재 PowerShell 세션에 Wiki LLM key를 넣으려면:

```powershell
$env:PPTX_WIKI_API_KEY = "your-key"
```

Explorer에서 BAT 파일로 실행하려면 환경변수를 Windows 사용자 환경에 영구
등록하거나, `config.yml`의 `api_key`에 직접 넣고 `api_key_env: ""`로 바꿉니다.
후자는 평문 secret이므로 config를 공유하거나 커밋하면 안 됩니다.

이후 실행 시 인수는 PPTX 하나뿐입니다.

```powershell
.\run-windows.ps1 "D:\documents\source.pptx"
```

또는 PPTX를 `run-windows.bat` 위로 drag-and-drop할 수 있습니다. 스크립트는 항상
자신과 같은 디렉터리의 고정된 `config.yml`만 읽으며, PPTX 옆이나 현재 작업
디렉터리의 설정 파일을 자동으로 신뢰하지 않습니다.

Windows 기본 설정은 다음과 같습니다.

- PowerPoint 자체 렌더러, 300 DPI
- 출력 경로 `output/<PPTX 이름>-<원본 hash>/`
- 기존 결과 디렉터리가 비어 있지 않으면 중단
- 외부 링크 이미지·OLE·외부 데이터 관계가 있으면 렌더 전에 중단
- API key 환경변수는 PowerPoint/LibreOffice를 시작하는 동안 프로세스 환경에서 제거
- 원격 endpoint는 명시적으로 `network.allow_remote_endpoints: true`로 허용해야 함
- 원격 평문 HTTP는 추가로 `network.allow_insecure_http: true`가 필요함

`http://127.0.0.1`과 `http://localhost`는 로컬 endpoint로 허용됩니다. 문서 crop을
외부 서버로 보내는 설정은 데이터 반출에 해당하므로 `network` 옵션을 의도적으로
켜야 합니다.

## 설치

```bash
cd /home/wooseok/wsy/pptx-wiki
python3 -m venv .venv
.venv/bin/pip install -e '.[api]'
```

PPTX를 직접 렌더링하려면 LibreOffice와 Poppler의 `pdftocairo`가 필요합니다.
Windows에서 글꼴과 배치 보존이 특히 중요하면 PowerPoint로 미리 PNG/PDF를
내보낸 뒤 `--rendered-slides-dir`을 사용하는 편이 더 정확합니다.

## 1. 네이티브 구조만 추출

```bash
.venv/bin/pptx-wiki run input.pptx -o output/native
```

텍스트박스, 네이티브 표, 병합 셀, 그룹 좌표, 발표자 노트가 추출됩니다.
이 모드는 GPU와 외부 API가 필요 없습니다.

## 2. 보유 중인 OpenAI-compatible VLM과 LLM 사용

```bash
export OPENAI_BASE_URL=http://127.0.0.1:8000/v1
export OPENAI_API_KEY=local-key
export OPENAI_VLM_MODEL=my-vlm
export OPENAI_LLM_MODEL=my-llm

.venv/bin/pptx-wiki run input.pptx -o output/full \
  --ocr openai_vlm \
  --vlm-response-format json_schema \
  --synthesize
```

API key가 필요 없는 로컬 서버라면 `OPENAI_API_KEY`를 설정하지 않아도 됩니다.
`response_format=json_schema`를 지원하지 않는 서버는 `json_object` 또는 `none`을
사용하십시오.

## 3. 내장 로컬 OCR 모델 사용

네 worker는 Hugging Face 모델을 고정 commit으로 먼저 다운로드한 뒤 추론 때는
offline 모드만 사용합니다. `config.yml`에는 repo ID가 아니라 검토된 profile만
선택합니다.

```yaml
ocr:
  enabled: true
  backend: local_model
  local_model:
    # paddleocr_vl_16 | paddleocr_vl_15 | monkeyocr_v2_b | ovisocr2
    model: paddleocr_vl_16
    workers_directory: ./workers
    models_directory: ./models
    device: auto
    dtype: auto
    startup_timeout_seconds: 900
    request_timeout_seconds: 600
    max_new_tokens: 16384
    fallback_to_vlm: false
```

선택한 worker는 한 번만 시작되고 모델도 한 번만 로드됩니다. ROI마다 새 Python을
띄우지 않으므로 수 GB 모델을 반복 로드하지 않습니다. Paddle 1.5/1.6 worker는
`PP-DocLayoutV3 + PaddleOCRVL` 전체 파이프라인을 사용하고, 객체 ROI에서는 layout을
끄고 해당 text/table/chart/formula prompt를 사용합니다.

보유 VLM을 오류 시 fallback으로만 쓰려면 `fallback_to_vlm: true`로 바꾸고
`vlm_api`를 설정합니다. 외부 endpoint라면 crop 전송에 해당하므로 `network` 허용도
필요합니다.

## 4. 모델별 환경을 직접 점검하기

각 worker는 단건 실행과 장기 JSONL 실행을 모두 제공합니다. 설치된 실제 패키지
목록은 worker 디렉터리에 setup이 남긴 `*.freeze.txt`에서 확인합니다. 다운로드만
다시 검증하려면 해당 venv의 Python으로 실행합니다.

```powershell
.\workers\paddleocr_vl_16\.venv\Scripts\python.exe `
  .\workers\paddleocr_vl_16\download.py `
  --model-dir .\models\paddleocr_vl_16
```

worker 공통 결과 형식은 다음과 같습니다.

```json
{
  "text": "전체 텍스트",
  "markdown": "Markdown 또는 HTML 표",
  "html": null,
  "confidence": null,
  "blocks": [
    {
      "kind": "table",
      "text": "셀 텍스트",
      "markdown": "<table>...</table>",
      "html": "<table>...</table>",
      "bbox": [10, 20, 900, 500],
      "confidence": 0.9,
      "order": 0
    }
  ]
}
```

## 미리 렌더링한 슬라이드 사용

슬라이드 수와 이미지 수가 같아야 합니다. 파일명에 슬라이드 번호를 넣는 것을
권장합니다.

```bash
.venv/bin/pptx-wiki run input.pptx -o output/full \
  --rendered-slides-dir ./rendered-300dpi \
  --ocr openai_vlm \
  --synthesize
```

6–8pt 표가 많다면 300 DPI 또는 약 4000×2250 PNG에서 시작하고, 전체 슬라이드
해상도를 무작정 높이기보다 객체/표 ROI를 유지하는 편이 효과적입니다.

## 출력 구조

```text
output/
├── deck.json                 # 전체 native/OCR 중간 표현
├── qa.json                   # 추출 경고, OCR 실패, 숫자 충돌
├── source-assets/            # PPTX에 포함된 원본 이미지
├── rendered/                 # 슬라이드 렌더 이미지
├── roi/                      # 서로 겹치지 않는 모델 입력 crop
├── ocr-results/              # backend 원본 결과
├── corpus/
│   ├── manifest.json
│   ├── provenance.jsonl      # element 단위 근거와 bbox/hash
│   └── slides/slide-0001.md  # 원본 충실 Markdown
└── wiki/
    ├── index.md
    ├── *.md                  # LLM이 주제별로 재구성한 문서
    └── synthesis-report.json
```

Wiki의 사실 문장은 `[slide-N#element-id]` 형식으로 원본 block을 가리킵니다.
네이티브 표 두 개는 corpus와 wiki prompt에서 서로 다른 table boundary와 citation을
유지합니다. 숫자·퍼센트·단위를 새로 만들거나 계산한 LLM 출력은 검증에서 거부됩니다.

## 테스트

```bash
.venv/bin/pip install -e '.[dev,api]'
.venv/bin/pytest -q
```

테스트 fixture에는 한국어 텍스트박스 여러 개, 1pt 간격 표 두 개, 가로·세로
병합 셀, 이미지 표, 그룹 도형이 포함됩니다.

## 현재 의도적으로 남긴 검수 범위

- 한 이미지 안에 붙어 있는 여러 표의 내부 경계
- SmartArt의 시각적 edge/노드 관계
- 픽셀 차트에서 추정한 수치
- 회전·그림자 등으로 실제 core 영역이 겹치는 객체
- LibreOffice와 원본 PowerPoint의 글꼴/줄바꿈 차이

이 항목은 근거를 조용히 합치지 않고 OCR block 또는 QA 경고로 남겨 후속 검수가
가능하도록 하는 것이 원칙입니다.
