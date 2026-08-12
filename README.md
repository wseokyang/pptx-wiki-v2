# pptx-wiki

## 여러 PPTX를 한 번에 처리하기

`batch` 명령은 여러 PPTX를 아래 순서로 한 번에 처리합니다.

1. 각 PPTX를 원본 구조 그대로 `parsed`에 추출합니다. `ocr` 설정이 켜져 있으면 네이티브 추출로 읽지 못한 시각 객체만 OCR/VL 모델로 보완합니다.
2. 각 parsed 근거를 LLM으로 선별·재구성해 PPTX당 `semantic.md` 한 개를 만듭니다. 작성 가이드, 템플릿, 무관한 예시와 정확히 중복된 블록은 제외하되 모든 결정은 `decisions.jsonl`에 기록합니다.
3. 모든 `semantic.md`를 다시 읽고, 출처가 한정된 주제·엔터티를 만든 뒤 Quartz가 바로 읽을 수 있는 `quartz/content/`를 게시합니다.

PR 번호는 선택 정보가 아닙니다. 프로그램이 parsed provenance의 본문·표·OCR에서 `PR 번호`, `PR No.`, `의뢰번호` 라벨과 명시적인 `PR-...` 표기를 찾아 원문 값을 고정하며, 한 PPTX 안의 여러 번호도 모두 보존합니다. LLM이 번호를 빠뜨리거나 바꿀 수 없습니다. 번호가 이미지에만 있으면 `config.yml`에서 OCR/VL을 활성화해야 합니다. 번호를 찾지 못한 파일이 하나라도 있으면 LLM 호출 전에 실패합니다.

최초 한 번 Python 부트스트랩을 실행한 뒤 `config.yml`을 편집합니다. 가상환경을
activate할 필요는 없습니다. `run.py`가 프로젝트의 `.venv` Python을 자동으로
사용합니다.

```console
python bootstrap.py
python run.py "D:\reliability\incoming" --batch --recursive --config config.yml --output "D:\reliability\wiki-build"
```

파일을 직접 여러 개 지정해도 됩니다.

```console
python run.py PR-001.pptx PR-002.pptx --batch --config config.yml --output output/collection
```

`config.yml`은 `semantic.enabled: true`, `wiki.enabled: true`, `semantic.coverage_policy: selected`여야 의도한 불필요 내용 제거가 수행됩니다. LLM과 VLM은 서로 다른 `llm_api`, `vlm_api` 설정을 사용합니다.

주요 출력은 다음과 같습니다.

```text
collection/
  README.md                    # 아래 세 최종 산출물의 링크 모음
  collection-manifest.json
  sources/<source-id>/
    source.json
    parsed/                 # 원본 충실 추출물
    semantic/
      semantic.md           # PPTX당 의미 기반 정리 1개
      documents.jsonl
      decisions.jsonl
      manifest.json
  integrated/
    source-map.jsonl        # deck-local citation을 전역 citation으로 매핑
    entities.jsonl
    pages.jsonl
    coverage.jsonl
    manifest.json
  quartz/
    content/                # Quartz 프로젝트의 content/로 사용
      index.md
      topics/
      entities/
      prs/
      sources/
      evidence/
    quartz-manifest.json
    README.md
```

Quartz 프로젝트에 `quartz/content/`를 복사한 뒤 `npx quartz build`로 빌드할 수 있습니다. Quartz의 콘텐츠 루트와 frontmatter/wikilink 규칙은 [Authoring Content](https://quartz.jzhao.xyz/authoring-content), [Frontmatter](https://quartz.jzhao.xyz/plugins/Frontmatter), [Wikilinks](https://quartz.jzhao.xyz/features/wikilinks)를 따릅니다.

한국어 텍스트박스와 표가 많은 PPTX를 구조 우선 방식으로 추출한 뒤,
원본 충실 산출물(`parsed`)과 의미 기반 재정리 산출물(`semantic`)을 각각 남기고,
검증된 semantic 산출물을 Markdown Wiki로 게시하는 파이프라인입니다.

## 설계 원칙

1. 네이티브 텍스트박스와 표는 OCR하지 않고 PPTX OOXML에서 직접 읽습니다.
2. 각 `<a:tbl>` 객체는 간격과 관계없이 독립된 표로 유지합니다.
3. 그림·차트·SmartArt 같은 시각 객체만 개별 ROI로 렌더링합니다.
4. ROI의 원본 픽셀은 이웃 객체와 겹치지 않게 중간 경계에서 자릅니다.
5. 모델에 필요한 여백은 crop 이후 흰색 픽셀로 추가합니다.
6. `parsed`는 LLM 없이 만들며 이후 단계가 수정하지 않는 원본 근거입니다.
7. `semantic`에서만 LLM을 사용해 목적에 맞는 근거를 선택하고 주제별로 재구성합니다.
8. `wiki`는 LLM을 호출하지 않고 검증된 semantic 문서를 Markdown으로 게시합니다.
9. LLM 결과는 기존 block citation과 숫자만 사용할 수 있으며, 검증 실패 시
   원문 block을 그대로 쓰는 fail-closed fallback으로 전환합니다.

따라서 1pt 간격의 네이티브 표 두 개도 한 표로 OCR되지 않습니다. 단, 하나의
비트맵 이미지 자체에 여러 표가 들어 있는 경우에는 PPTX에 내부 경계가 없으므로
OCR/layout backend가 여러 block을 반환해야 합니다. 안전하게 분리할 수 없는
비트맵은 사람이 확인해야 합니다.

## Python 권장 실행법

필요 조건은 64-bit Python 3.10 이상과 데스크톱 Microsoft PowerPoint입니다.
Office COM 자동화 특성상 Windows Service/SYSTEM 계정보다 로그인된 일반 사용자
세션에서 실행하는 것을 권장합니다.

최초 한 번 다음 Python 스크립트를 실행합니다.

```console
python bootstrap.py
```

이 명령은 다음 작업만 수행합니다.

- 프로젝트 전용 `.venv`를 생성하거나 기존의 정상적인 `.venv`를 재사용
- `.venv`에 `pptx-wiki`와 API/Windows 의존성을 editable 모드로 설치
- `config.yml`이 없을 때만 `config.example.yml`을 복사
- 기존 `.venv`나 `config.yml`이 비정상이면 덮어쓰거나 삭제하지 않고 중단

개발·테스트 의존성까지 설치하려면 다음을 사용합니다.

```console
python bootstrap.py --dev
```

직접 환경을 만들고 싶다면 아래와 동일합니다. 활성화는 선택 사항이며, 설치할 때
가상환경의 Python을 명시하면 됩니다.

```console
python -m venv .venv
.venv\Scripts\python.exe -m pip install -e ".[api,windows]"
```

macOS/Linux에서는 마지막 명령의 인터프리터 경로만
`.venv/bin/python`으로 바꿉니다.

OCR 모델은 크고 런타임 의존성이 달라 메인 앱과 섞어 설치하지 않습니다. 이미지
OCR이 필요 없으면 `config.yml`에서 `ocr.enabled: false`, `backend: none`을 사용하고,
OpenAI-compatible VLM이 있으면 `backend: openai_vlm`으로 설정할 수 있습니다. 로컬
GPU OCR worker를 사용할 때만 [workers/README.md](./workers/README.md)의 별도 설치
절차가 필요합니다.

`bootstrap.py`가 [config.example.yml](./config.example.yml)을 복사해 만든 로컬
`config.yml`에서 OCR profile과 semantic 단계용 endpoint를 설정합니다.
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

semantic:
  enabled: true
  goal: "핵심 업무 내용만 보존하고 작성 가이드와 무관한 예시는 제외합니다."
  coverage_policy: selected

wiki:
  enabled: true
```

`api_key_env`를 사용한다면 VS Code 실행 환경이나 운영체제 사용자 환경에 해당
변수를 등록합니다. 또는 `config.yml`의 `api_key`에 직접 넣고 `api_key_env: ""`로
바꿀 수 있지만, 평문 secret이므로 config를 공유하거나 커밋하면 안 됩니다.

이후 VS Code에서 이 폴더를 열고 터미널에서 실행합니다. `run.py`가 프로젝트의
`.venv` Python을 자동으로 찾아 실행하므로 별도로 activate하지 않아도 됩니다.
진행 로그와 최종 결과는 같은 터미널에 계속 출력됩니다.

```console
python run.py "D:\documents\source.pptx"
```

기본적으로 `run.py`와 같은 디렉터리의 고정된 `config.yml`을 읽으며, PPTX 옆이나
현재 작업 디렉터리의 설정 파일을 자동으로 신뢰하지 않습니다. 다른 설정이나 출력
경로를 명시하려면 다음처럼 실행합니다.

```console
python run.py "D:\documents\source.pptx" --config "D:\settings\pptx-wiki.yml" --output "D:\results\source-wiki"
```

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

```console
python bootstrap.py
```

PPTX를 직접 렌더링하려면 LibreOffice와 Poppler의 `pdftocairo`가 필요합니다.
Windows에서 글꼴과 배치 보존이 특히 중요하면 PowerPoint로 미리 PNG/PDF를
내보낸 뒤 `--rendered-slides-dir`을 사용하는 편이 더 정확합니다.

## 단계별 실행

아래 예시는 설치된 가상환경의 Python을 직접 사용하므로 activate가 필요 없습니다.
macOS/Linux에서는 `.venv\Scripts\python.exe` 대신 `.venv/bin/python`을 사용합니다.

### 1. parsed — 원본 충실 추출

```console
.venv\Scripts\python.exe -m pptx_wiki.cli parse input.pptx -o output/my-deck
```

텍스트박스, 네이티브 표, 병합 셀, 그룹 좌표, 발표자 노트가 추출됩니다.
결과는 `output/my-deck/parsed/` 아래에 생성됩니다. 이미지 OCR을 사용하지 않는다면
GPU, LLM, 외부 API가 필요 없습니다.

### 2. semantic — 의미 기반 선택·재정리

```console
.venv\Scripts\python.exe -m pptx_wiki.cli organize output/my-deck/parsed -o output/my-deck/semantic --goal "핵심 업무 내용만 보존하고 작성 가이드와 무관한 예시는 제외합니다." --coverage-policy selected --llm-base-url http://127.0.0.1:8000/v1 --llm-model my-llm
```

이 단계만 OpenAI-compatible 텍스트 LLM이 필요합니다. `selected` 정책에서는 목적과
무관한 block을 제외할 수 있고, 제외된 citation은 `semantic/manifest.json`의
`omitted_citations`에 남습니다. `complete` 정책은 모든 근거를 semantic 문서에
포함합니다. 생성된 문서 본문과 선택 근거는 `semantic/documents.jsonl`에 저장됩니다.

API key가 필요 없는 로컬 서버라면 `OPENAI_API_KEY`를 설정하지 않아도 됩니다.

### 3. wiki — Markdown 게시

```console
.venv\Scripts\python.exe -m pptx_wiki.cli wiki output/my-deck/semantic --parsed output/my-deck/parsed -o output/my-deck/wiki
```

이 단계는 LLM을 호출하지 않습니다. semantic 문서와 parsed provenance의 hash 및
citation을 검증한 다음 `index.md`, 주제별 Markdown, `publish-report.json`을 만듭니다.

세 단계를 한 번에 실행하는 기존 단축 명령도 유지됩니다.

```console
.venv\Scripts\python.exe -m pptx_wiki.cli run input.pptx -o output/my-deck --synthesize --coverage-policy selected --llm-base-url http://127.0.0.1:8000/v1 --llm-model my-llm
```

시각 객체 OCR에 OpenAI-compatible VLM을 사용하려면 parse/run 명령에
`--ocr openai_vlm --vlm-base-url ... --vlm-model ...`을 추가합니다.
`response_format=json_schema`를 지원하지 않는 VLM 서버는 `json_object` 또는
`none`을 사용하십시오.

## 내장 로컬 OCR 모델 사용

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

## 모델별 환경을 직접 점검하기

각 worker는 단건 실행과 장기 JSONL 실행을 모두 제공합니다. 설치된 실제 패키지
목록은 worker 디렉터리에 setup이 남긴 `*.freeze.txt`에서 확인합니다. 다운로드만
다시 검증하려면 해당 venv의 Python으로 실행합니다.

```console
workers\paddleocr_vl_16\.venv\Scripts\python.exe workers\paddleocr_vl_16\download.py --model-dir models\paddleocr_vl_16
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

```console
.venv\Scripts\python.exe -m pptx_wiki.cli run input.pptx -o output/my-deck --rendered-slides-dir rendered-300dpi --ocr openai_vlm --synthesize
```

6–8pt 표가 많다면 300 DPI 또는 약 4000×2250 PNG에서 시작하고, 전체 슬라이드
해상도를 무작정 높이기보다 객체/표 ROI를 유지하는 편이 효과적입니다.

## 출력 구조

```text
output/my-deck/
├── parsed/                   # 1차 산출물: 원본 충실, LLM 불필요
│   ├── manifest.json
│   ├── deck.json             # 전체 native/OCR 중간 표현
│   ├── qa.json               # 추출 경고, OCR 실패, 숫자 충돌
│   ├── source-assets/        # PPTX에 포함된 원본 이미지
│   ├── rendered/             # 슬라이드 렌더 이미지
│   ├── roi/                  # 서로 겹치지 않는 모델 입력 crop
│   ├── ocr-results/          # backend 원본 결과
│   └── corpus/
│       ├── manifest.json
│       ├── provenance.jsonl  # element 단위 근거와 bbox/hash
│       └── slides/slide-0001.md
├── semantic/                 # 2차 산출물: LLM 의미 기반 선택·재정리
│   ├── manifest.json         # 선택/제외 citation, 입력·출력 hash
│   └── documents.jsonl       # 게시 전의 주제별 의미 문서
└── wiki/                     # 후속 게시 산출물: LLM 호출 없음
    ├── index.md
    ├── *.md
    └── publish-report.json
```

semantic과 Wiki의 사실 문장은 `[slide-N#element-id]` 형식으로 원본 block을 가리킵니다.
네이티브 표 두 개는 corpus와 wiki prompt에서 서로 다른 table boundary와 citation을
유지합니다. 숫자·퍼센트·단위를 새로 만들거나 계산한 LLM 출력은 검증에서 거부됩니다.

## 테스트

```console
python bootstrap.py --dev
.venv\Scripts\python.exe -m pytest -q
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
