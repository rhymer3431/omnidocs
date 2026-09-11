# OmniDocs Architecture

## 목표

1. HWP/HWPX/DOCX를 OMDX로 정규화해 같은 편집 UX와 같은 조판 결과로 다룬다.
2. 문서 편집/조판 엔진을 새로 구현하지 않는다.
3. 포맷별 코드를 어댑터에 가두어 나중에 더 강한 변환 엔진으로 교체할 수 있게 한다.
4. AI 기능은 현재 범위에서 제외한다.

## 0.2 의사결정 — OMDX canonical layer

초기 계획의 LibreOfficeKit은 장기적으로 강한 DOCX fidelity를 제공하지만 현재 개발 머신에는 LibreOffice SDK/런타임이 없으며 macOS 앱 패키징까지 포함하면 초기 통합 비용이 큽니다. 따라서 0.1에서는 rHWP를 공통 편집 코어로 선정했습니다.

`@rhwp/core` 0.8.6에는 HWP/HWPX parser, pagination, SVG/HTML rendering, text/paragraph/table/image editing, HWP/HWPX export API가 이미 존재합니다. `@rhwp/editor`는 그 코어를 사용하는 완성된 Studio를 iframe으로 임베드하고 command/HwpCtrl API를 제공합니다.

이 선택으로 OmniDocs가 직접 구현해야 할 부분은 **OMDX gateway, 포맷 adapter와 제품 shell**로 줄어듭니다.

0.2부터 외부 형식을 편집기의 내부 상태로 직접 유지하지 않습니다. 모든 입력은 다음 OMDX v1 package로
정규화합니다.

```text
document.omdx
├── mimetype
├── manifest.json
├── content/document.hwpx   <- 유일한 편집/layout authority
└── source/original.*       <- immutable round-trip source
```

편집 엔진은 `content/document.hwpx`만 로드합니다. 따라서 DOCX에서 왔는지 HWP에서 왔는지에 따라
다른 layout engine이 선택되는 일이 없습니다.

## 모듈

### `src/editor`

- `contracts.ts`: UI가 의존하는 최소 편집 엔진 계약
- `RhwpStudioEngine.ts`: rHWP Studio 통합 및 포맷 gateway 호출

UI는 rHWP의 구체 API를 직접 호출하지 않습니다. 따라서 후에 `LibreOfficeEngine` 또는 native desktop engine을 추가해도 파일 UI는 유지할 수 있습니다.

### `src/formats`

- `canonical.ts`: 포맷 어댑터가 공유하는 얇은 canonical IR envelope와 feature inventory 계약
- `sourceAdapters.ts`: HWP/HWPX/DOCX → canonical HWPX importer registry
- `canonicalExport.ts`: canonical HWPX → HWP/HWPX/DOCX exporter registry
- `hwpxDocxWriter.ts`: canonical HWPX → WordprocessingML 직접 serializer (HTML 미사용)
- `rhwpRuntime.ts`: WASM 초기화와 공통 document 생성
- `omdx.ts`: OMDX ZIP/manifest/checksum codec
- `omdxGateway.ts`: HWP/HWPX/DOCX/OMDX → canonical OMDX gateway
- `docxLayout.ts`: WordprocessingML page/margin/header/footer sidecar parser
- `docxFeatures.ts`: OOXML feature inventory와 native/normalized/approximated/source-only 분류
- `docxAdapter.ts`: DOCX ↔ canonical HWPX bridge

현재 DOCX import는 `DOCX -> Mammoth semantic model + OOXML layout metadata -> HwpDocument -> OMDX` 경로입니다.

Mammoth transform에서 이미 읽힌 explicit alignment/indent/font/fontSize/highlight를 버리지 않고 HTML style로
보강합니다. underline도 custom style map으로 보존합니다. 페이지 크기/여백과 기본 머리말/꼬리말은
`word/document.xml`과 relationship part에서 직접 추출합니다.

현재 DOCX export는 `canonical HWPX XML -> WordprocessingML parts -> DOCX ZIP` 경로입니다. 문단/런의
직접 서식, 번호 목록, 표와 병합 셀, 이미지 resource, 기본 DrawingML 도형, 각주/미주 part, section page
geometry, column, header/footer를 OOXML로 직접 serialize합니다. floating 그림은 `wp:anchor`, floating 표는
`w:tblpPr`로 위치/정렬/offset/wrap 정보를 보존하고, HWP equation script의 분수·첨자·제곱근·합/곱/적분은
OMML로 구조화합니다. 페이지 렌더 HTML은 DOCX 저장 경로에서 사용하지 않습니다.

### `src/ui`

파일 열기, 저장 형식 선택, 상태 및 호환성 경고만 OmniDocs가 담당합니다. 실제 문서 toolbar/menu/page canvas는 rHWP Studio를 재사용합니다.

## OMDX와 Document IR의 관계

OMDX는 Paragraph/TextRun을 새로 정의하는 두 번째 layout model이 아닙니다. 자체 Paragraph/TextRun/Table/Layout
모델을 만들면 사실상 워드프로세서의 절반을 다시 구현하게 되므로, OMDX v1의 canonical payload는 rHWP가
이미 완전하게 편집할 수 있는 HWPX입니다.

원본 포맷의 unknown extension/metadata는 OMDX `source/original.*`에 그대로 보존하고 변환 손실 보고서는
manifest `compatibility`에 기록합니다.

`compatibility.features`에는 입력 문서에서 탐지된 주요 기능과 canonical 처리 수준을 기록합니다. 이 정보는
편집 엔진을 두 번째로 구현하기 위한 Document AST가 아니라, adapter fidelity를 계측하고 이후 OOXML/HWP
변환기 개선 우선순위를 정하기 위한 sidecar IR입니다. 실제 layout authority는 계속 canonical HWPX/rHWP입니다.

### 무손실 no-op 저장

OMDX manifest는 원본 checksum과 `canonicalSha256AtImport`를 기록합니다. 저장 직전 현재 canonical HWPX의
SHA-256이 import 시 값과 동일하고 대상 포맷도 원본과 같으면 serializer를 거치지 않고 original bytes를
그대로 반환합니다. 이 경로는 OmniDocs가 아직 해석하지 못하는 DOCX/HWP feature도 보존합니다.

## 단계별 로드맵

### M0 — 완료

- 실행 가능한 web shell
- 동일 rHWP Studio 편집면
- HWP/HWPX native round-trip
- DOCX basic import/export
- adapter boundary

### M1 — OMDX canonicalization — 현재 완료

- `.omdx` ZIP package v1
- canonical HWPX + immutable source payload
- SHA-256 integrity verification
- HWP/HWPX → HWPX normalization + rHWP content-loss report
- DOCX → canonical HWPX normalization
- DOCX page size/margins/default header/footer import
- DOCX explicit alignment/indent/font/font size/underline/highlight preservation
- no-op source-format byte-exact save
- OMDX/HWP/HWPX/DOCX Save As

### M2 — 호환성 테스트 확대

- `tests/fixtures/{hwp,hwpx,docx}` corpus
- open/save/reopen 자동 검증
- page count, text, table count, render snapshot 비교
- 실제 학교/공공기관 문서 fixture 추가

현재 첫 단계로 `npm run test:compat`에 HWP/HWPX export와 DOCX 기본 문단/목록/표 round-trip
smoke test를 추가했다. 이후 실제 fixture corpus를 이 경로에 누적한다.

### Self-hosted Studio

`@rhwp/editor` npm 패키지는 iframe SDK이며 Studio 정적 파일 자체를 포함하지 않는다. 기본 URL을
그대로 사용하면 외부 GitHub Pages가 런타임 의존성이 된다.

OmniDocs는 rHWP 0.8.6의 공개 Studio build를 `scripts/vendor-rhwp-studio.mjs`로
`public/rhwp/`에 vendoring하고 기본 `studioUrl`을 같은 origin의 `/rhwp/index.html`로 설정한다. 이 스크립트는
Studio HTML/JS/CSS/WASM/폰트/이미지 자원을 수집하고 JS bundle에 고정한 rHWP 버전이 포함돼 있는지
확인한다.

이 단계는 편집 엔진을 포크하는 것이 아니라 upstream의 검증된 편집 UI와 command 구현을 그대로
재사용하기 위한 배포 방식이다.

### M3 — DOCX fidelity 향상

- HTML 기반 DOCX export 제거 및 canonical HWPX → OOXML 직접 writer 도입 — 완료
- direct writer의 문단/런/표/병합 셀/이미지/section/header-footer 기본 매핑 — 완료
- floating image `wp:anchor` / floating table `w:tblpPr` 기본 위치 매핑 — 완료
- HWP equation → OMML 기본 변환(분수/상·하첨자/제곱근/합·곱·적분) — 완료
- 사각형/타원/선 → DrawingML word-processing shape 기본 매핑 — 완료
- 각주/미주 → `footnotes.xml` / `endnotes.xml` 및 본문 reference 매핑 — 완료
- 복잡 수식 문법, 자유곡선/복합 도형, field/차트/OLE의 정밀 매핑은 fixture 기반으로 확대
- 문서 호환 요구가 자체 writer의 범위를 넘는 경우 LibreOfficeKit/Collabora converter를 adapter 내부 대체 구현으로 도입

### M4 — Desktop packaging

웹 MVP 안정화 후 Tauri 또는 Electron shell을 추가합니다. 브라우저 단계에서 먼저 포맷/편집 동작을 안정화해 desktop packaging 문제와 문서 엔진 문제를 분리합니다.

### M5 — 제품 UI

rHWP Studio chrome을 줄이고 OmniDocs 자체 ribbon을 붙입니다. `studio.commands.list()/execute()`를 사용하여 기존 command 구현을 그대로 재사용합니다.

### M6 — AI

현재 범위 밖입니다. 추후 command API 위에 edit transaction layer를 추가합니다.

## 라이선스

- rHWP: MIT
- React/Vite: MIT
- mammoth: BSD-2-Clause
- JSZip: MIT

배포 시 `THIRD_PARTY_NOTICES.md`를 유지하고 실제 vendored Studio의 라이선스와 폰트 라이선스를 별도로 확인합니다.
