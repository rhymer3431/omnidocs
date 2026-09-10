# OMDX v1

OMDX는 OmniDocs의 편집용 canonical 문서 패키지다. HWP/HWPX/DOCX를 편집기별로 따로 유지하지 않고,
모든 입력을 OMDX로 정규화한 뒤 같은 rHWP DocumentCore/Studio에서 편집한다.

## 컨테이너

OMDX는 ZIP 컨테이너다. MIME type은 `application/vnd.omnidocs.document+zip`이다.

```text
document.omdx
├── mimetype
├── manifest.json
├── content/
│   └── document.hwpx
└── source/
    └── original.docx | original.hwp | original.hwpx
```

`content/document.hwpx`가 편집과 화면 조판의 유일한 canonical payload다. 따라서 원본이 DOCX든 HWP든
OmniDocs 화면에서는 같은 pagination/layout engine을 사용한다.

`source/original.*`은 round-trip fidelity용 immutable 원본이다. 사용자가 문서를 수정하지 않았다면
원래 형식으로 저장할 때 이 바이트를 그대로 반환할 수 있다. 수정된 경우에는 canonical HWPX에서
대상 포맷으로 새로 serialize한다.

## manifest.json

핵심 필드는 다음과 같다.

```json
{
  "schema": "org.omnidocs.omdx",
  "version": 1,
  "createdAt": "2026-09-10T00:00:00.000Z",
  "modifiedAt": "2026-09-10T00:00:00.000Z",
  "canonical": {
    "path": "content/document.hwpx",
    "format": "hwpx",
    "mediaType": "application/vnd.hancom.hwpx+zip",
    "sha256": "...",
    "engine": "rhwp",
    "engineVersion": "0.8.6",
    "pageCount": 1
  },
  "source": {
    "path": "source/original.docx",
    "format": "docx",
    "fileName": "original.docx",
    "sha256": "...",
    "canonicalSha256AtImport": "..."
  },
  "compatibility": {
    "importWarnings": [],
    "contentLoss": {},
    "layout": {}
  }
}
```

`canonicalSha256AtImport`는 원본을 정규화했을 당시 canonical HWPX의 SHA-256이다. 저장 시 현재
canonical SHA-256과 이 값이 같고 대상 형식도 원본 형식과 같으면 원본 파일을 그대로 반환한다.

## Import

```text
HWP  ── rHWP parse/export ──┐
HWPX ─ rHWP normalize ──────┼──> canonical HWPX ──> OMDX ──> Editor
DOCX ─ Mammoth + OOXML ─────┘
OMDX ─ integrity verify ───────────────────────────> Editor
```

DOCX는 Mammoth가 본문 구조/표/이미지를 읽고, OmniDocs가 WordprocessingML에서 페이지 크기·여백과
기본 머리말/꼬리말을 별도로 추출해 canonical HWPX에 적용한다. 명시적인 문단 정렬/들여쓰기,
run의 글꼴/크기/underline/highlight도 Mammoth document transform을 통해 HTML style로 보강한다.

## Export

```text
                     ┌──> .omdx  canonical + source + manifest
current canonical ───┼──> .hwpx  canonical 그대로
                     ├──> .hwp   rHWP serializer + content-loss report
                     └──> .docx  DOCX bridge
```

원본이 DOCX/HWP/HWPX이고 canonical이 import 이후 바뀌지 않았다면 해당 원본 형식 저장은 원본 바이트를
그대로 사용한다. 이 경로는 unknown/unsupported feature까지 보존하므로 가장 강한 no-op round trip이다.

## 호환성 원칙

1. 화면의 권위(authority)는 항상 canonical HWPX다.
2. 원본 포맷의 알 수 없는 데이터는 삭제하지 않고 OMDX source payload에 보존한다.
3. 변환 손실은 숨기지 않고 `compatibility`에 기록한다.
4. 포맷 adapter가 개선되어도 OMDX v1의 canonical/source 계약은 유지한다.
