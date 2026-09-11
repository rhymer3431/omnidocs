# OmniDocs

OmniDocs는 **HWP, HWPX, DOCX를 OMDX라는 하나의 내부 문서 형식으로 정규화한 뒤 같은 편집기와 같은 레이아웃 엔진에서 편집**하기 위한 워드프로세서 프로젝트입니다.

현재 0.1 MVP는 새 워드프로세서 엔진을 직접 구현하지 않고 다음 오픈소스를 재사용합니다.

- `@rhwp/editor` / `@rhwp/core` (MIT): 실제 편집, 조판, 표, 커서, Undo/Redo, HWP/HWPX 읽기/쓰기
- `mammoth` (BSD-2-Clause): DOCX → semantic HTML import
- `JSZip` (MIT): OMDX/OOXML 컨테이너 처리
- React + Vite: OmniDocs shell

## 핵심 구조

```text
HWP  ── rHWP ──────────┐
HWPX ─ rHWP ───────────┼──> OMDX ──> canonical document.hwpx
DOCX ─ Mammoth+OOXML ──┘                    │
OMDX ─ verify ──────────────────────────────┘
                                            │
                                  rHWP Document Core / Studio
                                            │
                          ┌─────────────────┼───────────────┐
                          ▼                 ▼               ▼
                         HWP               DOCX            OMDX
```

HWP와 DOCX에 서로 다른 편집기를 제공하지 않습니다. 파일을 여는 순간 OMDX로 정규화하고,
OMDX의 `content/document.hwpx`만 화면 조판의 기준으로 사용합니다. 따라서 가져온 원본 형식과 무관하게
편집 중에는 동일한 페이지네이션/표/폰트/커서 엔진을 사용합니다.

OMDX는 원본 파일도 `source/original.*`로 보존합니다. 문서를 수정하지 않고 원래 형식으로 저장하면
원본 바이트를 그대로 반환하므로 알 수 없는 Word/HWP 확장 기능도 손실되지 않습니다.

## 실행

```bash
npm install
npm run vendor:rhwp-studio
npm run dev
```

브라우저에서 Vite 주소를 연 뒤 `.hwp`, `.hwpx`, `.docx`, `.omdx` 파일을 선택합니다.

기본값은 `public/rhwp/`에 vendoring 된 rHWP Studio를 같은 origin의 `/rhwp/index.html` 경로에서 사용합니다.
따라서 실행 중 공개 GitHub Pages Studio에 의존하지 않습니다. 개발 또는 비교 목적일 때만 아래처럼
외부 Studio URL을 명시적으로 override할 수 있습니다.

```bash
VITE_RHWP_STUDIO_URL=https://your-studio.example.com/ npm run dev
```

## 검증

```bash
npm test
npm run test:compat
npm run build
```

`test:compat`은 OMDX 패키지/체크섬 테스트, DOCX layout/style bridge 테스트,
HWP/HWPX 저장 smoke test와 DOCX import smoke test를 함께 실행합니다.

## 현재 구현 범위

- HWP/HWPX: rHWP로 canonical HWPX를 생성한 뒤 OMDX 내부 문서로 편집
- DOCX: Mammoth + OOXML sidecar parser를 통한 OMDX import
- DOCX import 시 페이지 크기, 여백, 기본 머리말/꼬리말 복원
- DOCX import 시 명시적 문단 정렬/들여쓰기, 글꼴/크기, underline/highlight 보강
- DOCX export: canonical HWPX XML을 직접 읽어 WordprocessingML/OOXML package 생성
- DOCX export 시 문단/글자 서식, 번호 목록, 표/병합 셀, 이미지, 페이지 크기/여백, 다단, 머리말/꼬리말 직접 매핑
- HWPX floating image를 Word `wp:anchor`, floating table을 `w:tblpPr` 위치 정보로 직접 매핑
- HWPX 수식 script의 분수/제곱·첨자/제곱근/합·곱·적분을 편집 가능한 Word OMML로 직접 변환
- OMDX v1 ZIP package 및 SHA-256 무결성 검증
- 원본 HWP/HWPX/DOCX immutable payload 보존
- 수정하지 않은 원본 형식 저장 시 byte-exact round trip
- 파일 형식과 무관한 하나의 canonical 레이아웃/편집 화면
- 저장 형식 OMDX/HWP/HWPX/DOCX 전환
- 포맷 어댑터와 편집 엔진 인터페이스 분리
- self-hosted rHWP Studio 배포
- 기본 단위/포맷 호환성 smoke test

## 알려진 제한

편집된 문서를 DOCX로 저장할 때 HTML을 거치지 않습니다. floating 그림/표와 기본 수식은 OOXML 구조로
직접 변환하지만, HWP 전용 복잡 수식 문법, 비정형 tight/through contour, Word/HWP 전용 field,
일부 도형/차트/OLE는 아직 근사되거나 제외될 수 있습니다. 저장 시 해당 항목은 호환성 경고로 표시합니다.
원본 DOCX를 수정하지 않았다면 OMDX가 원본 바이트를 그대로 반환합니다.

OMDX 파일 규격은 [docs/OMDX.md](./docs/OMDX.md), 전체 설계는 [ARCHITECTURE.md](./ARCHITECTURE.md)를 참고하세요.
