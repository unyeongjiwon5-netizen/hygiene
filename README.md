# PA센터 공용 위생점검 앱 v16.5

v16.1 인라인 관리자 편집본을 Supabase 공용 데이터베이스에 연결한 버전입니다.

## 공용 저장 대상

- 점검자 목록 및 사용 상태
- 본동·부속동·화장실 점검항목과 순서
- 시설 및 화장실 점검기록
- 이상 내용과 수정이력
- 관리자 비밀번호

최근 점검자 조합만 사용자 기기의 브라우저에 저장됩니다.

## 환경변수

프로젝트 루트에 `.env.local`을 만들고 아래 값을 입력합니다.

```env
NEXT_PUBLIC_SUPABASE_URL=...
NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY=...
SUPABASE_SECRET_KEY=...
```

`SUPABASE_SECRET_KEY`는 브라우저 코드로 전달되지 않고 `/api/data` 서버 경로에서만 사용합니다. GitHub에 업로드하지 마세요.

## 로컬 실행

```bash
npm install
npm run dev
```

또는 기존 pnpm 환경에서는:

```bash
pnpm install
pnpm dev
```

접속: http://localhost:3000

## Vercel 환경변수

Vercel 프로젝트의 Settings → Environment Variables에 같은 3개 변수를 등록합니다. Production, Preview, Development를 모두 선택하는 것을 권장합니다.

## 동작 방식

- 앱 시작 시 Supabase 공용 데이터를 불러옵니다.
- 저장·수정·삭제 직후 공용 데이터를 다시 불러와 화면을 갱신합니다.
- 앱을 열어둔 동안 30초마다 최신 데이터를 다시 확인합니다.
- 데이터베이스 접근은 Next.js 서버 API를 거치므로 Secret Key가 브라우저에 노출되지 않습니다.


## v16.5 변경사항

- 캘린더를 연결형 표 구조로 변경하고 모바일에서 본/부 상태점 표시
- 상단 현황 카드 호버의 이중 테두리 제거
- 시설 월간 현황 첫 열 고정 및 기존 크기·간격 유지
- 시설 월간 현황의 v16.2 아이콘·텍스트 호버 복원
- 관리자 점검 제외일 등록·수정·삭제 기능 추가
- 점검 제외일은 `app_settings`의 `excluded_dates` 항목에 JSON으로 저장되므로 별도 Supabase 테이블 생성 불필요
- 주말 색상 변경은 적용하지 않음
