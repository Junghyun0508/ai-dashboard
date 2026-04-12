# ai-dashboard

Google Sheet 기반 AI 교육 대시보드입니다.

## 사용 구조
- 데이터 소스: Google Spreadsheet (지정 7개 탭)
  - `F_피벗(요일)`, `F_피벗(본부)`, `F_피벗`
  - `I_피벗(요일)`, `I_피벗(본부)`, `I_피벗`
  - `대시보드`
- 백엔드: Google Apps Script Web App (`gas/Code.gs`)
- 프론트: GitHub Pages (`v3/index.html`)

## 1) Apps Script 배포
1. 구글 드라이브에서 Apps Script 새 프로젝트 생성
2. `gas/Code.gs` 내용을 붙여넣기
3. `Deploy > New deployment > Web app`
   - Execute as: `Me`
   - Who has access: `Anyone`
4. 배포 후 Web app URL 복사

## 2) 프론트엔드 연결
1. `v3/index.html` 파일의 아래 값 교체:
   - `const GAS_URL = 'YOUR_DEPLOYED_GAS_WEBAPP_URL';`
2. 커밋/푸시
3. GitHub Pages가 `main` 또는 원하는 브랜치/폴더를 서빙하도록 설정

## 3) 확인 포인트
- 상단 update/KPI가 `대시보드` 탭 기준으로 표시되는지
- 일별 추이: `F_피벗`, `I_피벗`
- 본부별 차트: `F_피벗(본부)`, `I_피벗(본부)`
- 요일 비중/인기강좌: `F_피벗(요일)`, `I_피벗(요일)`