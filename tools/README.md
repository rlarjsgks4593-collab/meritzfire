# 빌드 도구 (원본 엑셀 없음, 암호화된 월 데이터만 보관)

## 일일 반영
1. 새 엑셀을 임시 위치에 둔다. (엑셀은 절대 저장소에 올리지 않는다)
2. `python3 tools/build.py ingest 엑셀.xlsx --basis "11월 3일"`
   - 엑셀의 마지막 `…월 실적` 컬럼으로 월이 정해진다. 같은 월이면 그 월만 덮어쓰고, 이전 월은 그대로 고정된다.
   - 결과: `tools/months/<월>.json`(암호화) 갱신, `tools/dist/index.html` 생성
3. `tools/dist/index.html` → 저장소 `index.html`, `tools/dist/ss/index.html` → 저장소 `ss/index.html`로 교체하고 `tools/months/`, `tools/config.json` 변경분과 함께 push

## 문구만 바꿀 때
`tools/config.json` 수정 후 `python3 tools/build.py` (엑셀 불필요)

## 주의
- `tools/salt.txt`는 절대 바꾸지 않는다. 바꾸면 저장된 월 데이터가 모두 열리지 않는다.
- `pending_months`: 아직 자료가 없는 월을 버튼으로 미리 보여준다. 해당 월 자료가 들어오면 자동으로 일반 월이 된다.
- SS컬렉션 설정(목표 금액, 대상 월)은 `config.json`의 `promo`.
- `ss/` 폴더는 예전 주소를 쓰는 사람을 메인 주소로 보내는 이동 페이지다.
