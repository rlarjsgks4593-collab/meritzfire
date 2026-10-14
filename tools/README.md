# 빌드 도구 (데이터 없음)

일일 갱신 순서:
1. 새 엑셀을 `tools/data/latest.xlsx`로 둔다. (엑셀은 절대 저장소에 올리지 않는다)
2. `python3 tools/build.py` 실행 → `tools/dist/index.html` 생성
3. 생성된 파일을 저장소 맨 위 `index.html`로 교체해서 push

문구(하단 안내, 기준일)는 `tools/config.json`에서 수정한다.
