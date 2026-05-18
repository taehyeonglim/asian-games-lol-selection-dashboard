# Asian Games LoL Selection Weight Dashboard

2026 아시안게임 리그 오브 레전드 대표 선발 1차 검토지표를 사용자가 직접 가중치로 조절해보는 정적 대시보드입니다.

## 바로 사용하기

브라우저에서 `index.html`을 열면 됩니다. GitHub Pages에 올릴 경우 저장소 루트의 `index.html`이 공개 페이지가 됩니다.

## 포함된 비교

- Top: Zeus vs Kiin
- Jungle: Canyon vs Oner
- Bot: Gumayusi vs Peyz vs Viper

Mid와 Support는 v1 대시보드에서 제외했습니다.

## 지표

대시보드는 `asian_games_lol_selection_data.xlsx`의 `Candidate_Summary` 시트를 읽어 생성됩니다.

- 지역 정규리그 성적
- 월드 챔피언십 성적
- KeSPA Cup 성적
- All-Pro 수상
- POG/POM 선정

기본 가중치는 `17 / 26 / 5 / 5 / 47`입니다. 모든 컴포넌트는 0-100 스케일 값을 사용하며, 슬라이더 합계는 항상 100%로 유지됩니다.

각 가중치 박스의 `고정` 체크박스를 켜면 해당 수치는 잠기고, 다른 슬라이더를 움직일 때 재분배 대상에서 제외됩니다.

## 데이터 주의사항

이 데이터는 공식 선발 가중치가 아니라 공개 지표 기반 검토용 데이터입니다. 특히 원딜 비교는 LPL POG/POM, 일부 LPL 정규리그 행, 상세 개인지표가 미완성이라 참고용으로 봐야 합니다.

현재 데이터에는 월즈 Finals MVP 같은 개인 수상 보너스가 별도 컴포넌트로 반영되어 있지 않습니다.

## 다시 생성하기

Python 3 환경에서 다음을 실행합니다.

```bash
pip install -r requirements.txt
python build_selection_dashboard.py
```

그러면 `selection_dashboard.html`과 GitHub Pages용 `index.html`이 함께 생성됩니다.

## GitHub Pages 배포

1. 이 폴더를 GitHub 저장소로 푸시합니다.
2. GitHub 저장소의 `Settings > Pages`로 이동합니다.
3. `Deploy from a branch`를 선택합니다.
4. Branch는 `main`, folder는 `/root`를 선택합니다.
5. 몇 분 뒤 `https://<owner>.github.io/<repo>/` 주소로 접속합니다.
