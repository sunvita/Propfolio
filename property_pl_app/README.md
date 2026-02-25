# Property P&L Portfolio Builder

PDF 업로드 → 자동 파싱 → Excel P&L 생성 (Streamlit 웹앱)

## 실행 방법

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. 앱 실행
streamlit run app.py
```

브라우저에서 `http://localhost:8501` 접속

## 기능

| 단계 | 내용 |
|---|---|
| Step 1 | 부동산 기본정보 설정 (부동산 수, FY 시작월, 자산정보) |
| Step 2 | PDF 업로드 (렌탈 명세서 / 은행 거래내역 / 공과금 청구서) |
| Step 3 | 파싱 결과 검토·수정·수동 입력 |
| Step 4 | Excel 생성 및 다운로드 |

## 지원 PDF 형식

- **렌탈/명세서**: Certainty Property 등 관리회사 명세서 (Money In/Out/EFT)
- **은행 거래내역**: 표 형식의 거래내역 (날짜·내역·금액 자동 추출)
- **공과금**: 전기·수도·가스·인터넷 청구서 (금액 자동 추출)

## Excel 출력 구조

- **부동산별 탭** (IP#1~IP#5): P&L 86칼럼 + KPI 요약 Table A
- **Summary 탭**: Table B (자산정보·수익률) + Table A (포트폴리오 성과 집계)
- 컬러 코딩: 파란=수동입력 / 검은=수식 / 초록=탭간참조 / 노란=FY합계 / 회색=템플릿
