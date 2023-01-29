# kiwoomAPI
키움증권 자동화 프로그램 구축

# 설계 설명
 - 키움증권 API 기반 설계
 - 키움증권 KOA 를 참고하여 원하는 기능 설계
 - dynamiccall 또는 키움 api 메서드를 통해 키움증권 서버에 원하는 기능 요청 -> 요청에 대한 이벤트 발생 시, 코딩으로 이벤트 수신 코드 필요

# 파일 설명
1. Auto_trade / 주식매매 자동화, 보유계좌 리스트 추출
2. findPattern / 조건검색 주식 시각화 (PANDAS plot 사용)
3. stockAnalysis / 조건검색 주식 엑셀저장
4. verifyCondition / 조건검색 주식 익일 수익 
