# Extract_Dart_FinData


## 요약

> DART (금융감독원 전자공시시스템) 에서 재무정보를 찾을 때 테이블화가 되어 있지 않고 일일이 수작업으로 입력해야 하는 불편함을 개선할 목적으로 제작

> 분기별로 테이블화된 재무정보를 타 산업군 데이터와 함께 이용하여 예측 모형 등을 제작하는 데 활용할 수 있음

> 개인별 API 코드
및 회사 정보 등을 입력하여 분기별로 재무상태표, 손익계산서, 현금흐름표
데이터를 각 회사별로 추출할 수 있음


## 설치 방법

> R (R console(GUI), R studio 상관 없음) 에서

> source("https://raw.githubusercontent.com/jeong7683/Extract_Dart_FinData/master/Extract_Dart_FinData.R") 입력 시 바로 실행됩니다.


## 소스 코드

> 함수의 경우 main 함수 역할을 하며 정보를 입력받는 StartExtract 함수와 실제로 API를 이용해 재무정보를 얻어 재무상태표, 손익계산서, 현금흐름표 데이터를 만드는 DartJSONtoExcel 함수로 이루어져 있습니다.

> StartExtract: API 입력 – 회사 종목코드 입력 – 데이터 추출을 시작할 날짜 입력 (- 데이터 추출을 끝낼 날짜 입력) // 괄호는 생략 가능

> DartJSONtoExcel: API 이용해 파일 다운로드 – 파일 입력 및 단위 통일(원) – 변수 통일(일부) – 연간, 누적데이터의 분기 환산 – 각 데이터별 폴더 분리 후 저장


## 라이센스 정보


> 저작권자: 정현진 (Hyunjin Jeong)

> 제작일: 2017년 4월 8일


