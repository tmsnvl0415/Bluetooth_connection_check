# 블루투스 이어폰 연결 상태 모니터링 자동화

블루투스 이어폰의 연결 및 미연결 상태를 실시간으로 모니터링하고, 전체 연결 시간 및 끊김 시간을 Excel 리포트로 자동 정리하는 Python 스크립트입니다.



## 주요 기능

- 이어폰 연결 상태를 주기적으로 체크
- 총 연결 시간 / 미연결 시간 자동 계산
- 종료 시점에 Excel 리포트 자동 생성 (`bluetooth_summary_YYYYMMDD_HHMM.xlsx`)
- 배치 파일 `.bat`로 실행 가능



## 실행 방법

1. Python 3 설치
2. 라이브러리 설치:

```
pip install openpyxl
```

3. 스크립트 실행:

```
python bluetooth_check.py
```

또는 `run_bt_check.bat` 배치파일을 더블클릭하면 자동 실행



## 실행 결과

파일 이름은 자동으로 `bluetooth_summary_YYYYMMDD.xlsx` 형식으로 저장 (하나의 워크북에 여러 시트가 포함)

- 프로젝트 이름, 전체 이슈 수, 잔여 이슈 수 요약
- 전체 이슈 요약 테이블
  └ 이벤트 단계(DVT/EVT 등) × 우선순위(A~D) × 상태(Open~Closed) 별로 정리된 표
- 이번 주 등록 이슈 테이블
  └ 상태/우선순위별로 이번 주에 생성된 이슈만 따로 정리한 표
- 누적 이슈 커브 그래프 (자동 이미지 삽입)
  └ Total Bug / Resolved / Not Resolved 추이를 날짜별 선그래프로 시각화
- bluetooth_connect_report.txt
  └ 스크립트 실행 중 이어폰 연결/미연결 상태가 시간별로 기록된 로그 파일



## 파일 구조 예시

```
📁 Bluetooth_connection_check/
 ┣ 📄 Bluetooth_connect.py           → 블루투스 연결 상태 모니터링 메인 스크립트
 ┣ 📄 run_bt_check.bat               → 자동 실행용 배치 파일 (더블클릭 실행)
 ┣ 📄 README.md                      → 프로젝트 설명 문서
 ┣ 📄 .gitignore                     → 자동 생성 리포트 제외 설정
 ┣ 📄 bluetooth_connect_report.txt   → 실시간 연결 로그 기록 파일 (자동 생성됨)
 ┗ 📄 bluetooth_summary_*.xlsx       → 실행 시 자동 생성되는 리포트 파일들

```


## 👩‍💻 작성자
김예지  
SQA Engineer
GitHub: [@tmsnvl0415](https://github.com/tmsnvl0415)
