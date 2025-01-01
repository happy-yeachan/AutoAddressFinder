
# Excel Address Tool
<img width="1312" alt="스크린샷 2025-01-02 오전 8 18 47" src="https://github.com/user-attachments/assets/0b1be29f-098a-40c6-aa5e-dcc86a18fbae" />
`Excel Address Tool`은 주소 데이터를 관리하고, 사용자 친화적인 UI를 통해 주소 검색, 수정, 삭제, 및 엑셀 파일로 데이터를 내보내는 기능을 제공하는 데스크탑 애플리케이션입니다. PyQt6와 SQLite를 사용하여 개발되었습니다.

## 주요 기능

- **주소 데이터 관리**:
  - 주소 데이터의 추가, 수정, 삭제를 지원.
  - 테이블에서 항목을 선택하면 데이터를 편집할 수 있는 입력 창에 자동으로 로드.
- **주소 검색**:
  - 이름을 입력하여 관련 주소를 빠르게 검색.
  - 검색 결과를 클릭하면 입력창에 자동으로 데이터가 채워짐.
- **엑셀 파일로 데이터 내보내기**:
  - 현재 데이터베이스의 모든 데이터를 엑셀 파일로 저장.
  - 데이터 정렬 및 열 너비 자동 조정 기능 포함.
- **사용자 인터페이스**:
  - 직관적인 PyQt6 기반의 UI.
  - 여러 버튼(추가, 수정, 삭제, 취소)으로 다양한 작업 수행 가능.

## 설치 방법

### 필수 조건

- Python 3.8 이상
- 가상환경을 사용하는 것을 권장합니다.

### 의존성 설치

```bash
git clone https://github.com/yourusername/excel-address-tool.git
cd excel-address-tool
python -m venv venv
source venv/bin/activate  # Windows의 경우 venv\Scripts\activate
pip install -r requirements.txt
```

### 실행

```bash
python main.py
```

## 사용 방법

### 데이터 추가
1. 입력창에 보내는 사람과 받는 사람의 정보를 입력.
2. 품목명과 갯수를 입력한 후 **[추가]** 버튼 클릭.

### 데이터 수정
1. 테이블에서 수정할 항목을 클릭하여 데이터를 입력창에 로드.
2. 필요한 데이터를 수정한 후 **[수정]** 버튼 클릭.

### 데이터 삭제
1. 테이블에서 삭제할 항목을 클릭.
2. **[삭제]** 버튼 클릭.

### 엑셀 파일로 내보내기
1. **[엑셀 추출]** 버튼 클릭.
2. 저장할 파일 이름과 경로를 선택 후 저장.

### 데이터 초기화
- **[초기화]** 버튼 클릭하여 모든 데이터를 삭제.

## 데이터베이스 구조

- **Orders Table**:
  - `id`: 고유 ID (자동 증가)
  - `sender_name`: 보내는 사람 이름
  - `sender_phone`: 보내는 사람 전화번호
  - `sender_address`: 보내는 사람 주소
  - `receiver_name`: 받는 사람 이름
  - `receiver_phone`: 받는 사람 전화번호
  - `receiver_address`: 받는 사람 주소
  - `item_name`: 품목명
  - `quantity`: 품목 갯수

- **Name-Address Table**:
  - `id`: 고유 ID (자동 증가)
  - `name`: 이름
  - `phone`: 전화번호
  - `address`: 주소

## 기술 스택

- **프로그래밍 언어**: Python
- **GUI**: PyQt6
- **데이터베이스**: SQLite
- **엑셀 처리**: pandas, openpyxl

## 기여

기여를 환영합니다! 문제를 발견하거나 개선 사항이 있다면 [이슈](https://github.com/yourusername/excel-address-tool/issues)를 생성하거나 풀 리퀘스트를 제출해주세요.

## 라이선스

이 프로젝트는 [MIT 라이선스](LICENSE) 하에 배포됩니다.
