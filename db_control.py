import os
import pandas as pd
import sqlite3

# SQLite3 데이터베이스 연결
conn = sqlite3.connect("addresses.db")
cursor = conn.cursor()

# 데이터 폴더 경로
data_folder = "./data"

# 테이블 생성
cursor.execute('''
CREATE TABLE IF NOT EXISTS name_address (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    phone TEXT,
    address TEXT
)
''')

# 데이터 폴더에서 모든 엑셀 파일 검색
excel_files = [os.path.join(data_folder, f) for f in os.listdir(data_folder) if f.endswith(('.xlsx', '.xls'))]

# 엑셀 파일 데이터를 데이터베이스에 삽입
for file in excel_files:
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file)

        # 데이터프레임에서 데이터 추출 및 삽입
        for _, row in df.iterrows():
            name = row.get("받는분성명", None)
            phone = row.get("받는분전화번호", None)
            address = row.get("받는분주소(전체, 분할)", None)

            if name and phone and address:  # 모든 필드가 존재할 경우만 삽입
                # 중복 확인
                cursor.execute('''
                SELECT COUNT(*) FROM name_address
                WHERE name = ? AND phone = ? AND address = ?
                ''', (name, phone, address))
                if cursor.fetchone()[0] == 0:  # 데이터가 없을 경우 삽입
                    cursor.execute('''
                    INSERT INTO name_address (name, phone, address)
                    VALUES (?, ?, ?)
                    ''', (name, phone, address))
                else:
                    print(f"중복된 데이터로 인해 삽입되지 않음: {name}, {phone}, {address}")

        print(f"'{file}'의 데이터를 성공적으로 처리했습니다.")
    except Exception as e:
        print(f"파일 '{file}' 처리 중 오류 발생: {e}")

# 변경 사항 저장 및 연결 닫기
conn.commit()
conn.close()
