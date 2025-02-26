import os
import json
import pandas as pd

# Json 중첩객체를 Excel로 변환하는 코드드

# JSON 파일들이 있는 폴더 경로 (예: "./data")
folder_path = "."

# 결과를 담을 리스트
records = []

# 폴더 내 모든 파일을 확인
for filename in os.listdir(folder_path):
    if filename.endswith(".json"):
        file_path = os.path.join(folder_path, filename)
        
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # TimeSeriesData 배열 가져오기 (없으면 빈 리스트)
        ts_data_list = data.get("TimeSeriesData", [])

        # TimeSeriesData 내에서 Total_Labeling 정보만 추출
        for ts_item in ts_data_list:
            total_label = ts_item.get("Total_Labeling", {})
            
            # 필요한 필드만 record로 만들기
            record = {
                "FileName": filename,  # 어느 파일에서 왔는지 확인용
                "DataType": total_label.get("DataType", ""),
                "Estimation": total_label.get("Estimation", ""),
                "Reason": total_label.get("Reason", "")
            }
            records.append(record)

# 모은 데이터를 DataFrame으로 만들기
df = pd.DataFrame(records)


# Reason 컬럼의 고유값(Unique) 개수 계산
unique_reason_count = df["Reason"].nunique()
print("Reason 컬럼 고유값 개수:", unique_reason_count)

# 실제 고유값 목록을 보고 싶다면
unique_reasons = df["Reason"].unique()
print("Reason 컬럼 고유값 목록:", unique_reasons)

# 혹은 각 고유값별로 몇 번 나오는지 보고 싶다면
reason_value_counts = df["Reason"].value_counts()
print(reason_value_counts)

# Reason 컬럼의 고유값(Unique) 개수
unique_reason_count = df["Reason"].nunique()

# Reason 컬럼의 실제 고유값 목록
unique_reasons = df["Reason"].unique()

# Reason 컬럼의 값별 등장 횟수
reason_value_counts = df["Reason"].value_counts()

# ====== 엑셀로 내보내기 ======
# 한 번에 여러 시트에 쓰려면 ExcelWriter 사용
output_file = "reason_analysis.xlsx"

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    # 1) 전체 데이터 df
    df.to_excel(writer, sheet_name="AllData", index=False)
    
    # 2) Reason 통계: 고유값 개수
    df_summary = pd.DataFrame({
        "Metric": ["Unique Reason Count"],
        "Value": [unique_reason_count]
    })
    df_summary.to_excel(writer, sheet_name="Summary", index=False)
    
    # 3) Reason 고유값 목록(각 고유값을 행으로)
    df_unique_reasons = pd.DataFrame({"Unique Reasons": unique_reasons})
    df_unique_reasons.to_excel(writer, sheet_name="UniqueValues", index=False)
    
    # 4) Reason 값별 등장 횟수
    # value_counts() 결과를 DataFrame으로 변환
    df_reason_counts = reason_value_counts.reset_index()
    df_reason_counts.columns = ["Reason", "Count"]
    df_reason_counts.to_excel(writer, sheet_name="ValueCounts", index=False)

print("엑셀 파일 생성 완료:", output_file)

# Excel 파일로 내보내기
#df.to_excel("output_Total_Labeling.xlsx", index=False)
