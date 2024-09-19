# %% 1. 필요한 패키지 호출 -------------------
import pandas as pd
import numpy as np
import random
import os
from openpyxl import load_workbook
from io import BytesIO
from IPython.display import display, HTML
import base64





# %% 2. 단과대학/학부명 입력 + 파일 호출 및 읽기 -------------------
# 파일 경로 지정 (사용자의 파일 경로에 맞게 수정)
file_path = "C:/Users/qosgu/OneDrive/바탕 화면/공공안전학부.xlsx"

# 엑셀 파일 읽기
df = pd.read_excel(file_path, engine='openpyxl')

# 파일 출력
df





# 이름과 전화번호 뒷자리 기준으로 중복 응답자 확인 및 별도 데이터프레임에 저장
duplicates = df[df.duplicated(subset=['이름', '전화번호 뒷자리'], keep=False)]
if len(duplicates) > 0:
    print("설문 중복 응답자:")
    for index, row in duplicates.iterrows():
        print(f"{row['이름']}, {row['전화번호 뒷자리']}")
    # [구분] 컬럼 추가 후 설문 중복 응답자 정보 기록
    duplicates['구분'] = '설문 중복 응답자'
else:
    print("설문 중복 응답자가 없습니다.")
    
# 중복 응답자를 데이터프레임에서 제거하고, 별도로 저장
df = df.drop_duplicates(subset=['이름', '전화번호 뒷자리'])
df_excluded = duplicates.copy()  # 중복 응답자를 별도 저장




# %% 3. 전공명 로드 및 정렬 -------------------
# 지망 관련 컬럼만 추출하여 정렬
sample_row = df.iloc[0]  # 응답 결과의 첫 번째 행을 선택
major_columns = [col for col in df.columns if '지망' in col]
major_name = sample_row[major_columns].unique()  # 중복되지 않게 전공명 추출
major_name = [major for major in major_name if pd.notna(major)]  # NaN 값 제거
major_name.sort()  # 전공명을 오름차순으로 정렬





# %% 4. 인원별 지망 중복 체크 -------------------
# 각 응답자마다 같은 전공을 여러 지망에 선택했는지 확인 및 별도 데이터프레임에 저장
duplicates_per_major = []

# 기본적으로 빈 데이터프레임 생성
df_major_duplicates = pd.DataFrame()

for index, row in df.iterrows():
    selected_majors = row[major_columns].values
    if len(selected_majors) != len(set(selected_majors)):
        duplicates_per_major.append(row)

if len(duplicates_per_major) > 0:
    print("전공 중복 체크자:")
    for entry in duplicates_per_major:
        print(f"{entry['이름']}, {entry['전화번호 뒷자리']}, {entry[major_columns].values}")
    # 전공 중복 체크자 데이터프레임 생성 후 [구분] 컬럼 추가
    df_major_duplicates = pd.DataFrame(duplicates_per_major)
    df_major_duplicates['구분'] = '전공 중복 체크자'
else:
    print("전공 중복 체크자가 없습니다.")

# 전공 중복 체크자를 데이터프레임에서 제거하고, 별도로 저장
if not df_major_duplicates.empty:
    df = df.drop(df_major_duplicates.index)
    df_excluded = pd.concat([df_excluded, df_major_duplicates], ignore_index=True)  # 기존 제외된 응답자와 결합





# %% 5. 전공별 배치 정원 입력 -------------------
# 전공별 배치 정원 입력 받기
major_limits = {}
for major in major_name:
    while True:
        try:
            limit = int(input(f"{major}의 배치 정원을 입력하세요: "))
            major_limits[major] = limit
            break
        except ValueError:
            print("유효한 숫자를 입력하세요.")

# 입력된 전공별 배치 정원을 줄바꿈하여 출력
print("\n입력한 전공별 배치 정원:")
for major, limit in major_limits.items():
    print(f"{major}: {limit}명")





# %% 6. 전공별 전체 지망 인원수 확인 -------------------
# 전공별 지망 인원수 확인
major_applications = {major: [0] * len(major_columns) for major in major_name}

for index, row in df.iterrows():
    for i, major in enumerate(major_columns):
        if row[major] in major_applications:
            major_applications[row[major]][i] += 1

# 데이터프레임으로 변환
major_applications_df = pd.DataFrame(major_applications, index=[f'{i+1}지망' for i in range(len(major_columns))])

major_applications_df





# %% 7. 1~n지망 배치 후 배치 결과 발표 -------------------
# 전공 배정 함수
def allocate_majors(df, major_columns, major_limits):
    # 배정 결과를 담을 딕셔너리 초기화
    assigned_students = {major: [] for major in major_name}
    remaining_students = df.copy()  # 아직 배정되지 않은 학생 리스트
    
    # 각 지망 순서대로 배정 진행 (1지망부터 마지막 지망까지)
    for priority in range(len(major_columns)):  # 지망 개수만큼 반복
        current_choice_col = major_columns[priority]
        print(f"\n===== {priority + 1}지망 배정 진행 중... =====\n")
        
        for major in major_name:
            # 각 전공에 해당 지망을 선택한 학생들을 필터링
            chosen_students = remaining_students[remaining_students[current_choice_col] == major]
            
            # 전공 배정 가능 인원이 남아 있는 경우
            available_slots = major_limits[major] - len(assigned_students[major])
            if available_slots > 0:
                if len(chosen_students) <= available_slots:
                    # 배정 가능 인원이 충분한 경우, 모두 배정
                    assigned_students[major].extend(chosen_students['이름'].tolist())
                    print(f"{major}: {len(chosen_students)}명 배정 완료")
                else:
                    # 배정 가능 인원보다 많은 경우, 초과 인원 처리
                    selected_students = chosen_students.sample(available_slots)
                    assigned_students[major].extend(selected_students['이름'].tolist())
                    print(f"{major}: {len(chosen_students) - available_slots}명 초과, {available_slots}명 무작위 배정")
            else:
                print(f"{major}: 이미 정원이 모두 찼습니다.")
        
        # 현재 지망에서 배정된 학생들을 전체 남은 학생 목록에서 제거
        remaining_students = remaining_students[~remaining_students['이름'].isin([s for students in assigned_students.values() for s in students])]
        print("\n\n")
        
        # 남은 학생이 없는 경우 종료
        if remaining_students.empty:
            print(f"\n{priority + 1}지망에서 모든 학생이 배정 완료되었습니다.\n")
            break
    
    return assigned_students, remaining_students





# -------------------
# 전공 배치 실행
assigned_students, remaining_students = allocate_majors(df, major_columns, major_limits)

# 배정 결과를 테이블로 출력
# 각 학생의 이름 + 전화번호 뒷자리로 데이터 구성
formatted_results = {
    major: [f"{name}({df[df['이름'] == name]['전화번호 뒷자리'].values[0]})" for name in students]
    for major, students in assigned_students.items()
}

# 최대 학생 수에 맞춰 테이블을 형성
max_students = max(len(students) for students in formatted_results.values())
for major in formatted_results:
    formatted_results[major] += [''] * (max_students - len(formatted_results[major]))  # 빈 공간 채우기




# 미배정 학생 처리
if not remaining_students.empty:
    print("\n전공 미배정 학생:")
    print(remaining_students[['이름', '전화번호 뒷자리']])
    # 전공 미배정 학생 데이터프레임 생성 후 [구분] 컬럼 추가
    remaining_students['구분'] = '전공 미배정 학생'
    # 미배정 학생을 별도 데이터프레임에 저장
    df_excluded = pd.concat([df_excluded, remaining_students], ignore_index=True)
else:
    print("모든 학생이 배정되었습니다.")




# %% 8. 최종 전공 배치 결과 발표 -------------------
# 최종 배치 결과 출력 및 결측치 재배치
final_result = pd.DataFrame(formatted_results)

# 공백 문자열을 NaN으로 변환하여 결측치로 인식하게 처리
final_result.replace(r'^\s*$', np.nan, regex=True, inplace=True)

# 각 열에 대해 텍스트 오름차순 정렬 후 NaN 값을 맨 뒤로 보냄
for col in final_result.columns:
    # NaN이 아닌 값만 따로 정렬
    non_null_values = final_result[col].dropna().sort_values()
    
    # NaN 값을 NaN의 개수만큼 추출
    nan_count = final_result[col].isna().sum()
    nan_values = [np.nan] * nan_count  # NaN 값 리스트
    
    # NaN이 아닌 값 + NaN 값을 결합하여 새로운 리스트 생성
    sorted_values = non_null_values.tolist() + nan_values
    
    # 새로운 리스트를 해당 열에 할당
    final_result[col] = pd.Series(sorted_values)

# 재배치된 최종 배치 결과 출력
final_result




# %% 9. 결과 엑셀 파일 다운로드 -------------------
# 엑셀 파일 저장 함수 수정
def create_download_link(df_result, df_excluded, filename="전공배치결과.xlsx"):
    towrite = BytesIO()
    
    # 엑셀 파일에 행번호(index)를 1번부터 포함하여 저장
    df_result.index = df_result.index + 1  # 행 번호를 1번부터 시작하도록 설정
    with pd.ExcelWriter(towrite, engine='openpyxl') as writer:
        # 전공 배치 결과를 첫 번째 시트에 저장
        df_result.to_excel(writer, sheet_name='전공 배치 결과', index=True)
        
        # 제외된 응답자를 두 번째 시트에 저장
        if not df_excluded.empty:
            # "구분" 컬럼이 포함된 경우 그대로 저장
            df_excluded.to_excel(writer, sheet_name='제외된 응답자', index=False)
        else:
            # "구분" 컬럼을 가진 빈 데이터프레임 생성 및 저장
            empty_excluded = pd.DataFrame(columns=['이름', '전화번호 뒷자리', '제외된 이유'])
            empty_excluded.to_excel(writer, sheet_name='제외된 응답자', index=False)

    towrite.seek(0)
    
    # base64 인코딩
    b64 = base64.b64encode(towrite.read()).decode()
    
    # 다운로드 링크 생성
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">엑셀 파일 다운로드</a>'
    return HTML(href)

# 다운로드 링크 생성 및 표시
display(create_download_link(final_result, df_excluded))
