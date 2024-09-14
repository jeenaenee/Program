# Temporary_Major_Assign_Program(2022)

# %% 0. 패키지 호출 
import pandas as pd     # type: ignore # 파일 호출 및 반환
import random           # 난수 생성을 통한 가전공 배치
import os               # 파일 경로 가져오기
import openpyxl         # type: ignore # VS에서 xlsx 파일을 불러오기 위한 패키지





# %% 1. 수요조사 결과 호출 및 할당------------------------------------------------------------------------------------------
C_name = str(input("단과대학명을 입력하세요.(ex. 사회과학대학): "))   # C_name : College_name
F_name = str(input("학부명을 입력하세요.(ex. 경제학부): "))          # F_name : Faculty_name
print(F_name, '\n가전공 수요조사 엑셀 파일 경로와 파일명을 복사하여 저장해주세요.\n')
print("예시 : M = pd.read_excel('파일 경로/.../파일명.xlsx') \n\t※ A는 임의의 문자, 파일 경로는 /로 구분\n")
M = pd.read_excel("C:/Users/qosgu/OneDrive/바탕 화면/공공안전학부.xlsx", engine='openpyxl')   # M : Major


# 불필요한 컬럼 제거 및 데이터프레임으로 변환 후 확인
M = M.drop(['TIMESTAMP', '가전공 배치를 위한 개인정보 동의'], axis=1)     # 불필요한 컬럼 삭
M = pd.DataFrame(M)     # 데이터프레임으로 변환
M                       # 데이터프레임 확인


# 중복 응답자 확인
M[M.duplicated(keep = False)]   # 중복응답자가 있는 경우, 인덱스 번호 출력





# %% 2. 학과명 로드 및 정렬------------------------------------------------------------------------------------------------
M_count = int(input("학부 내 전공 수를 입력하세요.(숫자만): "))    # 학부 내 전공 개수 입력


# 지망별 학과명 추출
M_name = sorted(list(M.loc[0][1:M_count+1]))    # 첫 번째 응답자의 1~n지망을 추출하고, 가나다 순으로 정렬하여 리스트로 반환
M_name      # 확인


# 전공별 1~n지망 분류 리스트 생성
for i in M_name:                # 학부 내 전공명을 i에 반복 할당
    for j in range(M_count):    # 학부 내 전공 개수를 j에 반복 할당
        globals()["{}_{}지망".format(i, j+1)] = []    # 전공별 1~5지망 빈 리스트 생성





# %% 3. 인원별 지망 중복 체크
# 지망 중복 응답자를 모을 리스트
duplicate = []

# 지망 내 전공 개수와 전체 전공 개수 불일치 시 duplicate에 추가
for i in range(len(M)):                             # 전체 응답자 수 만큼의 범위
    if len(set(M.loc[i][1:M_count+1])) != M_count:  # 전체 응답자중 응답한 전공의 개수와 전체 전공의 개수 불일치 여부 파악
        duplicate.append(M.loc[i])                  # ex. 학부 내 5개 전공이 있지만, 1~5지망 중 1~4개의 전공만 체크하면 추가
        print(M.loc[i])                             # 어떤 데이터인지 확인하기 위해 출력


# 중복 체크자 확인 결과
if len(duplicate) == 0:     # 중복 체크자가 없는 경우
    print("1~{}지망 중복 체크자가 없습니다.\n".format(M_count))
else:                       # 중복 체크자가 있는 경우
    print("1~{}지망 중복 체크자는 {}명 입니다.\n".format(M_count, len(duplicate)))
    print(duplicate)





# %% 4. 전공별 배치정원 입력-----------------------------------------------------------------------------------------------
# 58 / 41 / 57 / 27 / 28
print("-----전공별 배치정원을 입력해주세요.-----")
print("잘못 입력한 경우, \'Ctrl+c\' 커맨드 후 다시 실행 하세요.")
for i in range(M_count):
    globals()["{}_배치정원".format(M_name[i])] = int(input("{} 배치정원 : ".format(M_name[i])))
    




# %% 5. 전공별 배치 완료자 명단(새 리스트) 생성-----------------------------------------------------------------------------
for i in range(M_count):
    globals()["{}_가전공_배치".format(M_name[i])] = []





# %% 6. 전공별 1~n지망 배치 완료자 명단(새 리스트) 생성------------------------------------------------------------------------
for i in M_name:
    for j in range(M_count):
        globals()["{}_{}지망_배치".format(i, j+1)] = []





########################################################################################################################
# %% 7. 학과별 지망 인원수 확인----------------------------------------------------------------------------------------------
for j in range(M_count):
    for i in range(len(M)):
        if M.loc[i][1] == M_name[j]:
            globals()["{}_1지망".format(M_name[j])].append(M.loc[i][0])


# 전공별 1지망 지원자 수 출력
print("----- \t{} {} 신입생 가전공 수요조사 1지망 결과 요약\t -----".format(C_name ,F_name))
print("{} {} 신입생 가전공 수요조사 응답자는 총 {}명 입니다.\n".format(C_name ,F_name, len(M)))
for i in range(M_count):
    print("{}을 1지망으로 선택한 학우는 총 {}명 입니다.\n".format(M_name[i], len(globals()["{}_1지망".format(M_name[i])])))
    




# %% 8. 1지망 배치 후, 배치 결과 및 초과 학과 발표---------------------------------------------------------------------------
excess_1 = []     # 1지망 지원 결과 배치 정원을 넘어간 학과명을 저장하기 위한 리스트



for i in range(M_count):
    for j in globals()["{}_1지망".format(M_name[i])]:
        if len(globals()["{}_1지망".format(M_name[i])]) > globals()["{}_배치정원".format(M_name[i])]:
            print("***** {}_1지망_초과 *****\n".format(M_name[i]))
            excess_1.append(M_name[i])
            break
        else:
            globals()["{}_가전공_배치".format(M_name[i])].append(j)
            globals()["{}_1지망_배치".format(M_name[i])].append(j)
    if globals()["{}_가전공_배치".format(M_name[i])] != []:
        print("{} 1지망 배치 완료\n".format(M_name[i]))


# 1지망 미초과 인원 배치 결과
for i in range(M_count):
    print("1지망으로 {}에 배치된 인원".format(M_name[i]))
    print(globals()["{}_1지망_배치".format(M_name[i])])
    print("")





# %% 9. 1지망 초과 전공 ▶ 난수 생성 후 배정 및 배치 결과 발표----------------------------------------------------------------
for i in range(M_count):
    if globals()["{}_가전공_배치".format(M_name[i])] == []:
        random_number = random.sample(range(0, len(globals()["{}_1지망".format(M_name[i])])), globals()["{}_배치정원".format(M_name[i])])
        random_number.sort()
        for j in random_number:
            globals()["{}_가전공_배치".format(M_name[i])].append(globals()["{}_1지망".format(M_name[i])][j])
            globals()["{}_1지망_배치".format(M_name[i])].append(globals()["{}_1지망".format(M_name[i])][j])
    print("1지망에서 {}에 배치된 인원".format(M_name[i]))
    print(globals()["{}_1지망_배치".format(M_name[i])])
    print("")


# ▶ 1지망 가전공 배치 완료
for i in range(M_count):
    print("-----\t{}_1지망 배치 결과\t-----".format(M_name[i]))
    print(globals()["{}_1지망_배치".format(M_name[i])])
    print("")





# %% 10. 가전공 미배치(1지망 탈락) 인원 ▶ 2지망 인원수 및 잔여 배정 가능 인원--------------------------------------------------
# 초과된 전공에 대해서 1지망 배정이 된 인원들을 해당 전공 1지망 지원 명단에서 제외
for i in range(len(excess_1)):
    for j in globals()["{}_1지망_배치".format(excess_1[i])]:
        globals()["{}_1지망".format(excess_1[i])].remove(j)


# 전공별 가전공 배치된 사람들의 인덱스를 추출할 리스트 생성
for i in range(M_count):
    globals()["{}_가전공_배치_인덱스".format(M_name[i])] = []


# 전체 명단에서 이미 가전공 배치된 인원들의 인덱스를 전공별로 추출
for i in range(M_count):
    for j in range(len(M)):
        for k in range(len(globals()["{}_가전공_배치".format(M_name[i])])):
            if M.loc[j][0] == globals()["{}_가전공_배치".format(M_name[i])][k]:
                globals()["{}_가전공_배치_인덱스".format(M_name[i])].append(j)


# 제거할 인덱스 번호 결합
delete = []
for i in range(M_count):
    delete += globals()["{}_가전공_배치_인덱스".format(M_name[i])]

delete.sort()



# 1지망 배치가 끝난 가전공 미배치 인원 리스트 정리
delete_1 = M.drop(delete, axis=0)
delete_1 = delete_1.reset_index()
delete_1 = delete_1.drop('1지망', axis=1)

delete_1
########################################################################################################################




# %% 11. 7~10 반복(모든 지망 인원 배치 끝날 때까지)---------------------------------------------------------------------------
after_1 = M_name.copy()

for i in range(len(excess_1)):
    after_1.remove(excess_1[i])

# 가전공 배치 미완료 전공
after_1


# 2지망 학과별 지원자 수 확인
for i in range(len(after_1)):
    globals()["{}_2지망".format(after_1[i])] = []

for i in range(len(delete_1)):
    for j in range(len(after_1)):

        if delete_1.loc[i][2] == str(after_1[j]):
            globals()["{}_2지망".format(after_1[j])].append(delete_1.loc[i][1])


# 전공별 2지망 지원자 수 출력
print("----- \t{} {} 신입생 가전공 수요조사 2지망 결과 요약\t -----".format(C_name ,F_name))
for i in range(len(after_1)):
    print("{}을 2지망으로 선택한 학우는 총 {}명 입니다.\n".format(after_1[i], len(globals()["{}_2지망".format(after_1[i])])))


# 전공별 배정인원 및 잔여 배정 가능 인원 수 출력
print("----- \t{} {} 전공별 배정 현황 및 잔여 배정 가능 인원수\t -----".format(C_name ,F_name))
for i in range(M_count):
    print("{}에 배정된 인원 수 : {}명 \n남은 배정인원 수 : {}명\n".format(M_name[i], len(globals()["{}_가전공_배치_인덱스".format(M_name[i])]), globals()["{}_배치정원".format(M_name[i])]-len(globals()["{}_가전공_배치_인덱스".format(M_name[i])])))


# 전공별 잔여 정원 계산
for i in range(len(after_1)):
    globals()["{}_잔여정원".format(after_1[i])] = globals()["{}_배치정원".format(after_1[i])]-len(globals()["{}_가전공_배치".format(after_1[i])])


# 2지망 배치 후, 배치 결과 및 초과 학과 발표
excess_2 = []   # 2지망 배치에서 초과된 학과명을 담기 위한 리스트

for i in range(len(after_1)):
    for j in globals()["{}_2지망".format(after_1[i])]:
        if len(globals()["{}_2지망".format(after_1[i])]) > globals()["{}_배치정원".format(after_1[i])]:
            print("***** {}_2지망_초과 *****\n".format(after_1[i]))
            excess_2.append(after_1[i])
            break
        else:
            globals()["{}_가전공_배치".format(after_1[i])].append(j)
            globals()["{}_2지망_배치".format(after_1[i])].append(j)
    if globals()["{}_가전공_배치".format(after_1[i])] != []:
        print("{} 2지망 배치 완료\n".format(after_1[i]))


# 2지망 미초과 인원 배치 결과
for i in range(M_count):
    print("2지망으로 {}에 배치된 인원".format(M_name[i]))
    print(globals()["{}_2지망_배치".format(M_name[i])])
    print("")


# 2지망 초과 전공 ▶ 난수 생성 후 배정 및 배치 결과 발표----------------------------------------------------------------
for i in range(len(after_1)):
    if globals()["{}_가전공_배치".format(after_1[i])] == []:
        random_number = random.sample(range(0, len(globals()["{}_1지망".format(after_1[i])])), globals()["{}_배치정원".format(after_1[i])])
        random_number.sort()
        for j in random_number:
            globals()["{}_가전공_배치".format(after_1[i])].append(globals()["{}_2지망".format(after_1[i])][j])
            globals()["{}_2지망_배치".format(after_1[i])].append(globals()["{}_2지망".format(after_1[i])][j])
    print("2지망에서 {}에 배치된 인원".format(after_1[i]))
    print(globals()["{}_2지망_배치".format(after_1[i])])
    print("")


# ▶ 2지망 가전공 배치 완료
for i in range(len(after_1)):
    print("-----\t{}_2지망 배치 결과\t-----".format(after_1[i]))
    print(globals()["{}_2지망_배치".format(after_1[i])])
    print("")
    
    
    
# 가전공 미배치(2지망 탈락) 인원 ▶ 3지망 인원수 및 잔여 배정 가능 인원--------------------------------------------------
# 초과된 전공에 대해서 2지망 배정이 된 인원들을 해당 전공 2지망 지원 명단에서 제외
for i in range(len(excess_2)):
    for j in globals()["{}_2지망_배치".format(excess_2[i])]:
        globals()["{}_2지망".format(excess_2[i])].remove(j)


# 전공별 가전공 배치된 사람들의 인덱스를 추출할 리스트 생성
for i in range(len(after_1)):
    globals()["{}_가전공_배치_인덱스".format(after_1[i])] = []


# 전체 명단에서 이미 가전공 배치된 인원들의 인덱스를 전공별로 추출
for i in range(len(after_1)):
    for j in range(len(M)):
        for k in range(len(globals()["{}_가전공_배치".format(after_1[i])])):
            if M.loc[j][0] == globals()["{}_가전공_배치".format(after_1[i])][k]:
                globals()["{}_가전공_배치_인덱스".format(after_1[i])].append(j)


# 제거할 인덱스 번호 결합
delete = []
for i in range(len(after_1)):
    delete += globals()["{}_가전공_배치_인덱스".format(after_1[i])]

delete.sort()



# 2지망 배치가 끝난 가전공 미배치 인원 리스트 정리
delete_2 = M.drop(delete, axis=0)
delete_2 = delete_2.reset_index()
delete_2 = delete_2.drop('2지망', axis=1)

delete_2




after_2 = M_name.copy()

for i in range(len(excess_2)):
    after_2.remove(excess_2[i])

# 가전공 배치 미완료 전공
after_2


# 2지망 학과별 지원자 수 확인
for i in range(len(after_2)):
    globals()["{}_3지망".format(after_2[i])] = []

for i in range(len(delete_2)):
    for j in range(len(after_2)):

        if delete_2.loc[i][3] == str(after_2[j]):
            globals()["{}_3지망".format(after_2[j])].append(delete_2.loc[i][1])


# 전공별 3지망 지원자 수 출력
print("----- \t{} {} 신입생 가전공 수요조사 3지망 결과 요약\t -----".format(C_name ,F_name))
for i in range(len(after_2)):
    print("{}을 3지망으로 선택한 학우는 총 {}명 입니다.\n".format(after_2[i], len(globals()["{}_3지망".format(after_2[i])])))


# 전공별 배정인원 및 잔여 배정 가능 인원 수 출력
print("----- \t{} {} 전공별 배정 현황 및 잔여 배정 가능 인원수\t -----".format(C_name ,F_name))
for i in range(M_count):
    print("{}에 배정된 인원 수 : {}명 \n남은 배정인원 수 : {}명\n".format(M_name[i], len(globals()["{}_가전공_배치_인덱스".format(M_name[i])]), globals()["{}_배치정원".format(M_name[i])]-len(globals()["{}_가전공_배치_인덱스".format(M_name[i])])))


# 전공별 잔여 정원 계산
for i in range(len(after_2)):
    globals()["{}_잔여정원".format(after_2[i])] = globals()["{}_배치정원".format(after_2[i])]-len(globals()["{}_가전공_배치".format(after_2[i])])


# 3지망 배치 후, 배치 결과 및 초과 학과 발표
excess_3 = []   # 3지망 배치에서 초과된 학과명을 담기 위한 리스트

for i in range(len(after_2)):
    for j in globals()["{}_3지망".format(after_2[i])]:
        if len(globals()["{}_3지망".format(after_2[i])]) > globals()["{}_배치정원".format(after_2[i])]:
            print("***** {}_3지망_초과 *****\n".format(after_2[i]))
            excess_3.append(after_2[i])
            break
        else:
            globals()["{}_가전공_배치".format(after_2[i])].append(j)
            globals()["{}_3지망_배치".format(after_2[i])].append(j)
    if globals()["{}_가전공_배치".format(after_2[i])] != []:
        print("{} 3지망 배치 완료\n".format(after_2[i]))


# 3지망 미초과 인원 배치 결과
for i in range(M_count):
    print("3지망으로 {}에 배치된 인원".format(M_name[i]))
    print(globals()["{}_3지망_배치".format(M_name[i])])
    print("")


# 3지망 초과 전공 ▶ 난수 생성 후 배정 및 배치 결과 발표----------------------------------------------------------------
for i in range(len(after_2)):
    if globals()["{}_가전공_배치".format(after_2[i])] == []:
        random_number = random.sample(range(0, len(globals()["{}_2지망".format(after_2[i])])), globals()["{}_배치정원".format(after_2[i])])
        random_number.sort()
        for j in random_number:
            globals()["{}_가전공_배치".format(after_2[i])].append(globals()["{}_3지망".format(after_2[i])][j])
            globals()["{}_3지망_배치".format(after_2[i])].append(globals()["{}_3지망".format(after_2[i])][j])
    print("3지망에서 {}에 배치된 인원".format(after_2[i]))
    print(globals()["{}_3지망_배치".format(after_2[i])])
    print("")


# ▶ 3지망 가전공 배치 완료
for i in range(len(after_2)):
    print("-----\t{}_3지망 배치 결과\t-----".format(after_2[i]))
    print(globals()["{}_3지망_배치".format(after_2[i])])
    print("")
    
    
    
# 가전공 미배치(3지망 탈락) 인원 ▶ 3지망 인원수 및 잔여 배정 가능 인원--------------------------------------------------
# 초과된 전공에 대해서 3지망 배정이 된 인원들을 해당 전공 3지망 지원 명단에서 제외
for i in range(len(excess_3)):
    for j in globals()["{}_3지망_배치".format(excess_3[i])]:
        globals()["{}_3지망".format(excess_3[i])].remove(j)


# 전공별 가전공 배치된 사람들의 인덱스를 추출할 리스트 생성
for i in range(len(after_2)):
    globals()["{}_가전공_배치_인덱스".format(after_2[i])] = []


# 전체 명단에서 이미 가전공 배치된 인원들의 인덱스를 전공별로 추출
for i in range(len(after_2)):
    for j in range(len(M)):
        for k in range(len(globals()["{}_가전공_배치".format(after_2[i])])):
            if M.loc[j][0] == globals()["{}_가전공_배치".format(after_2[i])][k]:
                globals()["{}_가전공_배치_인덱스".format(after_2[i])].append(j)


# 제거할 인덱스 번호 결합
delete = []
for i in range(len(after_2)):
    delete += globals()["{}_가전공_배치_인덱스".format(after_2[i])]

delete.sort()



# 3지망 배치가 끝난 가전공 미배치 인원 리스트 정리
delete_3 = M.drop(delete, axis=0)
delete_3 = delete_3.reset_index()
delete_3 = delete_3.drop('3지망', axis=1)

delete_3





after_3 = M_name.copy()

for i in range(len(excess_3)):
    after_3.remove(excess_3[i])

# 가전공 배치 미완료 전공
after_3


# 3지망 학과별 지원자 수 확인
for i in range(len(after_3)):
    globals()["{}_4지망".format(after_3[i])] = []

for i in range(len(delete_3)):
    for j in range(len(after_3)):

        if delete_3.loc[i][4] == str(after_3[j]):
            globals()["{}_4지망".format(after_3[j])].append(delete_3.loc[i][1])


# 전공별 4지망 지원자 수 출력
print("----- \t{} {} 신입생 가전공 수요조사 4지망 결과 요약\t -----".format(C_name ,F_name))
for i in range(len(after_3)):
    print("{}을 4지망으로 선택한 학우는 총 {}명 입니다.\n".format(after_3[i], len(globals()["{}_4지망".format(after_3[i])])))


# 전공별 배정인원 및 잔여 배정 가능 인원 수 출력
print("----- \t{} {} 전공별 배정 현황 및 잔여 배정 가능 인원수\t -----".format(C_name ,F_name))
for i in range(M_count):
    print("{}에 배정된 인원 수 : {}명 \n남은 배정인원 수 : {}명\n".format(M_name[i], len(globals()["{}_가전공_배치_인덱스".format(M_name[i])]), globals()["{}_배치정원".format(M_name[i])]-len(globals()["{}_가전공_배치_인덱스".format(M_name[i])])))


# 전공별 잔여 정원 계산
for i in range(len(after_3)):
    globals()["{}_잔여정원".format(after_3[i])] = globals()["{}_배치정원".format(after_3[i])]-len(globals()["{}_가전공_배치".format(after_3[i])])


# 4지망 배치 후, 배치 결과 및 초과 학과 발표
excess_4 = []   # 4지망 배치에서 초과된 학과명을 담기 위한 리스트

for i in range(len(after_3)):
    for j in globals()["{}_4지망".format(after_3[i])]:
        if len(globals()["{}_4지망".format(after_3[i])]) > globals()["{}_배치정원".format(after_3[i])]:
            print("***** {}_4지망_초과 *****\n".format(after_3[i]))
            excess_4.append(after_3[i])
            break
        else:
            globals()["{}_가전공_배치".format(after_3[i])].append(j)
            globals()["{}_4지망_배치".format(after_3[i])].append(j)
    if globals()["{}_가전공_배치".format(after_3[i])] != []:
        print("{} 4지망 배치 완료\n".format(after_3[i]))


# 4지망 미초과 인원 배치 결과
for i in range(M_count):
    print("4지망으로 {}에 배치된 인원".format(M_name[i]))
    print(globals()["{}_4지망_배치".format(M_name[i])])
    print("")


# 4지망 초과 전공 ▶ 난수 생성 후 배정 및 배치 결과 발표----------------------------------------------------------------
for i in range(len(after_3)):
    if globals()["{}_가전공_배치".format(after_3[i])] == []:
        random_number = random.sample(range(0, len(globals()["{}_3지망".format(after_3[i])])), globals()["{}_배치정원".format(after_3[i])])
        random_number.sort()
        for j in random_number:
            globals()["{}_가전공_배치".format(after_3[i])].append(globals()["{}_4지망".format(after_3[i])][j])
            globals()["{}_4지망_배치".format(after_3[i])].append(globals()["{}_4지망".format(after_3[i])][j])
    print("4지망에서 {}에 배치된 인원".format(after_3[i]))
    print(globals()["{}_4지망_배치".format(after_3[i])])
    print("")


# ▶ 4지망 가전공 배치 완료
for i in range(len(after_3)):
    print("-----\t{}_4지망 배치 결과\t-----".format(after_3[i]))
    print(globals()["{}_4지망_배치".format(after_3[i])])
    print("")
    
    
    
# 가전공 미배치(4지망 탈락) 인원 ▶ 4지망 인원수 및 잔여 배정 가능 인원--------------------------------------------------
# 초과된 전공에 대해서 4지망 배정이 된 인원들을 해당 전공 4지망 지원 명단에서 제외
for i in range(len(excess_4)):
    for j in globals()["{}_4지망_배치".format(excess_4[i])]:
        globals()["{}_4지망".format(excess_4[i])].remove(j)


# 전공별 가전공 배치된 사람들의 인덱스를 추출할 리스트 생성
for i in range(len(after_3)):
    globals()["{}_가전공_배치_인덱스".format(after_3[i])] = []


# 전체 명단에서 이미 가전공 배치된 인원들의 인덱스를 전공별로 추출
for i in range(len(after_3)):
    for j in range(len(M)):
        for k in range(len(globals()["{}_가전공_배치".format(after_3[i])])):
            if M.loc[j][0] == globals()["{}_가전공_배치".format(after_3[i])][k]:
                globals()["{}_가전공_배치_인덱스".format(after_3[i])].append(j)


# 제거할 인덱스 번호 결합
delete = []
for i in range(len(after_3)):
    delete += globals()["{}_가전공_배치_인덱스".format(after_3[i])]

delete.sort()



# 3지망 배치가 끝난 가전공 미배치 인원 리스트 정리
delete_4 = M.drop(delete, axis=0)
delete_4 = delete_4.reset_index()
delete_4 = delete_4.drop('4지망', axis=1)

delete_4




# %% 12. 전공 미배치 인원(데이터 중복 검사)
unassigned_students = delete_4  # 전공 미배치 인원 명단
unassigned_duplicates = unassigned_students[unassigned_students.duplicated()]
print("전공 미배치 인원 중 중복 데이터:")
print(unassigned_duplicates)

# %% 13. 최종 가전공 배치 결과 발표
for major in M_name:  # M_name은 전공명 리스트
    assigned_students = globals()[f"{major}_가전공_배치"]
    print(f"{major} 전공 배치 결과:")
    print(assigned_students)

# %% 14. 학부 및 가전공별 엑셀 파일 변환
for major in M_name:
    assigned_students_df = pd.DataFrame(globals()[f"{major}_가전공_배치"])
    assigned_students_df.to_excel(f"{major}_배치결과.xlsx", index=False)
