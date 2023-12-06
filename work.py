import traceback
import argparse
import os.path
import openpyxl
import uuid
import json
import time
import requests
import re
import os
import shutil

class work:

    # 액셀 파일 만들기
    def make_excel(self, result_dict, lost_dict, outputPath):
        try:
            # 새로운 엑셀 워크북 생성
            workbook = openpyxl.Workbook()

            # 워크북의 활성시트(첫 번쨰 시트) 선택
            sheet = workbook.active
            i = 0
            j = 0

            # 실제 데이터
            for key in result_dict.keys():
                # Key 저장
                cell = sheet[str(chr(65 + j)) + str(i + 1)]
                # 셀에 저장될 값
                cell.value = key

                j += 1

                # value 저장
                for value in result_dict[key]:
                    # EX: A1, A2, B1,B2, etc..
                    cell = sheet[str(chr(65 + j)) + str(i + 1)]
                    # 셀에 저장될 값
                    cell.value = value
                    j += 1
                j = 0 if j % 2 == 0 else j
                i = i + 1 if j % 2 == 0 else i

            # 유실된 데이터가 있을 때만
            if len(lost_dict.keys()) != 0:
                #  i + 2 유실된값 라벨
                i +=2
                cell = sheet[str(chr(65 + j)) + str(i + 1)]
                # 셀에 저장될 값
                cell.value = "데이터 처리 과정에서 이상을 감지한 값들 입니다, 참고해주세요."
                i += 1
                cell = sheet[str(chr(65 + j)) + str(i + 1)]
                # 셀에 저장될 값
                cell.value = "원인: 1. 한칸에 두 줄,2. 빛,3. 회전,4. 화질,5. 필기,6. 구겨짐"
                i += 1
                cell = sheet[str(chr(65 + j)) + str(i + 1)]
                # 셀에 저장될 값
                cell.value = "앞뒤로 밀려 있거나 원래 빈칸 일 경우도 있습니다."
                # 3칸 아래부터 작성
                i += 1

                # 손실데이터
                for key in lost_dict.keys():
                    # Key 저장
                    cell = sheet[str(chr(65 + j)) + str(i + 1)]
                    # 셀에 저장될 값
                    cell.value = key

                    j += 1
                    # value 저장
                    for value in lost_dict[key]:
                        # EX: A1, A2, B1,B2, etc..
                        cell = sheet[str(chr(65 + j)) + str(i + 1)]
                        # 셀에 저장될 값
                        cell.value = value
                        j += 1
                    j = 0
                    i += 1
            workbook.save(outputPath)
            workbook.close()
            print("액셀 파일 생성완료")
        except Exception as e:
            print("작업을 실패했습니다, 에러를 확인해주세요.")
            print(e)
            print(traceback.print_exc())
            exit(1)

    # 문자열 위치 재배치
    def relocate_data(self, result_buffer):
        relocate_dict = {}
        # 유실된 dictionary
        lost_dict = {}
        try:
            buffer_list = []
            one_word = ""
            # 편집된 버퍼
            edited_buffer = " ".join(result_buffer).replace("~", "")

            # 숫자 위치 값
            num_position = re.findall("\\d+",edited_buffer)

            # 복사
            copied_num_position = num_position.copy()
            # 중복제거 및 정렬 된 숫자
            filtered_num = sorted(set(map(int, copied_num_position)))

            # 첫번째 열
            column_1 = int(num_position[0])
            # 두번째 열
            column_2 = int(num_position[1])
            # 첫번째 열 리스트
            column_list_1 = filtered_num[filtered_num.index(column_1):filtered_num.index(column_2)]
            # 두번째 열 리스트
            column_list_2 = filtered_num[filtered_num.index(column_2):]
            # merge 된 리스트
            merged_list = column_list_1.copy()

            # 단순히 숫자와 숫자 사이를 찾는 것은 어려웠습니다
            # 왜냐하면 중간에 OCR 이 잘못인식하여 0 같은 가비지 넘버를 만들어냈기 때문입니다.
            # 따라서 [0] [1] 번째 처음 시작 숫자를 체크하고,
            # 기준으로 정렬과 필터링을 거친뒤
            # 다시 merge 해주었습니다.
            i = 0
            index = 0
            while True:
                if len(column_list_2) == i:
                    break
                # 첫번째 열보다 두번쨰 열이 크다면 단순히 두번째 열 값들을 추가해줍니다.
                if len(column_list_1) < i:
                    merged_list.append(column_list_2[i])
                else:
                    index +=1
                    merged_list.insert(index, column_list_2[i])
                    index +=1
                i+=1

            final_num_position = list(map(str, merged_list))


            # 숫자가 무조건 처음에 나오기 때문에, 처음과 끝을 알 수 있음
            for z in range(len(final_num_position)):

                # ( 숫자 영어 한글 ) 한 그룹 단위

                #  마지막 위치 값이 아닐때
                if z != len(final_num_position) - 1:
                    start_point = work.findIndex(edited_buffer, final_num_position[z])
                    end_point = work.findIndex(edited_buffer, final_num_position[z + 1])
                    # 편집된 버퍼[현재 위치값: 다음 위치 값]
                    one_word = edited_buffer[start_point:end_point]
                else:
                    # 편집된 버퍼[마지막 위치값 : 끝]
                    one_word = edited_buffer[work.findIndex(edited_buffer, final_num_position[z]):]

                num = final_num_position[z]
                # 한 단어에서 한국어 위치값
                korea = re.findall(r'[가-힣]+', one_word)
                # 한 단어에서 영어 위치값
                english = re.findall(r'[a-zA-Z]+', one_word)

                buffer_list.append(" ".join(english))
                buffer_list.append(" ".join(korea))
                relocate_dict[num] = buffer_list
                buffer_list = []

                # 에러 처리
                if  len(korea) == 0 or len(english) == 0:
                    size = []
                    if len(korea) < len(english):
                        size = english
                    elif len(english) < len(korea):
                        size = korea
                    else:
                        size = korea

                    # 둘다 Null 일때
                    if len(korea) == 0 and len(english) == 0:
                        lost_dict[num] = ["Null", "Null"]
                    else:
                        for i in range(len(size)):
                            lost_eng = english[i] if i <= len(english)-1 else "Null"
                            lost_kor = korea[i] if i <= len(korea)-1 else "Null"
                            lost_dict[num] = [lost_eng, lost_kor]
            print("데이터 편집 정상")
        except Exception as e:
            print(e)
            print(traceback.print_exc())
            exit(1)
        return relocate_dict, lost_dict


    # 문자열에서 인덱스를 찾아주는 함수
    def findIndex(str_main, str_sub):
        index = str_main.find(str_sub)

        while True:
            # 마지막에 숫자가 나올때
            # 공백이 없으므로 바로 리턴
            if str_main[index:].find(" ") == -1:
                return index

            space = index + str_main[index:].find(" ")
            if str_main[index:space] == str_sub:
                return index
            else:
                index = str_main.find(str_sub, index + 1)

    # 경로에서 파일명 가져오기
    def get_filename_from_path(file_path):
        return os.path.basename(file_path)

    # 문자가 같은면 True , 틀리면 False , 영어는 소문자로 바꿔서 비교
    def find_lower_index(word, word_lower):
        if word_lower.lower() == word.lower():
            return True            
        return False

    # 데이터 편집 하는 함수
    def data_grab(self, call_api_result):
        iskorea = False
        isenglish = False
        isnumber = False

        # 찾은 필드 그룹 개수
        find_field_group_count = 0
        # 총 필드 그룹 수
        group_count = 0
        # 필드 다음 데이터 시작 여부
        next_data = False
                
        # 값 버퍼
        result_buffer = []
        
        text = ""
        # 숫자, 영어, 한국 한 그룹이 몇개씩 있는지 확인
        for field in call_api_result['images'][0]['fields']:
            # 실 문자열 값
            text = field['inferText']
            
            # 한글 field 찾기
            if not iskorea: iskorea = any(work.find_lower_index(text, word) for word in ["의미", "뜻", "meaning", "한글","korea"])
            # 숫자 field 찾기
            if not isnumber: isnumber = any(work.find_lower_index(text,word) for word in ["번호","num","숫자","number","no"])
            # 영어 field 찾기
            if not isenglish: isenglish = any(work.find_lower_index(text,word) for word in ["word","영어","단어","english"])

            if iskorea and isnumber and isenglish:
                group_count +=1
                iskorea = False
                isnumber = False
                isenglish = False
            
        for field in call_api_result['images'][0]['fields']:
            # 실 문자열 값
            text = field['inferText']
            
            # 한글 field 찾기
            if not iskorea: iskorea = any(work.find_lower_index(text, word) for word in ["의미", "뜻", "meaning", "한글","korea"])
            # 숫자 field 찾기
            if not isnumber: isnumber = any(work.find_lower_index(text,word) for word in ["번호","num","숫자","number","no"])
            # 영어 field 찾기
            if not isenglish: isenglish = any(work.find_lower_index(text,word) for word in ["word","영어","단어","english"])

            if iskorea and isnumber and isenglish:
                find_field_group_count +=1
                iskorea = False
                isnumber = False
                isenglish = False

            # 필드 다음 줄 부터 추가하기 위한 조건
            if find_field_group_count == group_count and not next_data:
                next_data = True            
                continue            
            # 실 데이터 추가
            if next_data:
                result_buffer.append(text)

        print("데이터 가공 정상")
        return result_buffer

    # 인자 받는 함수
    def getargs(self):
        # ArgumentParser 객체 생성
        parser = argparse.ArgumentParser(description='네이버 클로버 OCR을 이용한, 영어 단어 추출 프로그램입니다.')
        # 인수 추가
        parser.add_argument('--input', type=str, help='Input file path')
        parser.add_argument('--output', type=str, help='Output file path')
        parser.add_argument('--secret_key', type=str, help='OCR secretKey')
        parser.add_argument('--api_url', type=str, help='NAVER Clova API Url')

        # 명령행 인수 파싱
        args = parser.parse_args()

        inputPath = args.input
        outputPath = args.output
        secret_key = args.secret_key
        api_url = args.api_url

        for arg in args.__dict__:
            if args.__dict__[arg] is None:
                print(f"{arg} 인자를 받지 못했습니다.")
                exit(1)

        print("인자 정상적으로 받음")
        return api_url, secret_key, inputPath, outputPath


    # ocr api 호출 함수
    def call_ocr_api(self, inputPath, secret_key, api_url):

        paths = inputPath.split(",")
        files = [('file', open(path, 'rb')) for path in paths]
        request_json = {'images': [{'format': 'jpg',
                                    'name': 'demo'
                                    }],
                        'requestId': str(uuid.uuid4()),
                        'version': 'V2',
                        'timestamp': int(round(time.time() * 1000))
                        }
        payload = {'message': json.dumps(request_json).encode('UTF-8')}
        headers = {
            'X-OCR-SECRET': secret_key,
        }
        response = requests.request("POST", api_url, headers=headers, data=payload, files=files)
        print("OCR Api 정상 호출")
        return response.json()

    # 이미지 파일 크기 20MB 이하인지 확인
    def checkImageSize(self, inputPath):     
        fileSizeMB = [os.path.getsize(f"{path}") / (1024 * 1024) for path in inputPath.split(",")]  # 파일 사이즈
        for i in range(len(fileSizeMB)):    
            if re.search(r'\.(jpg|jpeg|png|gif|bmp)$', inputPath.split(",")[i], re.IGNORECASE):
                if fileSizeMB[i] > 20:
                    print("이미지 파일이 20MB 보다 큽니다.")
                    exit(1)
                else:
                    print("이미지 20MB 이하 정상")
    # TestData 가져오기
    def getTestJson(self, test_path):
        # if __debug__:
        file_path = os.path.abspath(test_path)  # JSON 파일 경로
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)            
            print("json 파일 읽기 정상")
            return data
        
    def copy_file(self, source_path, destination_path):
        try:
            shutil.copy(source_path, destination_path)
            print(f"{source_path}를 {destination_path}로 복사했습니다.")
        except FileNotFoundError:
            print("파일을 찾을 수 없습니다.")
            exit(1)
        except PermissionError:
            print("권한이 없습니다.")
            exit(1)
        except Exception as e:
            print(f"오류 발생: {e}")
            exit(1)