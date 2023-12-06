import traceback
from work import work


class Main:

    def start(self):
        try:
            # Api 에서 가져온 json 값
            # call_api_result = {}
            # 결과 dictionary
            result_dict = {}
            # 처리 전 버퍼
            result_buffer = []
            # 유실된 데이터
            lost_dict = {}

            wk = work()
            # 인자 값 받기
            api_url, secret_key, inputPath, outputPath = wk.getargs()
            
            wk.checkImageSize(inputPath)
            

            # Release 일때, API 호출
            call_api_result = wk.call_ocr_api(inputPath, secret_key, api_url)
            
            # 디버그 일때, 저장해둔 json 값으로 테스트
            # call_api_result = wk.getTestJson('.json')

            # 데이터 가져오기
            result_buffer = wk.data_grab(call_api_result)

            # 데이터 전처리 및 편집
            result_dict, lost_dict = wk.relocate_data(result_buffer)            

            # 글자 위치 그대로 액셀 작성
            wk.make_excel(result_dict, lost_dict, outputPath)            

            print("정상적으로 모든 작업을 완료하였습니다.")
        except Exception as e:
            print("작업을 실패했습니다, 에러를 확인해주세요.")
            print(e)
            print(traceback.print_exc())


if __name__ == '__main__':
    program = Main()
    program.start()