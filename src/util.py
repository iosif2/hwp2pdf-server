import os
import pythoncom
import win32com.client
from win32com.client import Dispatch
import logging

logger = logging.getLogger(__name__)

# 임시 파일 저장 경로 설정 (실제 경로로 변경 필요)
TEMP_DIR = "C:\\temp_hwp_files"
if not os.path.exists(TEMP_DIR):
    os.makedirs(TEMP_DIR)


# HWP to PDF 변환 함수
def convert_hwp_to_pdf(hwp_file_path: str, output_pdf_path: str) -> bool:
    """
    HWP 파일을 PDF로 변환합니다.
    이 함수는 Windows 환경과 한컴오피스 설치가 필요합니다.
    """
    pythoncom.CoInitialize()  # 현재 스레드에 대한 COM 라이브러리 초기화
    hwp = None  # Hwp 객체를 try 블록 외부에서 접근할 수 있도록 초기화

    try:
        
        #한글 파일을 열기 위해 HWP변수에 함수를 저장합니다.
        hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
        # 보안 경고 팝업이 뜨는 것을 방지합니다.
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        # HWP 파일을 엽니다.
        if ".hwpx" in hwp_file_path:
            hwp.Open(hwp_file_path, "HWPX", "forceopen:true")
        else:
            hwp.Open(hwp_file_path, "HWP", "forceopen:true")

        # HWP 파일을 PDF 포맷으로 저장합니다.
        hwp.SaveAs(output_pdf_path, "PDF")

        logger.info(f"성공적으로 PDF로 변환 완료: {output_pdf_path}")
        return True
    except Exception as e:
        # 변환 과정에서 발생하는 모든 예외를 로깅합니다.
        logger.error(f"HWP to PDF 변환 중 오류 발생: {e}")
        return False
    finally:
        # HWP 객체가 성공적으로 생성되었다면, 리소스 누수를 방지하기 위해 종료합니다.
        if hwp:
            hwp.XHwpDocuments.Item(0).Close(isDirty=False)
            hwp.Quit()
        # COM 라이브러리 사용을 종료합니다.
        pythoncom.CoUninitialize()
