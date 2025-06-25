from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
import os
import uuid
import logging
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor
import asyncio

# HWP 변환 로직 임포트
from .util import convert_hwp_to_pdf, TEMP_DIR

logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)  # 로깅 설정


@asynccontextmanager
async def lifespan(app: FastAPI):
    # 애플리케이션 시작 시 1개의 워커를 가진 ThreadPoolExecutor 생성
    # COM 객체는 스레드에 안전하지 않으므로, 한 번에 하나의 변환만 처리하도록 보장합니다.
    executor = ThreadPoolExecutor(max_workers=1)
    app.state.executor = executor
    logger.info("ThreadPoolExecutor가 시작되었습니다.")
    yield
    # 애플리케이션 종료 시 Executor 종료
    logger.info("ThreadPoolExecutor를 종료합니다.")
    executor.shutdown(wait=True)


app = FastAPI(
    title="HWP to PDF Converter API",
    description="HWP 및 HWPX 파일을 PDF로 변환합니다.",
    version="1.1.0",
    lifespan=lifespan,
)


async def convert_hwp_to_pdf_async_wrapper(
    executor: ThreadPoolExecutor, hwp_path: str, pdf_path: str
):
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(executor, convert_hwp_to_pdf, hwp_path, pdf_path)


def cleanup_files(*file_paths):
    for path in file_paths:
        if os.path.exists(path):
            try:
                os.remove(path)
                logger.info(f"임시 파일을 정리했습니다: {path}")
            except Exception as e:
                logger.error(f"임시 파일 정리 실패 {path}: {e}")


@app.post("/convert/hwp-to-pdf")
async def convert_hwp(request: Request, file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith((".hwp", ".hwpx")):
        raise HTTPException(
            status_code=400, detail="HWP 또는 HWPX 파일만 업로드할 수 있습니다."
        )

    unique_id = str(uuid.uuid4())
    file_ext = os.path.splitext(file.filename)[1]
    input_path = os.path.join(TEMP_DIR, f"{unique_id}{file_ext}")
    output_path = os.path.join(TEMP_DIR, f"{unique_id}.pdf")

    try:
        contents = await file.read()
        with open(input_path, "wb") as buffer:
            buffer.write(contents)
        logger.info(f"임시 파일 저장 완료: {input_path}")
    except Exception as e:
        logger.error(f"파일 저장 실패: {e}")
        raise HTTPException(
            status_code=500, detail="업로드된 파일을 저장하는데 실패했습니다."
        )

    success = await convert_hwp_to_pdf_async_wrapper(
        request.app.state.executor, input_path, output_path
    )

    if not success:
        cleanup_files(input_path)
        raise HTTPException(
            status_code=500, detail="HWP 파일을 PDF로 변환하는데 실패했습니다."
        )

    response_filename = f"{os.path.splitext(file.filename)[0]}.pdf"

    return FileResponse(
        output_path,
        media_type="application/pdf",
        filename=response_filename,
        background=BackgroundTask(cleanup_files, input_path, output_path),
    )
