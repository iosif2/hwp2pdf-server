from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask
import os
import uuid
import logging
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor
import asyncio
import subprocess
from pathlib import Path

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


def convert_hwp_to_hwpx_using_maven(input_hwp: str, output_hwpx: str) -> bool:
    try:
        current_file = Path(__file__).resolve()
        hwp2hwpx_dir = current_file.parent / "hwp2hwpx"

        if not hwp2hwpx_dir.exists():
            logger.error(f"hwp2hwpx 디렉토리를 찾을 수 없습니다: {hwp2hwpx_dir}")
            return False

        jar_path = hwp2hwpx_dir / "target" / "hwp2hwpx-1.0.0.jar"
        lib_path = hwp2hwpx_dir / "target" / "lib"

        if not jar_path.exists():
            logger.error(f"패키징된 JAR을 찾을 수 없습니다: {jar_path}")
            return False

        # 의존성 JAR 존재 여부 확인
        if not lib_path.exists():
            logger.error(f"의존성 폴더(target/lib)가 없습니다. mvn dependency:copy-dependencies 실행 필요: {lib_path}")
            return False
        has_dep = any(lib_path.glob("*.jar"))
        if not has_dep:
            logger.error(f"target/lib에 의존성 JAR이 없습니다. mvn dependency:copy-dependencies 실행 필요: {lib_path}")
            return False

        classpath = f"{jar_path}{os.pathsep}{lib_path}/*"
        cmd = [
            "java",
            "-cp",
            classpath,
            "kr.dogfoot.hwp2hwpx.ConvertExample",
            input_hwp,
            output_hwpx,
        ]

        logger.info(f"JAR을 사용하여 HWP를 HWPX로 변환 중: {input_hwp} -> {output_hwpx}")
        result = subprocess.run(
            cmd,
            cwd=str(hwp2hwpx_dir),
            capture_output=True,
            text=True,
            check=False,
        )

        if result.returncode == 0 and os.path.exists(output_hwpx):
            logger.info(f"JAR 변환 성공: {output_hwpx}")
            return True
        else:
            logger.error(f"JAR 변환 실패 (returncode: {result.returncode})")
            if result.stdout:
                logger.error(f"JAR stdout: {result.stdout}")
            if result.stderr:
                logger.error(f"JAR stderr: {result.stderr}")
            return False

    except FileNotFoundError:
        logger.error("java 실행 파일을 찾을 수 없습니다.")
        return False
    except Exception as e:
        logger.error(f"JAR 변환 중 오류 발생: {e}")
        return False


async def convert_with_retry(
    executor: ThreadPoolExecutor, 
    input_path: str, 
    output_path: str,
    original_ext: str
) -> tuple[bool, str]:
    success = await convert_hwp_to_pdf_async_wrapper(executor, input_path, output_path)
    if success and os.path.exists(output_path):
        return True, input_path
    if success and not os.path.exists(output_path):
        logger.error(f"PDF가 생성되지 않았습니다: {output_path}")

    logger.warning(f"첫 번째 변환 시도 실패: {input_path}")

    if original_ext.lower() == '.hwp':
        hwpx_path = os.path.splitext(input_path)[0] + '.hwpx'
        if os.path.exists(hwpx_path):
            logger.info(f"확장자를 .hwpx로 변경하여 재시도: {hwpx_path}")
            success = await convert_hwp_to_pdf_async_wrapper(executor, hwpx_path, output_path)
            if success and os.path.exists(output_path):
                return True, hwpx_path
            if success:
                logger.error(f"PDF가 생성되지 않았습니다: {output_path}")
            logger.warning(f"확장자 변경 재시도 실패: {hwpx_path}")
        else:
            logger.info(f"확장자를 .hwpx로 변경한 파일이 존재하지 않음: {hwpx_path}")

        converted_hwpx_path = os.path.join(TEMP_DIR, f"{uuid.uuid4()}.hwpx")
        logger.info(f"Maven을 사용하여 HWP를 HWPX로 변환 후 재시도")
        if convert_hwp_to_hwpx_using_maven(input_path, converted_hwpx_path):
            success = await convert_hwp_to_pdf_async_wrapper(
                executor, converted_hwpx_path, output_path
            )
            if success and os.path.exists(output_path):
                return True, converted_hwpx_path
            if success:
                logger.error(f"PDF가 생성되지 않았습니다: {output_path}")
            else:
                cleanup_files(converted_hwpx_path)
        else:
            logger.error(f"Maven HWP->HWPX 변환 실패: {input_path}")

    elif original_ext.lower() == '.hwpx':
        hwp_path = os.path.splitext(input_path)[0] + '.hwp'
        if os.path.exists(hwp_path):
            logger.info(f"확장자를 .hwp로 변경하여 재시도: {hwp_path}")
            success = await convert_hwp_to_pdf_async_wrapper(executor, hwp_path, output_path)
            if success and os.path.exists(output_path):
                return True, hwp_path
            if success:
                logger.error(f"PDF가 생성되지 않았습니다: {output_path}")
            logger.warning(f"확장자 변경 재시도 실패: {hwp_path}")
        else:
            logger.info(f"확장자를 .hwp로 변경한 파일이 존재하지 않음: {hwp_path}")

    return False, input_path


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

    success, final_input_path = await convert_with_retry(
        request.app.state.executor, input_path, output_path, file_ext
    )

    if not success:
        cleanup_files(input_path)
        if final_input_path != input_path and os.path.exists(final_input_path):
            cleanup_files(final_input_path)
        raise HTTPException(
            status_code=500, detail="HWP 파일을 PDF로 변환하는데 실패했습니다."
        )

    response_filename = f"{os.path.splitext(file.filename)[0]}.pdf"

    files_to_cleanup = [input_path, output_path]
    if final_input_path != input_path and os.path.exists(final_input_path):
        files_to_cleanup.append(final_input_path)

    return FileResponse(
        output_path,
        media_type="application/pdf",
        filename=response_filename,
        background=BackgroundTask(cleanup_files, *files_to_cleanup),
    )
