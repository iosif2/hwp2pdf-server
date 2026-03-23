import logging
import importlib
import os
from functools import lru_cache
from pathlib import Path
from typing import Protocol, cast

try:
    import winreg
except ImportError:
    winreg = None

logger = logging.getLogger(__name__)

TEMP_DIR = "C:\\temp_hwp_files"

SECURITY_MODULE_NAME = os.getenv("HWP_SECURITY_MODULE_NAME", "FilePathCheckerModule")
SECURITY_MODULE_DLL_ENV = "HWP_SECURITY_MODULE_DLL"
SECURITY_MODULE_REGISTRY_PATH = r"SOFTWARE\HNC\HwpAutomation\Modules"
SECURITY_MODULE_FILENAMES = (
    "FilePathCheckerModule.dll",
    "FilePathCheckerModuleExample.dll",
)


class Win32GenCache(Protocol):
    def EnsureDispatch(self, prog_id: str) -> object: ...


class Win32ClientModule(Protocol):
    gencache: Win32GenCache


class PyHwpFactory(Protocol):
    def __call__(self, *, visible: bool = ...) -> "HwpClient": ...


class PythonComModule(Protocol):
    def CoInitialize(self) -> None: ...
    def CoUninitialize(self) -> None: ...


class HwpDocument(Protocol):
    def Close(self, *, isDirty: bool) -> None: ...


class HwpDocuments(Protocol):
    @property
    def Count(self) -> int: ...

    def Item(self, index: int) -> HwpDocument: ...


class HwpClient(Protocol):
    def RegisterModule(self, module_type: str, module_name: str) -> bool: ...
    def Open(self, path: str, file_format: str, option: str) -> None: ...
    def SaveAs(self, path: str, file_format: str) -> None: ...
    def Quit(self) -> None: ...
    def open(self, path: str) -> None: ...
    def save_as(self, path: str, file_format: str = ...) -> None: ...
    def quit(self) -> None: ...

    XHwpDocuments: HwpDocuments


def _load_pythoncom() -> PythonComModule:
    module = importlib.import_module("pythoncom")
    return cast(PythonComModule, cast(object, module))


def _load_win32com_client() -> Win32ClientModule:
    module = importlib.import_module("win32com.client")
    return cast(Win32ClientModule, cast(object, module))


def _load_pyhwpx_factory() -> PyHwpFactory:
    module = importlib.import_module("pyhwpx")
    return cast(PyHwpFactory, getattr(module, "Hwp"))


def _ensure_temp_dir() -> None:
    os.makedirs(TEMP_DIR, exist_ok=True)


@lru_cache(maxsize=1)
def _discover_security_module_path() -> str | None:
    configured_path = os.getenv(SECURITY_MODULE_DLL_ENV)
    candidate_paths: list[Path] = []

    if configured_path:
        candidate_paths.append(Path(configured_path))

    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        root = os.getenv(env_name)
        if not root:
            continue
        base_path = Path(root)
        for dll_name in SECURITY_MODULE_FILENAMES:
            candidate_paths.extend(base_path.glob(f"Hnc/**/{dll_name}"))

    checked: set[str] = set()
    for candidate in candidate_paths:
        normalized = str(candidate)
        if normalized in checked:
            continue
        checked.add(normalized)
        if candidate.is_file():
            return normalized

    return None


def _ensure_security_module_registration() -> str | None:
    dll_path = _discover_security_module_path()
    if not dll_path:
        logger.warning(
            "보안 모듈 DLL을 찾지 못했습니다. 환경 변수 %s 또는 한글 설치 경로를 확인하세요.",
            SECURITY_MODULE_DLL_ENV,
        )
        return None

    if winreg is None:
        logger.warning(
            "winreg를 사용할 수 없어 보안 모듈 레지스트리를 등록하지 못했습니다."
        )
        return dll_path

    try:
        registry_key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER, SECURITY_MODULE_REGISTRY_PATH
        )
        winreg.SetValueEx(
            registry_key, SECURITY_MODULE_NAME, 0, winreg.REG_SZ, dll_path
        )
        winreg.CloseKey(registry_key)
        logger.info(
            "보안 모듈 레지스트리 등록 완료: %s -> %s", SECURITY_MODULE_NAME, dll_path
        )
    except OSError as exc:
        logger.warning("보안 모듈 레지스트리 등록 실패: %s", exc)

    return dll_path


def _create_hwp_client() -> HwpClient:
    try:
        hwp_factory = _load_pyhwpx_factory()

        try:
            hwp = hwp_factory(visible=False)
        except TypeError:
            hwp = hwp_factory()
        logger.info("pyhwpx를 사용해 한글 객체를 초기화했습니다.")
        return hwp
    except ImportError:
        logger.info("pyhwpx가 설치되어 있지 않아 win32com으로 대체합니다.")
        win32_client = _load_win32com_client()
        return cast(
            HwpClient,
            win32_client.gencache.EnsureDispatch("HWPFrame.HwpObject"),
        )


def _register_security_module(hwp: HwpClient) -> None:
    if not hasattr(hwp, "RegisterModule"):
        logger.info(
            "현재 HWP 객체는 RegisterModule 메서드를 직접 노출하지 않습니다. pyhwpx 기본 동작을 사용합니다."
        )
        return

    dll_path = _ensure_security_module_registration()
    register_targets = [SECURITY_MODULE_NAME]
    if dll_path:
        register_targets.append(dll_path)

    for target in register_targets:
        try:
            result = hwp.RegisterModule("FilePathCheckDLL", target)
            logger.info("RegisterModule 실행: target=%s result=%s", target, result)
            if result:
                return
        except Exception as exc:
            logger.warning("RegisterModule 실패: target=%s error=%s", target, exc)


def _open_document(hwp: HwpClient, hwp_file_path: str) -> None:
    if hasattr(hwp, "open"):
        hwp.open(hwp_file_path)
        return

    file_format = "HWPX" if hwp_file_path.lower().endswith(".hwpx") else "HWP"
    hwp.Open(hwp_file_path, file_format, "forceopen:true")


def _save_as_pdf(hwp: HwpClient, output_pdf_path: str) -> None:
    if hasattr(hwp, "save_as"):
        try:
            hwp.save_as(output_pdf_path, "PDF")
        except TypeError:
            hwp.save_as(output_pdf_path)
        return

    hwp.SaveAs(output_pdf_path, "PDF")


def _close_hwp(hwp: HwpClient) -> None:
    close_error: Exception | None = None

    if hasattr(hwp, "XHwpDocuments"):
        try:
            documents = cast(HwpDocuments, getattr(hwp, "XHwpDocuments"))
            document_count = documents.Count
            if document_count > 0:
                documents.Item(0).Close(isDirty=False)
        except Exception as exc:
            close_error = exc

    quit_method = getattr(hwp, "quit", None)
    if callable(quit_method):
        try:
            _ = quit_method()
            return
        except Exception as exc:
            close_error = exc

    quit_method = getattr(hwp, "Quit", None)
    if callable(quit_method):
        try:
            _ = quit_method()
            return
        except Exception as exc:
            close_error = exc

    if close_error is not None:
        logger.warning("한글 객체 종료 중 경고 발생: %s", close_error)


# HWP to PDF 변환 함수
def convert_hwp_to_pdf(hwp_file_path: str, output_pdf_path: str) -> bool:
    """
    HWP 파일을 PDF로 변환합니다.
    이 함수는 Windows 환경과 한컴오피스 설치가 필요합니다.
    """
    pythoncom = _load_pythoncom()
    pythoncom.CoInitialize()  # 현재 스레드에 대한 COM 라이브러리 초기화
    hwp = None  # Hwp 객체를 try 블록 외부에서 접근할 수 있도록 초기화

    try:
        _ensure_temp_dir()
        hwp = _create_hwp_client()
        _register_security_module(hwp)
        _open_document(hwp, hwp_file_path)
        _save_as_pdf(hwp, output_pdf_path)

        logger.info(f"성공적으로 PDF로 변환 완료: {output_pdf_path}")
        return True
    except Exception as e:
        # 변환 과정에서 발생하는 모든 예외를 로깅합니다.
        logger.error(f"HWP to PDF 변환 중 오류 발생: {e}")
        return False
    finally:
        # HWP 객체가 성공적으로 생성되었다면, 리소스 누수를 방지하기 위해 종료합니다.
        if hwp:
            _close_hwp(hwp)
        # COM 라이브러리 사용을 종료합니다.
        pythoncom.CoUninitialize()
