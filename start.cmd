@echo off
cd src\hwp2hwpx
mvn clean compile
if %errorlevel% neq 0 (
    echo Maven 빌드 실패
    exit /b %errorlevel%
)
cd ..\..
uv run uvicorn src.main:app --host 0.0.0.0 --port 80 --reload