@echo off
setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

if not exist src\hwp2hwpx\target\hwp2hwpx-1.0.0.jar (
    echo hwp2hwpx JAR이 없어 빌드합니다...
    cd src\hwp2hwpx
    mvn clean package -DskipTests
    if %errorlevel% neq 0 (
        echo Maven 패키징 실패
        exit /b %errorlevel%
    )
    cd ..\..
) else (
    echo hwp2hwpx JAR 발견: src\hwp2hwpx\target\hwp2hwpx-1.0.0.jar
)

uv run uvicorn src.main:app --host 0.0.0.0 --port 80 --reload