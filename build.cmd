setlocal

set SCRIPT_DIR=%~dp0
cd /d "%SCRIPT_DIR%"

echo hwp2hwpx JAR/의존성 준비...
cd src\hwp2hwpx
mvn clean package dependency:copy-dependencies -DskipTests -DoutputDirectory=target/lib -DincludeScope=runtime
if %errorlevel% neq 0 (
    echo Maven 패키징/의존성 복사 실패
    exit /b %errorlevel%
)
cd ..\..