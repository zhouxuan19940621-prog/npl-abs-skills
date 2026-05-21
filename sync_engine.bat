@echo off
REM 同步预测引擎到 GitHub
REM 用法: 双击运行，或在命令行执行 sync_engine.bat

set SOURCE=C:\Users\zhoux\PyCharmMiscProject\63_clean_v14_gs.py
set REPO=C:\Users\zhoux\npl-skills-repo
set TARGET=%REPO%\core-engine\63_clean_v14_gs.py

echo === ABS预测引擎 GitHub同步 ===

copy "%SOURCE%" "%TARGET%" >nul

cd /d "%REPO%"
git add core-engine/63_clean_v14_gs.py >nul

git diff --cached --quiet
if %errorlevel% equ 0 (
    echo 无变更，跳过推送。
) else (
    git commit -m "sync: update 63_clean_v14_gs.py"
    git push origin master
    echo 已推送到 GitHub。
)
