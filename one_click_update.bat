@echo off
chcp 65001 > nul & cls
cd /d %~dp0
setlocal enabledelayedexpansion

echo 🔍 正在检查 .gitignore 中是否屏蔽了 .xlsx 文件...
set found=
if exist ".gitignore" (
    for /f "delims=" %%i in ('findstr /R /C:"^\s*\*\.xlsx\s*$" .gitignore') do (
        set found=true
    )
)
if defined found (
    echo ❌ 警告：你的 .gitignore 中包含了 "*.xlsx"，这将导致 Excel 文件无法上传到 GitHub！
    echo ✅ 请打开 .gitignore 文件并删除或注释掉该行：  *.xlsx
    echo.
    pause
    goto :eof
) else (
    echo ✅ .gitignore 设置正常，未屏蔽 .xlsx 文件。
)

:: 先清空变量
set msg=
set filelist=

for /f "delims=" %%f in ('git status -s') do (
    set line=%%f
    set file=!line:~3!

    if /i "!file!"=="wangye.py" (
        set description=网页脚本
    ) else if /i "!file!"=="requirements.txt" (
        set description=依赖项
    ) else if /i "!file!"==".gitignore" (
        set description=Git忽略规则
    ) else if /i "!file:~-5!"==".xlsx" (
        if /i "!file!"=="school_data.xlsx" (
            set description=学校数据
        ) else if /i "!file!"=="招生专业.xlsx" (
            set description=招生专业
        ) else (
            set description=Excel文件
        )
    ) else (
        set description=!file!
    )

    echo !filelist! | findstr /C:"!description!" >nul
    if errorlevel 1 (
        set filelist=!filelist!!description!, 
    )
)

:: 这里用启用延迟扩展后再赋值给 %msg%
if defined filelist (
    setlocal enabledelayedexpansion
    set msg=更新: !filelist:~0,-2!
    endlocal & set msg=%msg%
) else (
    set msg=更新脚本和数据
)

echo.
echo 📁 正在添加所有更改的文件...
git add .

echo.
set /p usermsg=✏️ 请输入提交说明（留空将自动生成: %msg%）：
if not "%usermsg%"=="" set msg=%usermsg%

echo 🔁 正在提交更改...
git commit -m "%msg%"
git push

echo.
set /p open=是否要打开 Streamlit Cloud 查看结果？(Y/N)：
if /I "%open%"=="Y" (
    start https://streamlit.io/cloud
)

echo.
echo ✅ 一键更新完成，提交说明: %msg%
pause > nul