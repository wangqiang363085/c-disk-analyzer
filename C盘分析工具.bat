@echo off
chcp 936 >nul
title C盘空间分析工具
setlocal enabledelayedexpansion

set desktop=%USERPROFILE%\Desktop
set report=%desktop%\C盘分析报告.txt
set csvfile=%desktop%\C盘目录说明.csv

color 0F
echo ============================================
echo       C 盘空间分析工具 (兼容 XP~Win11)
echo       正在扫描，请稍候...
echo ============================================
echo.

REM ===== 获取磁盘容量 =====
echo 获取磁盘信息...
for /f "tokens=2 delims==" %%a in ('wmic logicaldisk where "DeviceID='C:'" get Size /value 2^>nul') do set total=%%a
for /f "tokens=2 delims==" %%a in ('wmic logicaldisk where "DeviceID='C:'" get FreeSpace /value 2^>nul') do set free=%%a

if "%total%"=="" (
    echo ⚠ 无法获取磁盘信息，尝试备用方案...
    REM XP 备用方案
    for /f "tokens=3" %%a in ('dir C:\ 2^>nul ^| find "可用字节"') do set freetxt=%%a
    REM 总容量用fsutil
    for /f "tokens=2 delims=:" %%a in ('fsutil volume diskfree C: 2^>nul ^| find "总数"') do set totaltxt=%%a
)

set /a used=total-free
set /a total_gb=total/1000000000
set /a used_gb=used/1000000000
set /a free_gb=free/1000000000

if %total_gb% leq 0 (
    echo ⚠ 无法获取容量信息，请以管理员身份运行
    echo.
    pause
    exit /b
)

set /a used_pct=used*100/total

echo 总容量: %total_gb% GB
echo 已用:   %used_gb% GB (%used_pct%%%)
echo 可用:   %free_gb% GB

if %free_gb% lss 20 echo ⚠ 状态: 严重不足!
if %free_gb% geq 20 if %free_gb% lss 50 echo ⚡ 状态: 空间紧张
if %free_gb% geq 50 echo ✅ 状态: 正常
echo.

REM ===== 扫描根目录下的大文件夹 =====
echo 正在扫描一级目录...
echo.
echo 排名 大小         占比  目录名
echo ----------------------------------------

set tempfile=%TEMP%\dirlist.txt
set outfile=%TEMP%\dirsizes.txt

REM 清空临时文件
if exist %tempfile% del %tempfile%
if exist %outfile% del %outfile%

REM 扫描每个一级目录（跳过系统保护目录）
for /d %%i in (C:\*) do (
    set dirname=%%~nxi
    set skip=0
    if /i "!dirname!"=="System Volume Information" set skip=1
    if /i "!dirname!"=="$Recycle.Bin" set skip=1
    if /i "!dirname!"=="Recovery" set skip=1
    if /i "!dirname!"=="Windows.old" set skip=1

    if !skip!==0 (
        cls
        echo 正在扫描: %%i ...
        REM 用 dir 计算大小（最底层的兼容方式，XP 到 Win11 都支持）
        set size=0
        for /f "tokens=3 delims= " %%s in ('dir "%%i" /s /-c 2^>nul ^| findstr /i "文件" ^| findstr /v "目录"') do (
            set size=%%s
        )
        REM 去掉逗号
        set size=!size:,=!
        if defined size (
            echo %%i !size!>>%tempfile%
        )
    )
)

REM 排序（冒泡法，兼容所有版本）
cls
echo 正在排序...
set idx=0
for /f "tokens=1,* delims= " %%a in (%tempfile%) do (
    set dir[!idx!]=%%a
    set size[!idx!]=%%b
    set /a idx+=1
)

REM 开始输出
for /l %%i in (1,1,15) do set top[%%i]=0
set count=0
set total_scan=0

REM 简单排序取前15
set /a max_items=idx
for /l %%i in (0,1,%max_items%) do (
    if defined size[%%i] set /a total_scan+=!size[%%i]!
)

REM 输出到屏幕
set line=0
(for /f "tokens=1,* delims= " %%a in ('sort /r %tempfile%') do (
    set /a line+=1
    if !line! leq 15 (
        set /a pct=%%b*100/total_scan
        set s=%%b
        if !s! geq 1000000000 (
            set /a gb=s/1000000000
            call :printline !gb! GB  !pct!%%  %%~nxa
        ) else if !s! geq 1000000 (
            set /a mb=s/1000000
            call :printline !mb! MB  !pct!%%  %%~nxa
        ) else (
            call :printline !s! B  !pct!%%  %%~nxa
        )
    )
)) 2>nul

echo.
echo ============================================

REM ===== 扫描常用缓存目录 =====
echo.
echo --- 缓存目录检查 ---
set caches=(
    "C:\Windows\Temp"
    "C:\Windows\SoftwareDistribution\Download"
    "%TEMP%"
    "%USERPROFILE%\AppData\Local\Microsoft\Windows\INetCache"
)
if exist "%USERPROFILE%\AppData\Local\npm-cache" call :showcache "npm缓存" "%USERPROFILE%\AppData\Local\npm-cache"
if exist "%USERPROFILE%\AppData\Local\pip\cache" call :showcache "pip缓存" "%USERPROFILE%\AppData\Local\pip\cache"
call :showcache "Windows临时文件" "C:\Windows\Temp"
call :showcache "用户临时文件" "%TEMP%"
call :showcache "Windows更新缓存" "C:\Windows\SoftwareDistribution\Download"

REM ===== 写报告文件 =====
echo 正在生成报告...
(
echo ============================================
echo   C 盘空间分析报告
echo   生成时间: %DATE% %TIME%
echo   计算机:   %COMPUTERNAME%
echo ============================================
echo.
echo 总容量: %total_gb% GB
echo 已用:   %used_gb% GB (%used_pct%%%)
echo 可用:   %free_gb% GB
echo.
echo --- 一级目录大小排行 ---
) > %report%

REM 写入报告
set line=0
(for /f "tokens=1,* delims= " %%a in ('sort /r %tempfile%') do (
    set /a line+=1
    if !line! leq 20 (
        set /a pct=%%b*100/total_scan
        set s=%%b
        if !s! geq 1000000000 (
            set /a gb=s/1000000000
            echo !gb! GB  !pct!%%  %%~nxa>>%report%
        ) else if !s! geq 1000000 (
            set /a mb=s/1000000
            echo !mb! MB  !pct!%%  %%~nxa>>%report%
        ) else (
            echo !s! B  !pct!%%  %%~nxa>>%report%
        )
    )
)) 2>nul

REM ===== 生成 CSV =====
(
echo 目录路径,说明,能否删除
echo C:\Windows,Windows系统核心文件,不能
echo C:\Windows\Temp,Windows临时文件,能
echo C:\Windows\SoftwareDistribution\Download,Windows更新缓存,能
echo C:\Program Files,已安装程序(64位),不能
echo C:\Program Files (x86),已安装程序(32位),不能
echo C:\ProgramData,程序共享数据,不能
echo C:\Users,用户数据,视情况
echo C:\Users\%%USERNAME%%\AppData\Local\Temp,用户临时文件,能
echo C:\tmp,临时目录,能
echo hiberfil.sys,休眠文件(未启用),能
echo pagefile.sys,虚拟内存文件,不能
) > %csvfile%

REM 报告添加清理建议
(
echo.
echo --- 清理建议 ---
echo.
echo 一、能马上清理的:
echo   - 临时文件: cleanmgr 或手动删除 C:\Windows\Temp 和 %%TEMP%%
echo   - Windows 更新缓存: cleanmgr 或删除 SoftwareDistribution\Download
echo   - npm缓存: npm cache clean --force
echo   - pip缓存:  pip cache purge
echo.
echo 二、需在软件内清理:
echo   - 剪映/美图/WPS等软件: 在软件设置里清理缓存
echo   - 微信/钉钉/飞书: 设置里清理聊天缓存
echo.
echo 三、绝对不能碰:
echo   - C:\Windows 系统核心
echo   - C:\ProgramData 程序数据
echo   - C:\Program Files 通过设置卸载，勿手动删
echo   - pagefile.sys 虚拟内存
echo.
echo 四、注意事项:
echo   - 大文件放到 D 盘
echo   - 每月跑一次 cleanmgr
echo   - 视频软件做完后清缓存
echo   - 换大容量 SSD 解决根本问题
echo.
echo ============================================
) >> %report%

REM 算一下 pagefile.sys 和休眠文件
if exist C:\hiberfil.sys (
    for %%f in (C:\hiberfil.sys) do set hiber_size=%%~zf
    set /a hiber_gb=hiber_size/1000000000
    if !hiber_gb! gtr 0 echo 休眠文件: !hiber_gb! GB (可释放)>>%report%
)
if exist C:\pagefile.sys (
    for %%f in (C:\pagefile.sys) do set page_size=%%~zf
    set /a page_gb=page_size/1000000000
    if !page_gb! gtr 0 echo 虚拟内存: !page_gb! GB>>%report%
)

del %tempfile% 2>nul

cls
color 0A
echo ============================================
echo   ✅ 分析完成！
echo.
echo   报告文件: %report%
echo   CSV 速查表: %csvfile%
echo ============================================
echo.
echo 按任意键退出...
pause >nul
goto :eof

:printline
echo %*
goto :eof

:showcache
if exist "%~2" (
    set csize=0
    for /f "tokens=3 delims= " %%s in ('dir "%~2" /s /-c 2^>nul ^| findstr /i "文件" ^| findstr /v "目录"') do set csize=%%s
    set csize=!csize:,=!
    if defined csize if !csize! gtr 0 (
        set /a cmb=csize/1000000
        echo   %~1: !cmb! MB
    )
)
goto :eof
