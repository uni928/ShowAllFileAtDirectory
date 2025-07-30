@echo off
setlocal enabledelayedexpansion

:: �J�����g�f�B���N�g����CSV��
set "CURRENT_DIR=%CD%"
for %%I in (.) do set "FOLDER_NAME=%%~nxI"
set "BASE_CSV_NAME=%FOLDER_NAME%.csv"
set "CSV_NAME=%BASE_CSV_NAME%"
set "counter=1"

:checkFile
if exist "%CSV_NAME%" (
    set "CSV_NAME=%FOLDER_NAME%_%counter%.csv"
    set /a counter+=1
    goto checkFile
)

:: �w�b�_�[�s�iShift-JIS��echo�j
echo �g���q,����,�t�@�C����,�t�@�C���̃p�X,�T�C�Y,�쐬����,�ŏI�X�V����,�ŏI�A�N�Z�X����,���� > "%CSV_NAME%"

:: �t�@�C���ꗗ�����W
set "TEMP_FILE=~filelist.tmp"
del "%TEMP_FILE%" >nul 2>&1
(for /r %%F in (*) do echo %%F) > "%TEMP_FILE%"
sort "%TEMP_FILE%" > "%TEMP_FILE%.sorted"

:: �g���q���Ƃ̊i�[�p�ꎞ�t�@�C��
del ext_*.tmp >nul 2>&1

:: ���t�@�C�����̃J�E���g
set /a FILE_TOTAL=0
for /f "usebackq delims=" %%F in ("%TEMP_FILE%.sorted") do set /a FILE_TOTAL+=1
set /a FILE_DONE=0

:: �g���q���ƂɃu���b�N�o�́i�{���j
set "LAST_EXT="
set "EXT_COUNT=0"
setlocal EnableDelayedExpansion
set "GROUP_LINES="

for /f "usebackq delims=" %%F in ("%TEMP_FILE%.sorted") do (
    set "FILE=%%~fF"
    set "EXT=%%~xF"
    set "NAME=%%~nxF"
    set "FOLDER=%%~dpF"
    set "SIZE=%%~zF"

    REM wmic ���t�擾�i��O���̃G���[����j
    set "ESC_PATH=!FILE:\=\\!"
    set "CREATE_DATE="
    set "MODIFY_DATE="
    set "ACCESS_DATE="

    for /f "skip=1 tokens=1,2 delims==" %%a in ('wmic datafile where name^="!ESC_PATH!" get CreationDate /value 2^>nul') do if "%%a"=="CreationDate" set "CREATE_DATE=%%b"
    for /f "skip=1 tokens=1,2 delims==" %%a in ('wmic datafile where name^="!ESC_PATH!" get LastModified /value 2^>nul') do if "%%a"=="LastModified" set "MODIFY_DATE=%%b"
    for /f "skip=1 tokens=1,2 delims==" %%a in ('wmic datafile where name^="!ESC_PATH!" get LastAccessed /value 2^>nul') do if "%%a"=="LastAccessed" set "ACCESS_DATE=%%b"

    set "CREATE_DATE=!CREATE_DATE:~0,4!/!CREATE_DATE:~4,2!/!CREATE_DATE:~6,2! !CREATE_DATE:~8,2!:!CREATE_DATE:~10,2!:!CREATE_DATE:~12,2!"
    set "MODIFY_DATE=!MODIFY_DATE:~0,4!/!MODIFY_DATE:~4,2!/!MODIFY_DATE:~6,2! !MODIFY_DATE:~8,2!:!MODIFY_DATE:~10,2!:!MODIFY_DATE:~12,2!"
    set "ACCESS_DATE=!ACCESS_DATE:~0,4!/!ACCESS_DATE:~4,2!/!ACCESS_DATE:~6,2! !ACCESS_DATE:~8,2!:!ACCESS_DATE:~10,2!:!ACCESS_DATE:~12,2!"

    :: �����擾
    set "ATTR_STR="
    for /f "tokens=1" %%a in ('attrib "%%F"') do (
        echo %%a | findstr /C:"A" >nul && set "ATTR_STR=!ATTR_STR!A"
        echo %%a | findstr /C:"R" >nul && set "ATTR_STR=!ATTR_STR!R"
        echo %%a | findstr /C:"H" >nul && set "ATTR_STR=!ATTR_STR!H"
        echo %%a | findstr /C:"S" >nul && set "ATTR_STR=!ATTR_STR!S"
    )
    if "!ATTR_STR!"=="" set "ATTR_STR=NONE"

    :: �o�͍s��ۑ��i�g���q���L�[�Ƃ����ꎞ�t�@�C���ɒ~�ρj
    set "LINE=,,!NAME!,!REL_PATH!,!HUMAN_SIZE!,!CREATE_DATE!,!MODIFY_DATE!,!ACCESS_DATE!,!ATTR_STR!"
    echo !LINE!>> "ext_!EXT!.tmp"


    set /a EXT_COUNT+=1
    REM ���΃p�X�擾�iCURRENT_DIR�Ƃ̍����j
set "REL_PATH=!FILE:%CURRENT_DIR%\=!"

:: �T�C�Y�\�L�� B / KB / MB �ɕϊ�
set /a SIZE_B=!SIZE!
set "HUMAN_SIZE="
if !SIZE_B! LSS 1024 (
    set "HUMAN_SIZE=!SIZE_B! B"
) else if !SIZE_B! LSS 1048576 (
    set /a SIZE_KB_INT=!SIZE_B! / 1024
    set /a SIZE_KB_DEC=^!SIZE_B! %% 1024 * 10 / 1024
    set "HUMAN_SIZE=!SIZE_KB_INT!.!SIZE_KB_DEC! KB"
) else (
    set /a SIZE_MB_INT=!SIZE_B! / 1048576
    set /a SIZE_MB_DEC=^!SIZE_B! %% 1048576 * 100 / 1048576
    set /a SIZE_MB_DEC_1=!SIZE_MB_DEC! / 10
    set "HUMAN_SIZE=!SIZE_MB_INT!.!SIZE_MB_DEC_1! MB"
)

set "LINE=,,!NAME!,!REL_PATH!,!HUMAN_SIZE!,!CREATE_DATE!,!MODIFY_DATE!,!ACCESS_DATE!,!ATTR_STR!"
set "GROUP_LINES=!GROUP_LINES! "!LINE!""

:: �i���\���i���s�t���j
set /a FILE_DONE+=1
set /a PROGRESS=100 * FILE_DONE / FILE_TOTAL
echo �i��: !PROGRESS!%% (!FILE_DONE!/!FILE_TOTAL!)
)
echo.

:: �o�́i�g���q���ɂ܂Ƃ߂ďo�́A�t�@�C�����Ń\�[�g�j
for %%E in (ext_*.tmp) do (
    set "EXT=%%~nE"
    set "EXT=!EXT:~4!"
    set /a COUNT=0
    for /f %%C in ('find /c /v "" ^< "%%E"') do set /a COUNT=%%C
    echo !EXT!,!COUNT!,,,,, >> "%CSV_NAME%"
    
    sort "%%E" > "%%E.sorted"
    type "%%E.sorted" >> "%CSV_NAME%"
    
    del "%%E" >nul 2>&1
    del "%%E.sorted" >nul 2>&1
)

:: �N���[���A�b�v
del "%TEMP_FILE%" >nul 2>&1
del "%TEMP_FILE%.sorted" >nul 2>&1

echo.
echo ? CSV�o�͂��������܂��� �� %CSV_NAME%
pause
exit /b
