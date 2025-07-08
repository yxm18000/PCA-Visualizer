@echo off
set VENV_DIR=venv

echo ���z�����m�F���܂�...

REM venv�t�H���_���Ȃ���΃Z�b�g�A�b�v�����s
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo.
    echo --------------------------------------------------
    echo ����Z�b�g�A�b�v���J�n���܂��B
    echo ����ɂ͐��������邱�Ƃ�����܂�...
    echo --------------------------------------------------
    echo.

    REM ���z�����쐬
    python -m venv %VENV_DIR%
    if %errorlevel% neq 0 (
        echo ���z���̍쐬�Ɏ��s���܂����B
        echo Python���C���X�g�[������APATH���ʂ��Ă��邩�m�F���Ă��������B
        pause
        exit
    )

    echo ���C�u�������C���X�g�[�����܂�...
    call %VENV_DIR%\Scripts\pip.exe install pandas scikit-learn matplotlib mplcursors
    if %errorlevel% neq 0 (
        echo ���C�u�����̃C���X�g�[���Ɏ��s���܂����B�C���^�[�l�b�g�ڑ����m�F���Ă��������B
        pause
        exit
    )

    echo.
    echo --------------------------------------------------
    echo �Z�b�g�A�b�v���������܂����B
    echo --------------------------------------------------
    echo.
)

echo PCA Visualizer���N�����܂�...
call %VENV_DIR%\Scripts\python.exe gui_pca_app.py

echo �A�v���P�[�V�������I�����܂����B
pause