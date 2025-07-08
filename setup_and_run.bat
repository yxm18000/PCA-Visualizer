@echo off
set VENV_DIR=venv

echo 仮想環境を確認します...

REM venvフォルダがなければセットアップを実行
if not exist "%VENV_DIR%\Scripts\python.exe" (
    echo.
    echo --------------------------------------------------
    echo 初回セットアップを開始します。
    echo これには数分かかることがあります...
    echo --------------------------------------------------
    echo.

    REM 仮想環境を作成
    python -m venv %VENV_DIR%
    if %errorlevel% neq 0 (
        echo 仮想環境の作成に失敗しました。
        echo Pythonがインストールされ、PATHが通っているか確認してください。
        pause
        exit
    )

    echo ライブラリをインストールします...
    call %VENV_DIR%\Scripts\pip.exe install pandas scikit-learn matplotlib mplcursors
    if %errorlevel% neq 0 (
        echo ライブラリのインストールに失敗しました。インターネット接続を確認してください。
        pause
        exit
    )

    echo.
    echo --------------------------------------------------
    echo セットアップが完了しました。
    echo --------------------------------------------------
    echo.
)

echo PCA Visualizerを起動します...
call %VENV_DIR%\Scripts\python.exe gui_pca_app.py

echo アプリケーションを終了しました。
pause