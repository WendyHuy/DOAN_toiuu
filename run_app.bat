@echo off
chcp 65001 > nul
echo ========================================================
echo     DO AN TOI UU HOA - GENETIC ALGORITHM (TSP TP.HCM)
echo     Giang vien huong dan: ThS. Tran Ngoc Viet
echo ========================================================
echo.
echo 1. Chay ung dung Giao dien do hoa truc quan (GUI)
echo 2. Chay ung dung Dong lenh mo phong (CLI)
echo 3. Chay kiem thu thuat toan (Test GA)
echo.
set /p choice="Nhap lua chon cua ban [1/2/3, Enter = 1]: "

if "%choice%"=="2" (
    "C:\Users\Huy\AppData\Local\Programs\Python\Python313\python.exe" run_cli.py
) else if "%choice%"=="3" (
    "C:\Users\Huy\AppData\Local\Programs\Python\Python313\python.exe" test_ga.py
) else (
    "C:\Users\Huy\AppData\Local\Programs\Python\Python313\python.exe" app_gui.py
)

pause
