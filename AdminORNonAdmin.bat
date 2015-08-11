@echo off

call :isAdmin
if %errorlevel% == 0 (
echo Running with admin rights.
) else (
echo Error: Access denied.
)

pause >nul
exit /b

:isAdmin
fsutil dirty query %systemdrive% >nul
exit /b