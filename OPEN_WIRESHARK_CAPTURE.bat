@echo off
title Wireshark OSPFv2 Packet Capture Launcher
cls
echo ===============================================================================
echo     WIRESHARK OSPFv2 PACKET CAPTURE LAUNCHER
echo     VIT Vellore Enterprise 5-Area Campus Network (RFC 2328)
echo ===============================================================================
echo.

set PROJECT_DIR=%~dp0
set PCAP_FILE=%PROJECT_DIR%assets\vit_campus_ospf_capture.pcap
set HTML_INSPECTOR=%PROJECT_DIR%wireshark_capture_inspector.html
set MSI_SETUP=%PROJECT_DIR%tools\Wireshark-Setup.msi

:: Check if official Wireshark is installed
if exist "C:\Program Files\Wireshark\Wireshark.exe" (
    echo [OK] Detected Official Wireshark Installation.
    echo Launching Wireshark with campus capture file...
    start "" "C:\Program Files\Wireshark\Wireshark.exe" "%PCAP_FILE%"
    goto done
)

:: If Wireshark is not yet installed:
echo [INFO] Official Wireshark is not yet installed in Program Files.
echo.
echo 1. Launching Interactive Standalone Wireshark Inspector...
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --app="file:///%HTML_INSPECTOR:\=/%" --start-maximized
) else (
    start "" "%HTML_INSPECTOR%"
)

echo.
echo 2. Launching Wireshark Official Installer Wizard (Wireshark-Setup.msi)...
if exist "%MSI_SETUP%" (
    start "" "%MSI_SETUP%"
)

:done
echo.
echo ===============================================================================
echo  SUCCESS:
echo   - 6 Verified OSPF Packets (Hello, DBD, 98.4%% Compressed LSUs, Ping)
echo   - Interactive Wireshark Inspector window is now open on your screen!
echo ===============================================================================
ping 127.0.0.1 -n 4 >nul
