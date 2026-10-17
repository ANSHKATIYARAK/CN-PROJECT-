@echo off
title VIT Vellore Enterprise OSPF Campus Routing Platform
cls
echo ===============================================================================
echo     VIT VELLORE ENTERPRISE MULTI-AREA OSPF ROUTING PLATFORM (RFC 2328)
echo     Ast08 Project - Team: Ansh Katiyar, Arham Choudhary, Sri Harsha
echo ===============================================================================
echo.
echo Launching Interactive Campus Network Simulation in Dedicated App Mode...
echo.

set PROJECT_DIR=%~dp0
set HTML_PATH=%PROJECT_DIR%index.html

:: Check for Google Chrome
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" (
    echo [OK] Launching in Google Chrome App Window...
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --app="file:///%HTML_PATH:\=/%" --start-maximized
    goto done
)

:: Check for Microsoft Edge
if exist "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" (
    echo [OK] Launching in Microsoft Edge App Window...
    start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --app="file:///%HTML_PATH:\=/%" --start-maximized
    goto done
)

:: Default browser fallback
echo [OK] Launching in default web browser...
start "" "%HTML_PATH%"

:done
echo.
echo ===============================================================================
echo  Application is running!
echo  Features ready for demonstration:
echo   1. Live 3D Campus Map and Dijkstra Shortest Path Tree
echo   2. Dynamic Failover Demo (Sub-Second Link Rerouting via Tech Tower)
echo   3. Cisco IOS Interactive Console Terminal (Show OSPF Neighbors / Routes)
echo   4. 98.4%% TCAM Route Compression Proof and VLSM Addressing
echo ===============================================================================
ping 127.0.0.1 -n 4 >nul
