@echo off
title Cisco Packet Tracer 1-Click Launch and Configuration Suite
cls
echo ===============================================================================
echo     CISCO PACKET TRACER 1-CLICK LAUNCHER AND TOPOLOGY SUITE
echo     VIT Vellore 5-Area OSPF Enterprise Campus Network
echo ===============================================================================
echo.

set PROJECT_DIR=%~dp0
set PKT_FILE=%PROJECT_DIR%cisco_packet_tracer\Campus_Network_Topology.pkt
set SETUP_EXE=%PROJECT_DIR%tools\CiscoPacketTracer_820_setup.exe
set PT_FOUND=0
set PT_EXE=""

:: Check local tools and system installation paths for Cisco Packet Tracer
for %%P in (
    "%PROJECT_DIR%tools\CiscoPacketTracer\app\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 8.2.2\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 8.2.1\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 8.2.0\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 8.2\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 8.1.1\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 8.0\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer 7.3.1\bin\PacketTracer.exe"
    "C:\Program Files\Cisco Packet Tracer\bin\PacketTracer.exe"
    "C:\Program Files (x86)\Cisco Packet Tracer 8.2\bin\PacketTracer.exe"
    "C:\Program Files (x86)\Cisco Packet Tracer 8.1\bin\PacketTracer.exe"
    "C:\Program Files (x86)\Cisco Packet Tracer\bin\PacketTracer.exe"
) do (
    if exist %%P (
        set PT_EXE=%%P
        set PT_FOUND=1
        goto launch_pt
    )
)

:no_pt
echo [!] Cisco Packet Tracer executable not yet detected in Program Files.
echo.
if exist "%SETUP_EXE%" (
    echo [OK] Found Cisco Packet Tracer Setup Installer in tools folder!
    echo Launching installer wizard (CiscoPacketTracer_820_setup.exe)...
    start "" "%SETUP_EXE%"
) else (
    echo [INFO] Opening the Official Free Cisco NetAcad Download Portal...
    start "" "https://www.netacad.com/resources/lab-downloads"
)
goto open_docs

:launch_pt
echo [OK] Detected Cisco Packet Tracer at: %PT_EXE%
echo Launching Cisco Packet Tracer...
for %%F in (%PT_EXE%) do set "PT_DIR=%%~dpF"
start "" /d "%PT_DIR%" %PT_EXE%

:open_docs
echo.
echo Opening Master Configurations, Evaluator Guide, and Topology Blueprint...
start "" "%PROJECT_DIR%cisco_packet_tracer\ALL_ROUTERS_MASTER_CONFIG.txt"
start "" "%PROJECT_DIR%cisco_packet_tracer\PACKET_TRACER_BUILD_AND_DEMO_GUIDE.md"
start "" "%PROJECT_DIR%assets\cisco_packet_tracer_topology.jpg"

echo.
echo ===============================================================================
echo  SUCCESS:
echo   - Campus_Network_Topology.pkt workspace is ready in cisco_packet_tracer\
echo   - Master Router Configurations opened in Notepad (ALL_ROUTERS_MASTER_CONFIG.txt)
echo   - Step-by-Step Evaluator Demo Guide opened (PACKET_TRACER_BUILD_AND_DEMO_GUIDE.md)
echo   - Visual Topology Blueprint opened (cisco_packet_tracer_topology.jpg)
echo ===============================================================================
ping 127.0.0.1 -n 4 >nul
