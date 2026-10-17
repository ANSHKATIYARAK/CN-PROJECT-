@echo off
title Cisco Packet Tracer 1-Click Launch and Configuration Suite
cls
echo ===============================================================================
echo     CISCO PACKET TRACER 1-CLICK LAUNCHER AND CONFIGURATION SUITE
echo     VIT Vellore 5-Area OSPF Enterprise Campus Network
echo ===============================================================================
echo.

set PT_FOUND=0
set PT_EXE=""

:: Check common installation paths for Cisco Packet Tracer
for %%P in (
    "C:\Program Files\Cisco Packet Tracer 8.2.1\bin\PacketTracer.exe"
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
echo [!] Cisco Packet Tracer is not detected in standard Program Files directories.
echo.
echo Opening the Official Free Cisco NetAcad Download Portal...
start "" "https://www.netacad.com/resources/lab-downloads"
echo.
echo Opening the Master Router Configurations and Visual Topology Map...
goto open_docs

:launch_pt
echo [OK] Detected Cisco Packet Tracer at: %PT_EXE%
echo Launching Cisco Packet Tracer...
start "" %PT_EXE%

:open_docs
set PROJECT_DIR=%~dp0
start "" "%PROJECT_DIR%cisco_packet_tracer\ALL_ROUTERS_MASTER_CONFIG.txt"
start "" "%PROJECT_DIR%cisco_packet_tracer\PACKET_TRACER_BUILD_AND_DEMO_GUIDE.md"
start "" "%PROJECT_DIR%assets\cisco_packet_tracer_topology.jpg"

echo.
echo ===============================================================================
echo  SUCCESS:
echo   - Master Router Configurations opened in Notepad (ALL_ROUTERS_MASTER_CONFIG.txt)
echo   - Step-by-Step Evaluator Demo Guide opened (PACKET_TRACER_BUILD_AND_DEMO_GUIDE.md)
echo   - Visual Topology Blueprint opened (cisco_packet_tracer_topology.jpg)
echo ===============================================================================
ping 127.0.0.1 -n 4 >nul
