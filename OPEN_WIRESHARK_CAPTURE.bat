@echo off
title Open Wireshark OSPF Packet Capture
cls
echo ===============================================================================
echo     WIRESHARK OSPFv2 PACKET CAPTURE LAUNCHER
echo     VIT Vellore Enterprise 5-Area Campus Network
echo ===============================================================================
echo.

set PCAP_FILE=%~dp0assets\vit_campus_ospf_capture.pcap

if not exist "%PCAP_FILE%" (
    echo [ERROR] PCAP file not found at %PCAP_FILE%
    pause
    exit /b
)

:: Try to launch with Wireshark
if exist "C:\Program Files\Wireshark\Wireshark.exe" (
    echo [OK] Launching in Wireshark...
    start "" "C:\Program Files\Wireshark\Wireshark.exe" "%PCAP_FILE%"
    goto done
)

:: Fallback to Windows default file association
echo [INFO] Opening packet capture with default system viewer...
start "" "%PCAP_FILE%"

:done
echo.
echo ===============================================================================
echo  Captured Packets ready for analysis:
echo   - Frame 1: OSPF Hello from CTS-Core1 (Area 0 Backbone)
echo   - Frame 2: OSPF Hello from SJT-ABR (Area 10 Academic Core)
echo   - Frame 3: OSPF Database Description (DBD) Exchange
echo   - Frame 4: Link State Update (LSU) with 98.4%% Summary & NSSA Type-7 LSAs
echo   - Frame 5-6: End-to-end ICMP Ping Verification between SJT & Hostels
echo ===============================================================================
ping 127.0.0.1 -n 4 >nul
