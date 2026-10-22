@echo off
title Campus Network - Automated Simulation Launcher
color 0A
echo =====================================================================
echo    VIT CAMPUS MULTI-AREA OSPF NETWORK SIMULATION LAUNCHER
echo =====================================================================
echo.
echo [1/2] Launching Cisco Packet Tracer (Enterprise Campus Topology)...
start "" "C:\Users\VICTUS\Downloads\CN PROJECT\tools\CiscoPacketTracer\app\bin\PacketTracer.exe" "C:\Users\VICTUS\Downloads\CN PROJECT\Campus_Network_Topology.pkt"

echo [2/2] Launching Wireshark (OSPF Packet Capture)...
start "" "C:\Program Files\Wireshark\Wireshark.exe" "C:\Users\VICTUS\Downloads\CN PROJECT\vit_campus_ospf_capture.pcap"

echo.
echo =====================================================================
echo    Both Cisco Packet Tracer and Wireshark are now running!
echo =====================================================================
ping 127.0.0.1 -n 3 >nul
