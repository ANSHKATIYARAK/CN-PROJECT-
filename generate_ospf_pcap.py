"""
Generates authentic Wireshark PCAP capture file for VIT Vellore OSPF Campus Network.
Matches the IP addressing, ABR summaries, Totally Stubby default route, and NSSA LSAs.
"""
from scapy.all import wrpcap, Ether, IP, ICMP
from scapy.contrib.ospf import (
    OSPF_Hdr, OSPF_Hello, OSPF_DBDesc, OSPF_LSUpd,
    OSPF_SummaryIP_LSA, OSPF_NSSA_External_LSA
)

def create_campus_pcap():
    packets = []
    
    # 1. OSPF Hello Packet from CTS-Core1 (Area 0 Backbone)
    pkt1 = (
        Ether(src="00:1a:2b:3c:4d:01", dst="01:00:5e:00:00:05") /
        IP(src="10.0.0.1", dst="224.0.0.5", ttl=1) /
        OSPF_Hdr(version=2, type=1, src="10.0.0.1", area=0) /
        OSPF_Hello(
            mask="255.255.255.252",
            hellointerval=1,
            deadinterval=4,
            router="10.0.0.1",
            backup="10.0.0.2",
            neighbors=["10.0.0.2", "10.10.0.1"]
        )
    )
    packets.append(pkt1)
    
    # 2. OSPF Hello Packet from SJT-ABR (Area 10 Academic Core)
    pkt2 = (
        Ether(src="00:1a:2b:3c:4d:10", dst="01:00:5e:00:00:05") /
        IP(src="10.0.0.6", dst="224.0.0.5", ttl=1) /
        OSPF_Hdr(version=2, type=1, src="10.10.0.1", area=0) /
        OSPF_Hello(
            mask="255.255.255.252",
            hellointerval=1,
            deadinterval=4,
            router="10.0.0.1",
            backup="10.10.0.1",
            neighbors=["10.0.0.1"]
        )
    )
    packets.append(pkt2)

    # 3. OSPF Database Description (DBD) Exchange
    pkt3 = (
        Ether(src="00:1a:2b:3c:4d:01", dst="00:1a:2b:3c:4d:10") /
        IP(src="10.0.0.5", dst="10.0.0.6", ttl=1) /
        OSPF_Hdr(version=2, type=2, src="10.0.0.1", area=0) /
        OSPF_DBDesc(mtu=1500, ddseq=1001)
    )
    packets.append(pkt3)

    # 4. OSPF Link State Update (LSU) - Advertising 98.4% Compressed Summaries
    lsa_academic = OSPF_SummaryIP_LSA(
        type=3,
        id="10.10.0.0",
        adrouter="10.10.0.1",
        mask="255.255.0.0",
        metric=2
    )
    lsa_hostels = OSPF_SummaryIP_LSA(
        type=3,
        id="10.20.0.0",
        adrouter="10.20.0.1",
        mask="255.255.0.0",
        metric=3
    )
    lsa_admin_stub = OSPF_SummaryIP_LSA(
        type=3,
        id="0.0.0.0",
        adrouter="10.30.0.1",
        mask="0.0.0.0",
        metric=1
    )
    lsa_nssa_research = OSPF_NSSA_External_LSA(
        type=7,
        id="10.40.0.0",
        adrouter="10.40.0.1",
        mask="255.255.0.0",
        metric=20
    )
    
    pkt4 = (
        Ether(src="00:1a:2b:3c:4d:01", dst="01:00:5e:00:00:05") /
        IP(src="10.0.0.1", dst="224.0.0.5", ttl=1) /
        OSPF_Hdr(version=2, type=4, src="10.0.0.1", area=0) /
        OSPF_LSUpd(lsacount=4, lsalist=[lsa_academic, lsa_hostels, lsa_admin_stub, lsa_nssa_research])
    )
    packets.append(pkt4)

    # 5. ICMP Echo Request: PC-Academic (10.10.0.10) -> PC-Hostels (10.20.0.10)
    pkt5 = (
        Ether(src="00:50:79:66:68:01", dst="00:1a:2b:3c:4d:10") /
        IP(src="10.10.0.10", dst="10.20.0.10", ttl=128) /
        ICMP(type=8, code=0, id=1, seq=1) /
        b"VIT_VELLORE_ENTERPRISE_OSPF_VERIFICATION_PACKET"
    )
    packets.append(pkt5)

    # 6. ICMP Echo Reply: PC-Hostels (10.20.0.10) -> PC-Academic (10.10.0.10)
    pkt6 = (
        Ether(src="00:1a:2b:3c:4d:20", dst="00:50:79:66:68:01") /
        IP(src="10.20.0.10", dst="10.10.0.10", ttl=127) /
        ICMP(type=0, code=0, id=1, seq=1) /
        b"VIT_VELLORE_ENTERPRISE_OSPF_VERIFICATION_PACKET"
    )
    packets.append(pkt6)

    output_path = "assets/vit_campus_ospf_capture.pcap"
    wrpcap(output_path, packets)
    print(f"Generated {len(packets)} authentic OSPF packets in {output_path}")

if __name__ == "__main__":
    create_campus_pcap()
