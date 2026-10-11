import matplotlib.pyplot as plt
import matplotlib.patches as patches

def generate_topology_diagram():
    fig, ax = plt.subplots(figsize=(11, 7.2), dpi=300)
    ax.set_xlim(0, 110)
    ax.set_ylim(0, 80)
    ax.axis('off')

    # Background
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    # Title Banner
    title_box = patches.FancyBboxPatch((5, 72), 100, 6, boxstyle="round,pad=0.5,rounding_size=1.5",
                                       facecolor="#003366", edgecolor="#002244", linewidth=1.5)
    ax.add_patch(title_box)
    ax.text(55, 75, "VIT VELLORE CAMPUS HIERARCHICAL 5-AREA OSPF ARCHITECTURE",
            ha='center', va='center', color='white', fontname='Arial', fontsize=12, fontweight='bold')

    def draw_node_card(x, y, w, h, bg_color, border_color, area_tag, area_color, title, desc_lines, ip_text):
        # Outer Card
        card = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.4,rounding_size=1.2",
                                      facecolor=bg_color, edgecolor=border_color, linewidth=1.5)
        ax.add_patch(card)
        # Header Badge
        badge = patches.FancyBboxPatch((x + 2, y + h - 4.2), w - 4, 3.2, boxstyle="round,pad=0.2,rounding_size=0.8",
                                       facecolor=area_color, edgecolor='none')
        ax.add_patch(badge)
        ax.text(x + w/2, y + h - 2.6, area_tag, ha='center', va='center',
                color='white', fontname='Arial', fontsize=9, fontweight='bold')
        
        # Router Name / Title
        ax.text(x + w/2, y + h - 6.2, title, ha='center', va='center',
                color='#0F172A', fontname='Arial', fontsize=10, fontweight='bold')
        
        # IP Text
        ax.text(x + w/2, y + h - 8.6, f"Router ID: {ip_text}", ha='center', va='center',
                color='#0369A1', fontname='Arial', fontsize=8.5, fontweight='bold')

        # Descriptions
        cur_y = y + h - 11.2
        for line in desc_lines:
            ax.text(x + w/2, cur_y, line, ha='center', va='center',
                    color='#334155', fontname='Calibri', fontsize=8)
            cur_y -= 2.2

    # Area 10: Academic Core (Top Left)
    draw_node_card(6, 45, 42, 22, "#F0F7FF", "#0284C7",
                   "AREA 10: ACADEMIC CORE (10.10.0.0/16)", "#0284C7",
                   "SJT-ABR (10.10.0.1) & TT-ABR (10.10.64.1)",
                   ["• Silver Jubilee Tower (18,500 Students, Fl 1-12)",
                    "• Technology Tower Software Labs (5,200 Students)",
                    "• Summary Route: 10.10.0.0/16 (Aggregates 4 subnets)",
                    "• Inter-Tower 40G Conduit Backup (Link l8)"],
                   "10.10.0.1 / 10.10.64.1")

    # Area 20: Residential Hostels (Top Right)
    draw_node_card(62, 45, 42, 22, "#FAF5FF", "#7C3AED",
                   "AREA 20: RESIDENTIAL HOSTELS (10.20.0.0/16)", "#7C3AED",
                   "MH-ABR (10.20.0.1) & WH-ABR (10.20.128.1)",
                   ["• Men's Hostels Blocks A-T (22,000 Resident Beds)",
                    "• Women's Hostels Blocks A-J (10,500 Resident Beds)",
                    "• Summary Route: 10.20.0.0/16 (Aggregates 6 subnets)",
                    "• Strict Failure Domain Isolation against flap storms"],
                   "10.20.0.1 / 10.20.128.1")

    # Area 0: Central Backbone Spine (Center)
    draw_node_card(25, 23, 60, 16, "#F1F5F9", "#0F172A",
                   "AREA 0: BACKBONE CORE SPINE (10.0.0.0/24)", "#003366",
                   "CTS Core 1 & 2 Datacenter Spine Switches",
                   ["• Dual Redundant Nexus/Catalyst 9500 (Centre for Technical Support)",
                    "• Ultra-low latency 100G & 40G Single-Mode Fiber Conduits",
                    "• Central FIB Supernetted Route Table (Only 4 Summary Prefixes)"],
                   "10.0.0.1 <===> 10.0.0.2")

    # Area 30: Administration (Bottom Left)
    draw_node_card(6, 2, 42, 16, "#FFFBEB", "#D97706",
                   "AREA 30: TOTALLY STUBBY (SECURITY SHIELD)", "#D97706",
                   "Admin-ABR (Gandhi Block Central Administration)",
                   ["• Finance, Payroll & Examination Database Servers",
                    "• Config: 'area 30 stub no-summary' (Blocks LSA 3, 4, 5)",
                    "• Zero Reconnaissance: Only Default Route 0.0.0.0/0"],
                   "10.30.0.1")

    # Area 40: Research Park (Bottom Right)
    draw_node_card(62, 2, 42, 16, "#ECFDF5", "#059669",
                   "AREA 40: NOT-SO-STUBBY AREA (NSSA)", "#059669",
                   "PRP-ABR (Pearl Research Park HPC Clusters)",
                   ["• High Performance Computing & AI Cluster Nodes",
                    "• External RIP/BGP Redistribution via Type-7 LSAs",
                    "• ABR translates Type-7 to Type-5 LSA into Area 0"],
                   "10.40.0.1")

    # Connecting Links (Arrows with Labels)
    def draw_link(x1, y1, x2, y2, label, speed_color="#0284C7", offset_text=(0, 0)):
        ax.annotate('', xy=(x2, y2), xytext=(x1, y1),
                    arrowprops=dict(arrowstyle='<->', color='#334155', lw=2, mutation_scale=14))
        mid_x = (x1 + x2) / 2 + offset_text[0]
        mid_y = (y1 + y2) / 2 + offset_text[1]
        bbox_props = dict(boxstyle='round,pad=0.25', facecolor='#FFFFFF', edgecolor='#CBD5E1', lw=1)
        ax.text(mid_x, mid_y, label, ha='center', va='center',
                fontsize=7.5, fontname='Arial', fontweight='bold', color=speed_color, bbox=bbox_props)

    # Area 10 <-> Area 0
    draw_link(27, 45, 38, 39, "100G Trunk l1 & l2\n10.0.0.4/30 & 10.0.0.8/30", "#0284C7", offset_text=(-5, 0))
    # Area 20 <-> Area 0
    draw_link(83, 45, 72, 39, "40G Trunk l3 & l4\n10.0.0.12/30 & 10.0.0.16/30", "#7C3AED", offset_text=(5, 0))
    # Area 0 <-> Area 30
    draw_link(40, 23, 27, 18, "10G Armored Fiber l5\n10.0.0.20/30", "#D97706", offset_text=(-5, 0))
    # Area 0 <-> Area 40
    draw_link(70, 23, 83, 18, "100G Optical Fiber l6\n10.0.0.24/30", "#059669", offset_text=(5, 0))

    plt.tight_layout()
    output_path = r"c:\Users\VICTUS\Downloads\CN PROJECT\figures\vit_ospf_5area_schematic.png"
    plt.savefig(output_path, dpi=300, facecolor=fig.get_facecolor(), edgecolor='none', bbox_inches='tight')
    plt.close()
    print(f"Generated schematic diagram: {output_path}")

if __name__ == "__main__":
    generate_topology_diagram()
