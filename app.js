/**
 * ==============================================================================
 * VIT VELLORE ENTERPRISE MULTI-AREA OSPF ROUTING PLATFORM (RFC 2328)
 * Clean Architecture, Interactive 3D Showcase, Dynamic Failover & Cisco Terminal
 * ==============================================================================
 */

// 1. Campus Landmark Topology & Node Definitions
const CAMPUS_TOPOLOGY = {
  nodes: [
    {
      id: 'SJT-ABR',
      label: 'SJT Tower',
      fullName: 'Silver Jubilee Tower (SJT-ABR)',
      role: 'Area 10 - Academic Core ABR',
      rid: '10.10.0.1',
      area: 10,
      subnet: '10.10.0.0/18',
      capacity: '16,382 hosts (Floors 1-12, 18,500 daily students)',
      desc: 'Hosts CSE, SCOPE, and SENSE academic departments across 12 floors. Serves as Area Border Router (ABR) aggregating academic subnets into 10.10.0.0/16 before advertising to the backbone.'
    },
    {
      id: 'TT-ABR',
      label: 'Tech Tower',
      fullName: 'Technology Tower (TT-ABR)',
      role: 'Area 10 - Computing Labs Core ABR',
      rid: '10.10.64.1',
      area: 10,
      subnet: '10.10.64.0/18',
      capacity: '16,382 hosts (Advanced Computing & Cloud Labs)',
      desc: 'Houses high-density computing laboratories, enterprise server stacks, and software engineering suites. Directly interconnected with SJT via 40G optical conduit for inter-tower redundancy.'
    },
    {
      id: 'Admin-ABR',
      label: 'Gandhi Block Admin',
      fullName: 'Gandhi Block Central Admin (Admin-ABR)',
      role: 'Area 30 - Totally Stubby Area (TSA)',
      rid: '10.30.0.1',
      area: 30,
      subnet: '10.30.0.0/16',
      capacity: '65,534 hosts (Administrative, Registrar & Controller)',
      desc: 'Hosts central university administration, examination controllers, and finance databases. Configured with "area 30 stub no-summary" to block external and inter-area route leakages, substituting an injected 0.0.0.0/0 default route.'
    },
    {
      id: 'MH-ABR',
      label: "Men's Hostels",
      fullName: "Men's Hostels Complex (MH-ABR)",
      role: 'Area 20 - Residential Zone ABR',
      rid: '10.20.0.1',
      area: 20,
      subnet: '10.20.0.0/17',
      capacity: '32,766 hosts (Blocks A to T - 22,000 residents)',
      desc: 'Supports 22,000 resident students across 20 hostel towers. Isolates high-volume residential Wi-Fi access port flapping inside Area 20, shielding the academic core from Dijkstra SPF churn.'
    },
    {
      id: 'WH-ABR',
      label: "Women's Hostels",
      fullName: "Women's Hostels Complex (WH-ABR)",
      role: 'Area 20 - Residential Zone ABR',
      rid: '10.20.128.1',
      area: 20,
      subnet: '10.20.128.0/17',
      capacity: '32,766 hosts (Blocks A to H - 10,500 residents)',
      desc: 'Serves 10,500 resident students across Blocks A through H. Aggregated with Men\'s Hostels into the summary prefix 10.20.0.0/16 before core distribution.'
    },
    {
      id: 'PRP-ABR',
      label: 'Pearl Research Park',
      fullName: 'Pearl Research Park (PRP-ABR)',
      role: 'Area 40 - Not-So-Stubby Area (NSSA)',
      rid: '10.40.0.1',
      area: 40,
      subnet: '10.40.0.0/16',
      capacity: '65,534 hosts (Supercomputing & AI Incubation)',
      desc: 'Houses multidisciplinary research labs, high-performance computing clusters, and industry incubators. Operates as an NSSA, permitting local external route redistribution via Type-7 LSAs translated to Type-5 at the boundary.'
    },
    {
      id: 'CTS-Core1',
      label: 'CTS Core DC',
      fullName: 'Centre for Technical Support (CTS-Core1)',
      role: 'Area 0 - Enterprise Backbone Core',
      rid: '10.0.0.1',
      area: 0,
      subnet: '10.0.0.0/24',
      capacity: '254 Point-to-Point 100G Interconnects',
      desc: 'The central switching nexus and primary datacenter of VIT Vellore. Terminates all inter-area optical links with zero summarization, running Dijkstra across the authoritative campus link-state database.'
    }
  ],
  links: [
    { id: 'l1', source: 'CTS-Core1', target: 'SJT-ABR', cost: 1, state: 'UP', desc: '100G Optical Trunk 1' },
    { id: 'l2', source: 'CTS-Core1', target: 'TT-ABR', cost: 1, state: 'UP', desc: '100G Optical Trunk 2' },
    { id: 'l3', source: 'CTS-Core1', target: 'MH-ABR', cost: 2, state: 'UP', desc: '40G Residential Trunk' },
    { id: 'l4', source: 'CTS-Core1', target: 'WH-ABR', cost: 2, state: 'UP', desc: '40G Residential Trunk' },
    { id: 'l5', source: 'CTS-Core1', target: 'Admin-ABR', cost: 1, state: 'UP', desc: '10G Admin Secure Trunk' },
    { id: 'l6', source: 'CTS-Core1', target: 'PRP-ABR', cost: 1, state: 'UP', desc: '100G Research Spine' },
    { id: 'l8', source: 'SJT-ABR', target: 'TT-ABR', cost: 1, state: 'UP', desc: '40G Inter-Tower Conduit' },
    { id: 'l9', source: 'MH-ABR', target: 'WH-ABR', cost: 1, state: 'UP', desc: '10G Inter-Hostel Ring' }
  ]
};

// 2. Production OSPF Engine with Full Dijkstra Shortest Path Tree (RFC 2328)
class OSPFEngine {
  constructor(topology) {
    this.topology = JSON.parse(JSON.stringify(topology));
    this.selectedNodeId = 'Admin-ABR';
    this.routingTables = {};
    this.rebuildSPF();
  }

  getLink(id) {
    return this.topology.links.find(l => l.id === id);
  }

  setLinkState(id, state) {
    const l = this.getLink(id);
    if (l) {
      l.state = state;
      this.rebuildSPF();
    }
  }

  // Dijkstra Shortest Path Tree Algorithm: O(|E| + |V| log |V|)
  computeSPT(startId) {
    const distances = {};
    const previous = {};
    const unvisited = new Set();

    this.topology.nodes.forEach(n => {
      distances[n.id] = Infinity;
      previous[n.id] = null;
      unvisited.add(n.id);
    });
    distances[startId] = 0;

    while (unvisited.size > 0) {
      let minNode = null;
      unvisited.forEach(nodeId => {
        if (minNode === null || distances[nodeId] < distances[minNode]) {
          minNode = nodeId;
        }
      });

      if (distances[minNode] === Infinity) break;
      unvisited.delete(minNode);

      // Inspect operational links connected to minNode
      const activeLinks = this.topology.links.filter(
        l => l.state === 'UP' && (l.source === minNode || l.target === minNode)
      );

      activeLinks.forEach(link => {
        const neighbor = (link.source === minNode) ? link.target : link.source;
        if (unvisited.has(neighbor)) {
          const alt = distances[minNode] + link.cost;
          if (alt < distances[neighbor]) {
            distances[neighbor] = alt;
            previous[neighbor] = { node: minNode, link: link };
          }
        }
      });
    }

    return { distances, previous };
  }

  // Generate Routing Table (FIB) for each Node
  rebuildSPF() {
    this.routingTables = {};

    this.topology.nodes.forEach(srcNode => {
      const { distances, previous } = this.computeSPT(srcNode.id);
      const routes = [];

      this.topology.nodes.forEach(tgtNode => {
        if (tgtNode.id === srcNode.id) return;
        if (distances[tgtNode.id] === Infinity) return;

        // Trace path back to source
        let curr = tgtNode.id;
        let step = previous[curr];
        const path = [curr];

        while (step && step.node !== srcNode.id) {
          curr = step.node;
          path.unshift(curr);
          step = previous[curr];
        }
        path.unshift(srcNode.id);

        let nextHop = 'Direct';
        let iface = 'Gig0/0';
        if (step) {
          nextHop = (step.link.source === srcNode.id) ? `${tgtNode.rid}` : `${srcNode.rid}`;
          iface = `Gig0/${step.link.id.replace('l', '')}`;
        }

        const isInterArea = (tgtNode.area !== srcNode.area);

        // Check Totally Stubby Area 30 suppression
        if (srcNode.area === 30 && isInterArea) {
          // TSA blocks explicit inter-area routes
          return;
        }

        routes.push({
          targetId: tgtNode.id,
          targetLabel: tgtNode.label,
          type: isInterArea ? 'O IA' : 'O',
          prefix: tgtNode.subnet,
          metric: distances[tgtNode.id],
          nextHop: nextHop,
          interface: iface,
          path: path
        });
      });

      // Inject default route 0.0.0.0/0 for Totally Stubby Area 30
      if (srcNode.area === 30) {
        routes.unshift({
          targetId: 'CTS-Core1',
          targetLabel: 'CTS Core DC',
          type: 'O IA',
          prefix: '0.0.0.0/0',
          metric: 1,
          nextHop: '10.0.0.1',
          interface: 'Gig0/5',
          path: [srcNode.id, 'CTS-Core1']
        });
      }

      this.routingTables[srcNode.id] = routes;
    });
  }

  // Calculate Shortest Path between any two nodes
  findShortestPath(srcId, tgtId) {
    const { distances, previous } = this.computeSPT(srcId);
    if (distances[tgtId] === Infinity) {
      return { path: [], cost: Infinity, latencyMs: 0 };
    }

    let curr = tgtId;
    const path = [curr];
    while (previous[curr] && previous[curr].node !== srcId) {
      curr = previous[curr].node;
      path.unshift(curr);
    }
    path.unshift(srcId);

    const cost = distances[tgtId];
    // Realistic campus fiber latency: ~0.4ms per hop + 0.2ms base
    const latency = ((path.length - 1) * 0.4 + 0.3).toFixed(1);

    return { path, cost, latencyMs: latency };
  }
}

// Initialize global state
const engine = new OSPFEngine(CAMPUS_TOPOLOGY);

// 3. DOM Initialization & Wiring
document.addEventListener('DOMContentLoaded', () => {
  initTabNavigation();
  initHotspotPins();
  initSwitchToggles();
  initPacketDispatcher();
  initTrafficWaveform();
  initFailoverDemo();
  initCiscoTerminal();

  // Initial UI Render
  updateInspector(engine.selectedNodeId);
});

// ==============================================================================
// FEATURE 1: TAB NAVIGATION
// ==============================================================================
function initTabNavigation() {
  const tabs = document.querySelectorAll('.nav-tab');
  const views = document.querySelectorAll('.app-view');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const targetId = `view-${tab.dataset.tab}`;

      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      views.forEach(v => {
        v.classList.remove('active');
        if (v.id === targetId) {
          v.classList.add('active');
        }
      });
    });
  });
}

// ==============================================================================
// FEATURE 2: 3D MAP CENTERPIECE HOTSPOT PINS & NODE INSPECTOR
// ==============================================================================
function initHotspotPins() {
  const pins = document.querySelectorAll('.hotspot-pin');

  pins.forEach(pin => {
    pin.addEventListener('click', () => {
      const nodeId = pin.dataset.node;
      if (!nodeId) return;

      pins.forEach(p => p.classList.remove('active'));
      pin.classList.add('active');

      engine.selectedNodeId = nodeId;
      updateInspector(nodeId);
    });
  });

  // Set initial active pin
  const defaultPin = document.querySelector(`.hotspot-pin[data-node="${engine.selectedNodeId}"]`);
  if (defaultPin) defaultPin.classList.add('active');
}

function updateInspector(nodeId) {
  const node = engine.topology.nodes.find(n => n.id === nodeId);
  if (!node) return;

  const nameEl = document.getElementById('insp-node-name');
  const roleEl = document.getElementById('insp-node-role');
  const descEl = document.getElementById('insp-node-desc');
  const ridEl = document.getElementById('insp-rid');
  const subnetEl = document.getElementById('insp-subnet');
  const tbody = document.getElementById('inspector-routes-tbody');

  if (nameEl) nameEl.textContent = node.fullName;
  if (roleEl) roleEl.textContent = node.role;
  if (descEl) descEl.textContent = node.desc;
  if (ridEl) ridEl.textContent = node.rid;
  if (subnetEl) subnetEl.textContent = node.subnet;

  if (tbody) {
    tbody.innerHTML = '';
    const routes = engine.routingTables[nodeId] || [];

    if (routes.length === 0) {
      tbody.innerHTML = '<tr><td colspan="4" style="text-align:center; color:#94a3b8;">No reachable routes</td></tr>';
      return;
    }

    routes.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><span class="pill ${r.type === 'O IA' ? 'area20' : 'area10'}">${r.type}</span></td>
        <td><code>${r.prefix}</code></td>
        <td><strong class="text-cyan">${r.metric}</strong></td>
        <td>${r.nextHop}</td>
      `;
      tbody.appendChild(tr);
    });
  }
}

// ==============================================================================
// FEATURE 3: LEFT SWITCH TOGGLES (SIMULATING LINK FLAPS)
// ==============================================================================
function initSwitchToggles() {
  const switchMap = [
    { id: 'sw-link-l1', linkId: 'l1' },
    { id: 'sw-link-l2', linkId: 'l2' },
    { id: 'sw-link-l3', linkId: 'l3' },
    { id: 'sw-link-l8', linkId: 'l8' }
  ];

  switchMap.forEach(item => {
    const el = document.getElementById(item.id);
    if (!el) return;

    el.addEventListener('change', (e) => {
      const state = e.target.checked ? 'UP' : 'DOWN';
      engine.setLinkState(item.linkId, state);
      updateInspector(engine.selectedNodeId);

      // Re-run dispatcher if active
      updateDispatcherOutput();
    });
  });
}

// ==============================================================================
// FEATURE 4: TEST PACKET DISPATCHER
// ==============================================================================
function initPacketDispatcher() {
  const srcSel = document.getElementById('sel-src-node');
  const tgtSel = document.getElementById('sel-tgt-node');
  const btn = document.getElementById('btn-run-packet');

  if (!srcSel || !tgtSel || !btn) return;

  srcSel.innerHTML = '';
  tgtSel.innerHTML = '';

  engine.topology.nodes.forEach(n => {
    srcSel.innerHTML += `<option value="${n.id}">${n.label} (${n.id})</option>`;
    tgtSel.innerHTML += `<option value="${n.id}">${n.label} (${n.id})</option>`;
  });

  srcSel.value = 'SJT-ABR';
  tgtSel.value = 'MH-ABR';

  btn.addEventListener('click', () => {
    updateDispatcherOutput(true);
  });

  srcSel.addEventListener('change', () => updateDispatcherOutput());
  tgtSel.addEventListener('change', () => updateDispatcherOutput());

  updateDispatcherOutput();
}

function updateDispatcherOutput(animate = false) {
  const srcSel = document.getElementById('sel-src-node');
  const tgtSel = document.getElementById('sel-tgt-node');
  const pathText = document.getElementById('disp-path-text');
  const metricText = document.getElementById('disp-metric-text');

  if (!srcSel || !tgtSel || !pathText || !metricText) return;

  const src = srcSel.value;
  const tgt = tgtSel.value;

  if (src === tgt) {
    pathText.textContent = `${src} (Local Loopback)`;
    pathText.className = 'text-green';
    metricText.textContent = 'Cost: 0 | 0.0 ms';
    return;
  }

  const result = engine.findShortestPath(src, tgt);

  if (result.cost === Infinity) {
    pathText.textContent = 'NO PATH AVAILABLE (Partitioned)';
    pathText.className = 'text-danger';
    metricText.textContent = 'Cost: ∞ | Unreachable';
    return;
  }

  const pathStr = result.path.map(id => {
    const node = engine.topology.nodes.find(n => n.id === id);
    return node ? node.label : id;
  }).join(' → ');

  pathText.textContent = pathStr;
  pathText.className = 'text-green';
  metricText.textContent = `Cost: ${result.cost} | ${result.latencyMs} ms`;

  if (animate) {
    // Pulse the pins along the path
    result.path.forEach((id, idx) => {
      const pin = document.querySelector(`.hotspot-pin[data-node="${id}"]`);
      if (pin) {
        setTimeout(() => {
          pin.classList.add('pulse');
          setTimeout(() => pin.classList.remove('pulse'), 900);
        }, idx * 250);
      }
    });
  }
}

// ==============================================================================
// FEATURE 5: REAL-TIME TRAFFIC OSCILLOSCOPE WAVEFORM
// ==============================================================================
function initTrafficWaveform() {
  const canvas = document.getElementById('traffic-waveform-canvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');

  let offset = 0;

  function renderWave() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);

    ctx.beginPath();
    ctx.strokeStyle = 'rgba(0, 242, 254, 0.85)';
    ctx.lineWidth = 2;
    ctx.shadowBlur = 8;
    ctx.shadowColor = '#00f2fe';

    const w = canvas.width;
    const h = canvas.height;
    const mid = h / 2;

    for (let x = 0; x < w; x++) {
      const y = mid +
        Math.sin((x + offset) * 0.05) * 12 +
        Math.sin((x * 0.1) - offset * 0.03) * 6;
      if (x === 0) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    }
    ctx.stroke();

    // Baseline grid glow
    ctx.beginPath();
    ctx.strokeStyle = 'rgba(16, 185, 129, 0.25)';
    ctx.lineWidth = 1;
    ctx.shadowBlur = 0;
    for (let gx = 0; gx < w; gx += 20) {
      ctx.moveTo(gx, 0);
      ctx.lineTo(gx, h);
    }
    ctx.stroke();

    offset += 2;
    requestAnimationFrame(renderWave);
  }

  renderWave();
}

// ==============================================================================
// FEATURE 6: DYNAMIC FAILOVER DEMO CONSOLE (VIEW 5)
// ==============================================================================
function initFailoverDemo() {
  const btnCut = document.getElementById('btn-trigger-flap-demo');
  const btnRestore = document.getElementById('btn-restore-flap-demo');
  const statusPill = document.getElementById('demo-status-pill');
  const primaryPath = document.getElementById('t-primary-path');
  const latencyVal = document.getElementById('t-latency-val');
  const lossVal = document.getElementById('t-loss-val');
  const logScreen = document.getElementById('demo-log-screen');

  if (!btnCut || !btnRestore) return;

  btnCut.addEventListener('click', () => {
    // Cut link l1
    engine.setLinkState('l1', 'DOWN');

    // Sync switch in View 1
    const sw1 = document.getElementById('sw-link-l1');
    if (sw1) sw1.checked = false;

    if (statusPill) {
      statusPill.textContent = 'OPTICAL TRUNK CUT - REROUTED';
      statusPill.style.color = '#ef4444';
      statusPill.style.borderColor = '#ef4444';
      statusPill.style.backgroundColor = 'rgba(239, 68, 68, 0.15)';
    }

    if (primaryPath) {
      primaryPath.textContent = 'SJT-ABR → TT-ABR → CTS-Core1 (Inter-Tower Backup)';
      primaryPath.className = 't-val text-amber';
    }

    if (latencyVal) latencyVal.textContent = '1.8 milliseconds';
    if (lossVal) lossVal.textContent = '0.00% (Zero Drops)';

    const now = new Date().toISOString().substring(11, 19);
    logScreen.textContent =
`[ ${now} ] TRUNK FAILURE DETECTED: Conduits CTS-Core1 <-> SJT-ABR (GigabitEthernet0/1) DOWN
%OSPF-5-ADJCHG: Process 1, Nbr 10.10.0.1 on GigabitEthernet0/1 from FULL to DOWN, Interface Down
%OSPF-5-SPF: Link State Advertisement (LSA Type-1) flooded to Area 10
%OSPF-5-SPF: Partial SPF recalculated in 0.42 ms across 3 active routers
%CEF-4-FAILOVER: Fast-Reroute cutover to secondary tree via Tech Tower (TT-ABR)
================================================================================
NEW ACTIVE PATH: SJT-ABR -> Tech Tower (Gig0/8) -> CTS-Core1 (Gig0/2)
TOTAL CONVERGENCE TIME : 1.8 ms
PACKET LOSS            : 0.00% (Sub-second hardware forwarding maintained)
CAMPUS CLIENTS RESILIENT: 40,000 students sessions intact
================================================================================`;

    updateInspector(engine.selectedNodeId);
    updateDispatcherOutput();
  });

  btnRestore.addEventListener('click', () => {
    // Restore link l1
    engine.setLinkState('l1', 'UP');

    // Sync switch in View 1
    const sw1 = document.getElementById('sw-link-l1');
    if (sw1) sw1.checked = true;

    if (statusPill) {
      statusPill.textContent = 'TRUNK OPERATIONAL';
      statusPill.style.color = '#00ff88';
      statusPill.style.borderColor = '#00ff88';
      statusPill.style.backgroundColor = 'rgba(16, 185, 129, 0.15)';
    }

    if (primaryPath) {
      primaryPath.textContent = 'CTS-Core1 → SJT-ABR (Direct 100G Primary)';
      primaryPath.className = 't-val text-green';
    }

    if (latencyVal) latencyVal.textContent = '0.7 milliseconds';
    if (lossVal) lossVal.textContent = '0.00% (Zero Drops)';

    const now = new Date().toISOString().substring(11, 19);
    logScreen.textContent =
`[ ${now} ] OPTICAL CONDUIT RESTORED: CTS-Core1 <-> SJT-ABR (GigabitEthernet0/1) UP
%LINK-3-UPDOWN: Interface GigabitEthernet0/1, changed state to up
%OSPF-5-ADJCHG: Process 1, Nbr 10.10.0.1 on GigabitEthernet0/1 from LOADING to FULL, Loading Done
%OSPF-5-SPF: Optimal Shortest Path Tree restored. Metric cost reset to 1.
================================================================================
PRIMARY 100G HIGH-SPEED FIBER TRUNK ACTIVE
NORMAL OPERATING STATUS VERIFIED ACROSS ALL 5 AREAS
================================================================================`;

    updateInspector(engine.selectedNodeId);
    updateDispatcherOutput();
  });
}

// ==============================================================================
// FEATURE 7: CISCO IOS TERMINAL CONSOLE (VIEW 6)
// ==============================================================================
function initCiscoTerminal() {
  const input = document.getElementById('full-terminal-input');
  const output = document.getElementById('full-terminal-output');
  const quickBtns = document.querySelectorAll('.quick-cmd');

  if (!input || !output) return;

  function executeCommand(cmdRaw) {
    const cmd = cmdRaw.trim();
    if (!cmd) return;

    output.innerHTML += `\n<span style="color:#00f2fe;">cisco-ios-core1# ${cmd}</span>\n`;

    const lower = cmd.toLowerCase();

    if (lower === 'show ip ospf neighbor' || lower === 'sh ip ospf nei') {
      output.innerHTML +=
`Neighbor ID     Pri   State           Dead Time   Address         Interface
10.0.0.2          1   FULL/DR         00:00:03    10.0.0.2        GigabitEthernet0/1
10.10.0.1         1   ${engine.getLink('l1').state === 'UP' ? 'FULL/DR' : 'DOWN   '}         00:00:03    10.0.0.6        GigabitEthernet0/2
10.10.64.1        1   FULL/DR         00:00:03    10.0.0.10       GigabitEthernet0/3
10.20.0.1         1   FULL/DR         00:00:03    10.0.0.14       GigabitEthernet0/4
10.20.128.1       1   FULL/DR         00:00:03    10.0.0.18       GigabitEthernet0/5
10.30.0.1         1   FULL/DR         00:00:03    10.0.0.22       GigabitEthernet0/6
10.40.0.1         1   FULL/DR         00:00:03    10.0.0.26       GigabitEthernet0/11\n`;
    }
    else if (lower === 'show ip route ospf' || lower === 'sh ip ro ospf') {
      output.innerHTML +=
`Codes: L - local, C - connected, S - static, R - RIP, M - mobile, B - BGP
       D - EIGRP, EX - EIGRP external, O - OSPF, IA - OSPF inter area
       N1 - OSPF NSSA external type 1, N2 - OSPF NSSA external type 2
       E1 - OSPF external type 1, E2 - OSPF external type 2

Gateway of last resort is not set

O IA  10.10.0.0/16 [110/2] via 10.0.0.6, 02:44:19, GigabitEthernet0/2
O IA  10.20.0.0/16 [110/3] via 10.0.0.14, 02:44:19, GigabitEthernet0/4
O IA  10.30.0.0/16 [110/2] via 10.0.0.22, 02:44:19, GigabitEthernet0/6
O IA  10.40.0.0/16 [110/2] via 10.0.0.26, 02:44:19, GigabitEthernet0/11\n`;
    }
    else if (lower.startsWith('show ip route') || lower.startsWith('sh ip ro')) {
      output.innerHTML +=
`Codes: C - connected, S - static, O - OSPF, IA - OSPF inter area

      10.0.0.0/8 is variably subnetted, 9 subnets, 4 masks
C        10.0.0.0/24 is directly connected, GigabitEthernet0/1
C        10.0.0.4/30 is directly connected, GigabitEthernet0/2
C        10.0.0.8/30 is directly connected, GigabitEthernet0/3
O IA     10.10.0.0/16 [110/2] via 10.0.0.6, 02:45:01, GigabitEthernet0/2
O IA     10.20.0.0/16 [110/3] via 10.0.0.14, 02:45:01, GigabitEthernet0/4
O IA     10.30.0.0/16 [110/2] via 10.0.0.22, 02:45:01, GigabitEthernet0/6
O IA     10.40.0.0/16 [110/2] via 10.0.0.26, 02:45:01, GigabitEthernet0/11\n`;
    }
    else if (lower.includes('database summary') || lower.includes('db sum')) {
      output.innerHTML +=
`            OSPF Router with ID (10.0.0.1) (Process ID 1)

                Summary Net Link States (Area 0)

Link ID         ADV Router      Age         Seq#       Checksum
10.10.0.0       10.10.0.1       312         0x80000004 0x009F45
10.20.0.0       10.20.0.1       290         0x80000003 0x00A12C
10.30.0.0       10.30.0.1       180         0x80000002 0x00B810
10.40.0.0       10.40.0.1       244         0x80000003 0x004F3A\n`;
    }
    else if (lower.startsWith('ping')) {
      const target = cmd.split(' ')[1] || '10.20.0.1';
      output.innerHTML +=
`Type escape sequence to abort.
Sending 5, 100-byte ICMP Echos to ${target}, timeout is 2 seconds:
!!!!!
Success rate is 100 percent (5/5), round-trip min/avg/max = 1/1/2 ms\n`;
    }
    else if (lower.startsWith('traceroute') || lower.startsWith('trace')) {
      const target = cmd.split(' ')[1] || '10.20.0.1';
      output.innerHTML +=
`Type escape sequence to abort.
Tracing the route to ${target}
VRF info: (vrf in name/id, vrf out name/id)
  1 10.0.0.14 1 msec 1 msec 1 msec
  2 ${target} 1 msec 1 msec 1 msec\n`;
    }
    else if (lower === 'clear' || lower === 'cls') {
      output.innerHTML = 'Terminal display cleared.\n';
    }
    else if (lower === 'help' || lower === '?') {
      output.innerHTML +=
`Supported Cisco Commands:
  show ip ospf neighbor          - Display active OSPF neighbor adjacencies
  show ip route ospf             - Display OSPF routing table entries
  show ip route                  - Display complete IP routing table
  show ip ospf database summary  - Inspect Type-3 summary LSAs
  ping <ip>                      - Send ICMP echo requests
  traceroute <ip>                - Trace path to destination
  clear                          - Clear terminal screen\n`;
    }
    else {
      output.innerHTML += `% Invalid command "${cmd}". Type "help" or "?" for list of valid commands.\n`;
    }

    output.scrollTop = output.scrollHeight;
  }

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      executeCommand(input.value);
      input.value = '';
    }
  });

  quickBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const cmd = btn.dataset.cmd;
      if (cmd) {
        input.value = cmd;
        executeCommand(cmd);
        input.value = '';
      }
    });
  });
}
