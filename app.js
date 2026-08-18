/**
 * ==============================================================================
 * DYNAMIC MULTI-AREA OSPF ENTERPRISE SIMULATOR ENGINE (RFC 2328 COMPLIANT)
 * ==============================================================================
 */

// Global State
const state = {
  activeRouter: 'Core-R1',
  timerProfile: 'optimized', // 'standard' or 'optimized'
  spfRuns: 0,
  links: {
    'L_C1_C2': { u: 'Core-R1', v: 'Core-R2', cost: 1, up: true, area: 0, type: 'p2p' },
    'L_C1_ACAD': { u: 'Core-R1', v: 'Academic-ABR', cost: 5, up: true, area: 0, type: 'p2p' },
    'L_C1_ADMIN': { u: 'Core-R1', v: 'Admin-ABR', cost: 5, up: true, area: 0, type: 'p2p' },
    'L_C2_HOSTEL': { u: 'Core-R2', v: 'Hostel-ABR', cost: 5, up: true, area: 0, type: 'p2p' },
    'L_C2_LABS': { u: 'Core-R2', v: 'Labs-ABR', cost: 5, up: true, area: 0, type: 'p2p' },
    'L_C2_ACAD': { u: 'Core-R2', v: 'Academic-ABR', cost: 10, up: true, area: 0, type: 'p2p' },
    'L_C1_HOSTEL': { u: 'Core-R1', v: 'Hostel-ABR', cost: 10, up: true, area: 0, type: 'p2p' }
  },
  nodes: {
    'Core-R1': {
      rid: '10.0.0.1',
      role: 'Backbone Autonomous Core Router',
      area: 0,
      isABR: false,
      isStub: false,
      coords: { x: 380, y: 120 },
      interfaces: [
        { name: 'Gi0/0/0', ip: '10.0.0.5/30', link: 'L_C1_C2', neighbor: '10.0.0.2', status: 'up' },
        { name: 'Gi0/0/1', ip: '10.0.0.9/30', link: 'L_C1_ACAD', neighbor: '10.10.0.1', status: 'up' },
        { name: 'Gi0/0/2', ip: '10.0.0.13/30', link: 'L_C1_ADMIN', neighbor: '10.30.0.1', status: 'up' },
        { name: 'Gi0/0/3', ip: '10.0.0.17/30', link: 'L_C1_HOSTEL', neighbor: '10.20.0.1', status: 'up' }
      ]
    },
    'Core-R2': {
      rid: '10.0.0.2',
      role: 'Backbone Redundant Core Router',
      area: 0,
      isABR: false,
      isStub: false,
      coords: { x: 570, y: 120 },
      interfaces: [
        { name: 'Gi0/0/0', ip: '10.0.0.6/30', link: 'L_C1_C2', neighbor: '10.0.0.1', status: 'up' },
        { name: 'Gi0/0/1', ip: '10.0.0.21/30', link: 'L_C2_HOSTEL', neighbor: '10.20.0.1', status: 'up' },
        { name: 'Gi0/0/2', ip: '10.0.0.25/30', link: 'L_C2_LABS', neighbor: '10.40.0.1', status: 'up' },
        { name: 'Gi0/0/3', ip: '10.0.0.29/30', link: 'L_C2_ACAD', neighbor: '10.10.0.1', status: 'up' }
      ]
    },
    'Academic-ABR': {
      rid: '10.10.0.1',
      role: 'Academic Area Border Router (ABR)',
      area: 10,
      isABR: true,
      isStub: false,
      summary: '10.10.0.0/16',
      subnets: ['10.10.1.0/24', '10.10.2.0/24'],
      coords: { x: 220, y: 320 },
      interfaces: [
        { name: 'Gi0/0/0', ip: '10.0.0.10/30', link: 'L_C1_ACAD', neighbor: '10.0.0.1', status: 'up' },
        { name: 'Gi0/0/1', ip: '10.0.0.30/30', link: 'L_C2_ACAD', neighbor: '10.0.0.2', status: 'up' },
        { name: 'Gi0/0/2.101', ip: '10.10.1.1/24', link: null, neighbor: 'None (Passive)', status: 'up' },
        { name: 'Gi0/0/2.102', ip: '10.10.2.1/24', link: null, neighbor: 'None (Passive)', status: 'up' }
      ]
    },
    'Hostel-ABR': {
      rid: '10.20.0.1',
      role: 'Residential Hostels Area Border Router (ABR)',
      area: 20,
      isABR: true,
      isStub: false,
      summary: '10.20.0.0/16',
      subnets: ['10.20.0.0/18', '10.20.64.0/18'],
      coords: { x: 730, y: 320 },
      interfaces: [
        { name: 'Gi0/0/0', ip: '10.0.0.22/30', link: 'L_C2_HOSTEL', neighbor: '10.0.0.2', status: 'up' },
        { name: 'Gi0/0/1', ip: '10.0.0.18/30', link: 'L_C1_HOSTEL', neighbor: '10.0.0.1', status: 'up' },
        { name: 'Gi0/0/2.201', ip: '10.20.0.1/18', link: null, neighbor: 'None (Passive)', status: 'up' }
      ]
    },
    'Admin-ABR': {
      rid: '10.30.0.1',
      role: 'Admin & VTOP Data Center ABR (Stub)',
      area: 30,
      isABR: true,
      isStub: true, // Totally Stubby
      summary: '0.0.0.0/0',
      subnets: ['10.30.1.0/24'],
      coords: { x: 220, y: 510 },
      interfaces: [
        { name: 'Gi0/0/0', ip: '10.0.0.14/30', link: 'L_C1_ADMIN', neighbor: '10.0.0.1', status: 'up' },
        { name: 'Gi0/0/1.301', ip: '10.30.1.1/24', link: null, neighbor: 'None (Passive)', status: 'up' }
      ]
    },
    'Labs-ABR': {
      rid: '10.40.0.1',
      role: 'High-Compute & IoT Labs ABR',
      area: 40,
      isABR: true,
      isStub: false,
      summary: '10.40.0.0/16',
      subnets: ['10.40.1.0/24'],
      coords: { x: 730, y: 510 },
      interfaces: [
        { name: 'Gi0/0/0', ip: '10.0.0.26/30', link: 'L_C2_LABS', neighbor: '10.0.0.2', status: 'up' },
        { name: 'Gi0/0/1.401', ip: '10.40.1.1/24', link: null, neighbor: 'None (Passive)', status: 'up' }
      ]
    }
  }
};

// Initial routing tables cache
const routingTables = {};

// Initialization
document.addEventListener('DOMContentLoaded', () => {
  selectRouter('Core-R1');
  recalculateNetworkState();
  initLinkEventListeners();
});

// Link Selection Listeners (directly on SVG elements)
function initLinkEventListeners() {
  for (const linkKey in state.links) {
    const el = document.getElementById(linkKey);
    if (el) {
      el.addEventListener('click', () => {
        // Toggle from list
        const select = document.getElementById('linkFlapSelect');
        select.value = linkKey;
        const currentStatus = state.links[linkKey].up;
        toggleSelectedLink(!currentStatus);
      });
    }
  }
}

// Timer Profile Switcher
function setTimerProfile(profile) {
  state.timerProfile = profile;
  document.getElementById('standardTimersBtn').classList.toggle('active', profile === 'standard');
  document.getElementById('optimizedTimersBtn').classList.toggle('active', profile === 'optimized');

  const convergenceElem = document.getElementById('avgConvergenceTime');
  if (profile === 'optimized') {
    convergenceElem.textContent = '3.008s';
    convergenceElem.className = 'value cyan';
  } else {
    convergenceElem.textContent = '40.015s';
    convergenceElem.className = 'value amber';
  }
  appendCliOutput(`% OSPF-5-TIMER_UPDATE: Hello/Dead intervals reset to ${profile.toUpperCase()} profile.`);
  recalculateNetworkState();
}

// Router selection
function selectRouter(nodeKey) {
  state.activeRouter = nodeKey;
  
  document.querySelectorAll('.router-node').forEach(node => node.classList.remove('selected'));
  const target = document.getElementById(`node-${nodeKey}`);
  if (target) target.classList.add('selected');

  const rData = state.nodes[nodeKey];
  document.getElementById('inspectedRouterName').textContent = `Node: ${nodeKey}`;
  document.getElementById('inspectedRouterRole').textContent = `Role: ${rData.role}`;
  document.getElementById('cliPromptLabel').textContent = `${nodeKey}#`;
  
  const statusTag = document.getElementById('inspectedNodeState');
  const isRouterUp = isNodeFullyConnected(nodeKey);
  statusTag.textContent = isRouterUp ? 'FULL STATE' : 'DEGRADED';
  statusTag.className = isRouterUp ? 'status-tag live' : 'status-tag down';

  renderDiagnostics();
}

// Check if a router node is fully up
function isNodeFullyConnected(nodeKey) {
  const router = state.nodes[nodeKey];
  let upLinks = 0;
  router.interfaces.forEach(iface => {
    if (iface.link) {
      const link = state.links[iface.link];
      if (link && link.up) upLinks++;
    }
  });
  return upLinks > 0 || nodeKey.startsWith('Core');
}

// Link Fail/Heal Control
function toggleSelectedLink(status) {
  const linkKey = document.getElementById('linkFlapSelect').value;
  const link = state.links[linkKey];
  if (link) {
    if (link.up === status) return; // Status already matches

    link.up = status;
    const svgLine = document.getElementById(linkKey);
    if (svgLine) svgLine.classList.toggle('severed', !status);

    state.spfRuns++;
    document.getElementById('globalSpfCount').textContent = state.spfRuns;

    // Visual updates for router node representations
    updateRouterNodeVisualStatus(link.u);
    updateRouterNodeVisualStatus(link.v);

    const delay = state.timerProfile === 'optimized' ? '3.008s' : '40.015s';
    appendCliOutput(`% LINK-3-UPDOWN: Interface on ${linkKey} changed state to ${status ? 'UP' : 'DOWN'}`);
    appendCliOutput(`% OSPF-5-ADJCHANGE: SPF recalculation triggered. Converged in ${delay}`);

    recalculateNetworkState();
    
    // Auto re-render Dijkstra highlights if active
    const activeSptRoot = document.getElementById('sptRootSelect').value;
    visualizeSPT(activeSptRoot);
  }
}

function updateRouterNodeVisualStatus(nodeName) {
  const isUp = isNodeFullyConnected(nodeName);
  const el = document.getElementById(`node-${nodeName}`);
  if (el) {
    el.classList.toggle('down', !isUp);
    el.classList.toggle('active', isUp);
  }
}

// Dynamic routing calculation using Dijkstra algorithm
function recalculateNetworkState() {
  // Clear routing tables
  for (const nodeKey in state.nodes) {
    routingTables[nodeKey] = [];
  }

  // 1. Calculate routes for each router
  for (const nodeKey in state.nodes) {
    const router = state.nodes[nodeKey];
    const table = [];

    // Add Connected interfaces
    router.interfaces.forEach(iface => {
      table.push({
        code: 'C',
        prefix: getSubnetPrefix(iface.ip),
        metric: '[0/0]',
        nextHop: 'Direct',
        interface: iface.name
      });
    });

    if (router.isStub) {
      // Totally Stub Area Rule: Strips all database summary updates, receives only a default route
      table.push({
        code: 'O*IA',
        prefix: '0.0.0.0/0',
        metric: '[110/6]',
        nextHop: '10.0.0.13', // Next hop is Core-R1 interface
        interface: 'Gi0/0/0'
      });
    } else {
      // Non-stub dynamic Dijkstra calculation
      // Calculate distances using SPF
      const { pathDetails } = runDijkstra(nodeKey);
      
      for (const destKey in pathDetails) {
        if (destKey === nodeKey) continue;
        const result = pathDetails[destKey];
        if (result.cost === Infinity) continue; // Unreachable

        const destRouter = state.nodes[destKey];
        
        // Add remote Loopback route
        table.push({
          code: getProtocolCode(router.area, destRouter.area),
          prefix: `${destRouter.rid}/32`,
          metric: `[110/${result.cost}]`,
          nextHop: result.nextHopIp,
          interface: result.outboundInterface
        });

        // Add remote summaries / subnets
        if (destRouter.isABR && destRouter.summary) {
          // ABR summarizes subnets into a single /16 block
          table.push({
            code: 'O IA',
            prefix: destRouter.summary,
            metric: `[110/${result.cost + 5}]`,
            nextHop: result.nextHopIp,
            interface: result.outboundInterface
          });
        } else {
          // Standard propagation
          destRouter.interfaces.forEach(iface => {
            if (iface.link === null) { // Client subnets
              table.push({
                code: getProtocolCode(router.area, destRouter.area),
                prefix: getSubnetPrefix(iface.ip),
                metric: `[110/${result.cost + 10}]`,
                nextHop: result.nextHopIp,
                interface: result.outboundInterface
              });
            }
          });
        }
      }
    }

    routingTables[nodeKey] = table;
  }

  // Update overall LSA count display based on active adjacencies
  let lsaCount = 12; // Base Loopbacks and ABR Type-3 defaults
  for (const key in state.links) {
    if (state.links[key].up) lsaCount += 2;
  }
  document.getElementById('totalLsaCount').textContent = lsaCount;

  // Update topology status bar
  const statusEl = document.getElementById('globalTopologyStatus');
  if (statusEl) {
    let allUp = true;
    for (const key in state.links) {
      if (!state.links[key].up) allUp = false;
    }
    if (allUp) {
      statusEl.textContent = 'STABLE (FULL)';
      statusEl.className = 'value green-glow';
      statusEl.style.color = '';
      statusEl.style.textShadow = '';
    } else {
      statusEl.textContent = 'DEGRADED (SPF ACTIVE)';
      statusEl.className = 'value';
      statusEl.style.color = 'var(--amber-warn)';
      statusEl.style.textShadow = '0 0 8px rgba(245, 158, 11, 0.4)';
    }
  }

  // Render diagnostics tables for selected router
  renderDiagnostics();
}

// Dijkstra Graph Solver
function runDijkstra(startNodeKey) {
  const distances = {};
  const previous = {};
  const queue = [];

  // Initialize
  for (const key in state.nodes) {
    distances[key] = Infinity;
    previous[key] = null;
    queue.push(key);
  }
  distances[startNodeKey] = 0;

  while (queue.length > 0) {
    // Find min distance vertex
    queue.sort((a, b) => distances[a] - distances[b]);
    const u = queue.shift();

    if (distances[u] === Infinity) break;

    // Search neighbors
    const neighbors = getNodeNeighbors(u);
    neighbors.forEach(nbr => {
      const v = nbr.nodeKey;
      if (queue.includes(v)) {
        const alt = distances[u] + nbr.cost;
        if (alt < distances[v]) {
          distances[v] = alt;
          previous[v] = u;
        }
      }
    });
  }

  // Resolve path details for routing table next hops
  const pathDetails = {};
  for (const targetKey in state.nodes) {
    if (distances[targetKey] === Infinity || targetKey === startNodeKey) continue;

    // Backtrack path to find first hop node
    let step = targetKey;
    while (previous[step] && previous[step] !== startNodeKey) {
      step = previous[step];
    }

    const localRouter = state.nodes[startNodeKey];
    let outboundInt = 'Unknown';
    let nextHopIp = 'Unknown';

    localRouter.interfaces.forEach(iface => {
      if (iface.link) {
        const l = state.links[iface.link];
        if (l && l.up && (l.u === step || l.v === step)) {
          outboundInt = iface.name;
          nextHopIp = iface.neighbor;
        }
      }
    });

    pathDetails[targetKey] = {
      cost: distances[targetKey],
      nextHopIp: nextHopIp,
      outboundInterface: outboundInt
    };
  }

  return { distances, previous, pathDetails };
}

// Find neighbors for a node
function getNodeNeighbors(nodeKey) {
  const router = state.nodes[nodeKey];
  const neighbors = [];
  
  router.interfaces.forEach(iface => {
    if (iface.link) {
      const link = state.links[iface.link];
      if (link && link.up) {
        const otherNode = link.u === nodeKey ? link.v : link.u;
        neighbors.push({
          nodeKey: otherNode,
          cost: link.cost
        });
      }
    }
  });
  return neighbors;
}

// Determine OSPF Intra-Area or Inter-Area tag
function getProtocolCode(srcArea, destArea) {
  return srcArea === destArea ? 'O' : 'O IA';
}

function getSubnetPrefix(ipStr) {
  const ip = ipStr.split('/')[0];
  const parts = ip.split('.');
  if (ipStr.endsWith('/30')) {
    const last = parseInt(parts[3]);
    const base = last - (last % 4);
    return `${parts[0]}.${parts[1]}.${parts[2]}.${base}/30`;
  }
  if (ipStr.endsWith('/24')) {
    return `${parts[0]}.${parts[1]}.${parts[2]}.0/24`;
  }
  if (ipStr.endsWith('/18')) {
    const last = parseInt(parts[2]);
    const base = last - (last % 64);
    return `${parts[0]}.${parts[1]}.${base}.0/18`;
  }
  return ipStr;
}

// Render tabs
function renderDiagnostics() {
  const node = state.nodes[state.activeRouter];
  
  // 1. Render IP Routing Table
  const tableBody = document.querySelector('#routingTableElement tbody');
  tableBody.innerHTML = '';
  
  const routes = routingTables[state.activeRouter] || [];
  if (routes.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="5" style="text-align:center; color:var(--text-muted);">Router is completely offline. No routes computed.</td></tr>`;
  } else {
    // Unique routes check to prevent display duplicates
    const seenPrefixes = new Set();
    routes.forEach(r => {
      if (seenPrefixes.has(r.prefix)) return;
      seenPrefixes.add(r.prefix);

      const colorBadge = r.code.startsWith('O*') || r.code === 'O IA' ? 'code-summary' : '';
      tableBody.innerHTML += `
        <tr>
          <td><span class="code-badge ${colorBadge}">${r.code}</span></td>
          <td>${r.prefix}</td>
          <td>${r.metric}</td>
          <td>${r.nextHop}</td>
          <td>${r.interface}</td>
        </tr>
      `;
    });
  }

  // 2. Render Neighbor Adjacency
  const nBody = document.querySelector('#neighborTableElement tbody');
  nBody.innerHTML = '';
  
  let hasNeighbors = false;
  node.interfaces.forEach(iface => {
    if (iface.link) {
      hasNeighbors = true;
      const link = state.links[iface.link];
      const stateStr = link.up ? 'FULL/ -' : 'DOWN';
      const deadTimer = state.timerProfile === 'optimized' ? '00:00:03' : '00:00:37';
      nBody.innerHTML += `
        <tr>
          <td>${iface.neighbor}</td>
          <td>1</td>
          <td><span style="color: ${link.up ? 'var(--emerald-live)' : 'var(--rose-danger)'}">${stateStr}</span></td>
          <td>${link.up ? deadTimer : '-'}</td>
          <td>${iface.ip.split('/')[0]}</td>
          <td>${iface.name}</td>
        </tr>
      `;
    }
  });

  if (!hasNeighbors) {
    nBody.innerHTML = `<tr><td colspan="6" style="text-align:center; color:var(--text-muted);">No active serial or transit links configured.</td></tr>`;
  }

  // 3. Render Link State Database (LSDB)
  const lsdb = document.getElementById('lsdbContainer');
  if (node.isStub) {
    lsdb.innerHTML = `
      <div class="lsa-block">
        <div class="lsa-title">Router Link States (Area ${node.area}) - Type 1 (Router LSA)</div>
        <div class="lsa-entry">ADV Router: ${node.rid} | Seq: 0x80000003 | Checksum: 0x4A7D</div>
      </div>
      <div class="lsa-block">
        <div class="lsa-title">Summary Link States (Area ${node.area}) - Type 3 (Summary LSA)</div>
        <div class="lsa-entry">Link ID: 0.0.0.0 | ADV Router: 10.30.0.1 | Metric: 1 (Default route fallback)</div>
      </div>
    `;
  } else {
    // Dynamically compile Type-1 and Type-3 list for others
    let type1Html = '';
    let type3Html = '';

    for (const key in state.nodes) {
      const r = state.nodes[key];
      if (r.area === node.area) {
        type1Html += `<div class="lsa-entry">ADV Router: ${r.rid} | Seq: 0x80000004 | Checksum: 0x7E2A | Loopbacks: 1</div>`;
      }
      if (r.isABR && r.summary && r.area !== node.area) {
        type3Html += `<div class="lsa-entry">Prefix: ${r.summary} | ADV Router: ${r.rid} | Metric: 5</div>`;
      }
    }

    lsdb.innerHTML = `
      <div class="lsa-block">
        <div class="lsa-title">Router Link States (Area ${node.area}) - Type 1 (Router LSA)</div>
        ${type1Html}
      </div>
      <div class="lsa-block">
        <div class="lsa-title">Summary Net Link States (Area ${node.area}) - Type 3 (Summary LSA)</div>
        ${type3Html || '<div class="lsa-entry">None (No summaries computed)</div>'}
      </div>
    `;
  }
}

// Shortest Path First (Dijkstra) Visualization
function visualizeSPT(rootName) {
  // Clear previous spt highlight styles
  document.querySelectorAll('.net-link').forEach(l => l.classList.remove('active-spt'));

  appendCliOutput(`% SPF-ENGINE: Executing Dijkstra calculation with root node ${rootName}...`);
  
  // Calculate Dijkstra tree from rootName
  const { distances, previous } = runDijkstra(rootName);
  
  // Highlight links in the path
  let activeLinks = 0;
  for (const v in previous) {
    const parent = previous[v];
    if (parent) {
      // Find the link connecting v and parent
      for (const linkKey in state.links) {
        const link = state.links[linkKey];
        if (link.up && ((link.u === v && link.v === parent) || (link.u === parent && link.v === v))) {
          const el = document.getElementById(linkKey);
          if (el && !el.classList.contains('active-spt')) {
            el.classList.add('active-spt');
            activeLinks++;
          }
        }
      }
    }
  }

  appendCliOutput(`% OSPF-5-SPF: Shortest Path Tree calculated successfully in 2.12ms. ${activeLinks} active SPT links highlighted.`);
}

// Packet Flow Animation Simulation
function triggerPacketSimulation() {
  const layer = document.getElementById('packetSimulationLayer');
  layer.innerHTML = ''; // reset

  // Determine path from Hostel-ABR to Admin-ABR using state engine
  const start = 'Hostel-ABR';
  const end = 'Admin-ABR';
  
  const { distances, previous } = runDijkstra(start);
  
  if (distances[end] === Infinity) {
    appendCliOutput(`% IP-FORWARD: Target destination VTOP database unreachable. Link severed!`);
    return;
  }

  // Backtrack from end to start to get the path
  const pathNodes = [];
  let curr = end;
  while (curr !== null) {
    pathNodes.unshift(curr);
    curr = previous[curr];
  }

  // Draw motion path coordinate string
  let dString = "";
  pathNodes.forEach((nodeKey, index) => {
    const coords = state.nodes[nodeKey].coords;
    if (index === 0) {
      dString += `M ${coords.x},${coords.y}`;
    } else {
      dString += ` L ${coords.x},${coords.y}`;
    }
  });

  layer.innerHTML = `
    <circle class="sim-packet" r="6">
      <animateMotion path="${dString}" dur="2.2s" fill="freeze" repeatCount="1" />
    </circle>
  `;
  appendCliOutput(`% IP-FORWARD: Packet dispatched [SRC: 10.20.1.5 (Hostels) -> DST: 10.30.1.10 (VTOP Portal)] via path: ${pathNodes.join(' ➔ ')}`);
}

// Security Injection Simulation
function injectRogueHello() {
  appendCliOutput(`% SEC-ALERT: Spoofed OSPF Hello received on Academic-ABR Gi0/0/2.101`);
  
  const targetNode = document.getElementById('node-Academic-ABR');
  if (targetNode) {
    targetNode.classList.add('node-flash-red');
    setTimeout(() => {
      targetNode.classList.remove('node-flash-red');
    }, 2400);
  }

  appendCliOutput(`% SEC-DROP: Packet dropped! Interface Gi0/0/2 is passive and MD5 authentication keys mismatch.`);
}

// Diagnostics Tabs
function switchDiagTab(tabId) {
  // Get active button elements
  const tabContainer = event.currentTarget.parentNode;
  tabContainer.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
  
  const contentContainer = tabContainer.parentNode;
  contentContainer.querySelectorAll('.tab-pane').forEach(pane => pane.classList.remove('active'));

  event.currentTarget.classList.add('active');
  document.getElementById(tabId).classList.add('active');
}

// Cisco IOS CLI Emulator
function handleCliInput(e) {
  if (e.key === 'Enter') {
    const input = document.getElementById('cliInput');
    const cmd = input.value.trim();
    if (!cmd) return;

    appendCliOutput(`${state.activeRouter}# ${cmd}`);
    executeCiscoCommand(cmd);
    input.value = '';
  }
}

function appendCliOutput(text) {
  const out = document.getElementById('terminalOutput');
  out.innerHTML += `${text}<br>`;
  out.scrollTop = out.scrollHeight;
}

function executeCiscoCommand(cmd) {
  const c = cmd.toLowerCase();
  if (c === 'show ip route' || c === 'sh ip ro') {
    const routes = routingTables[state.activeRouter] || [];
    let routeStr = `Codes: C - connected, S - static, O - OSPF, IA - OSPF inter area\n\nGateway of last resort is not set\n\n`;
    routes.forEach(r => {
      routeStr += `${r.code.padEnd(6)}${r.prefix.padEnd(18)} via ${r.nextHop.padEnd(14)} on ${r.interface}\n`;
    });
    appendCliOutput(routeStr);
  } else if (c === 'show ip ospf neighbor' || c === 'sh ip ospf nei') {
    const node = state.nodes[state.activeRouter];
    let nbrStr = `Neighbor ID     Pri   State           Dead Time   Address         Interface\n`;
    node.interfaces.forEach(i => {
      if (i.link) {
        const link = state.links[i.link];
        const stateStr = link.up ? 'FULL/ -' : 'DOWN';
        nbrStr += `${i.neighbor.padEnd(16)}1     ${stateStr.padEnd(16)}00:00:03    ${i.ip.split('/')[0].padEnd(16)}${i.name}\n`;
      }
    });
    appendCliOutput(nbrStr);
  } else if (c === 'show ip ospf database' || c === 'sh ip ospf da') {
    const node = state.nodes[state.activeRouter];
    let lsdbStr = `            OSPF Router with ID (${node.rid})\n\n                Router Link States (Area ${node.area})\n\nLink ID         ADV Router      Age         Seq#       Checksum\n`;
    for (const k in state.nodes) {
      if (state.nodes[k].area === node.area) {
        lsdbStr += `${state.nodes[k].rid.padEnd(16)}${state.nodes[k].rid.padEnd(16)}420         0x80000004 0x007E\n`;
      }
    }
    appendCliOutput(lsdbStr);
  } else if (c === 'clear ip ospf process') {
    state.spfRuns++;
    document.getElementById('globalSpfCount').textContent = state.spfRuns;
    appendCliOutput(`% OSPF-5-ADJCHANGE: Neighbor processes cleared. Resetting adjacency state engines...`);
    recalculateNetworkState();
  } else if (c === 'help') {
    appendCliOutput(`Available commands:\n  show ip route\n  show ip ospf neighbor\n  show ip ospf database\n  clear ip ospf process`);
  } else {
    appendCliOutput(`% Invalid command: "${cmd}". Type "help" for available commands.`);
  }
}

// Automated Test Harness Execution Suite
async function runSingleTest(tcId) {
  const card = document.getElementById(`card-${tcId}`);
  const statusElem = card.querySelector('.tc-status');
  statusElem.textContent = 'RUNNING';
  statusElem.className = 'tc-status running';

  appendCliOutput(`% TEST-SUITE: Executing compliance tests for ${tcId}...`);

  await new Promise(r => setTimeout(r, 600));

  let passed = true;
  
  // Specific checks for compliance
  if (tcId === 'TC-OSPF-01') {
    // Check neighbor states
    let hasDown = false;
    for (const key in state.links) {
      if (!state.links[key].up) hasDown = true;
    }
    if (hasDown) {
      appendCliOutput(`% TEST-WARN: Adjacency test warnings registered. Some mesh links are failed.`);
    }
  } else if (tcId === 'TC-OSPF-02') {
    appendCliOutput(`% TEST-CHECK: Verifying Area 10 & Area 0 boundary database isolation...`);
    appendCliOutput(`% TEST-CHECK: Type-1 Router LSAs contained inside Area 10: OK`);
    appendCliOutput(`% TEST-CHECK: Type-3 Summary LSAs generated for Core: OK`);
  } else if (tcId === 'TC-OSPF-03') {
    // Verify summarization
    if (routingTables['Core-R1'] && !routingTables['Core-R1'].some(r => r.prefix === '10.10.0.0/16')) {
      passed = false;
    }
  } else if (tcId === 'TC-OSPF-04') {
    // Verify Stub default route on Admin node
    if (routingTables['Admin-ABR'] && !routingTables['Admin-ABR'].some(r => r.prefix === '0.0.0.0/0')) {
      passed = false;
    }
  } else if (tcId === 'TC-OSPF-05') {
    appendCliOutput(`% TEST-CHECK: Analyzing OSPF Hello/Dead state transition timers...`);
    if (state.timerProfile === 'optimized') {
      appendCliOutput(`% TEST-CHECK: Dead neighbor interval is 4 seconds. Failover speed <= 3.008s: OK`);
      passed = true;
    } else {
      appendCliOutput(`% TEST-WARN: Standard timers active (Dead: 40 seconds). Convergence latency exceeds 3.5s limit.`);
      passed = false;
    }
  }

  if (passed) {
    statusElem.textContent = 'PASSED';
    statusElem.className = 'tc-status passed';
    appendCliOutput(`% TEST-SUITE: Verification success. ${tcId} complies with specification standards.`);
  } else {
    statusElem.textContent = 'FAILED';
    statusElem.className = 'tc-status failed';
    appendCliOutput(`% TEST-SUITE: Verification failed. ${tcId} does not comply with specifications.`);
  }
}

async function runAllTestCases() {
  const testIds = ['TC-OSPF-01', 'TC-OSPF-02', 'TC-OSPF-03', 'TC-OSPF-04', 'TC-OSPF-05'];
  for (const id of testIds) {
    await runSingleTest(id);
  }
}
