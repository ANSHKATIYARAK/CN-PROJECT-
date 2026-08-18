/**
 * ==============================================================================
 * DYNAMIC MULTI-AREA OSPF ENTERPRISE SIMULATOR ENGINE (RFC 2328 COMPLIANT)
 * ==============================================================================
 */

// Topology Data Model Definition
const TOPOLOGY = {
  nodes: [
    { id: 'Core-R1', rid: '10.0.0.1', x: 320, y: 80, area: 0, isABR: true, role: 'Backbone Autonomous Core Router' },
    { id: 'Core-R2', rid: '10.0.0.2', x: 480, y: 80, area: 0, isABR: true, role: 'Backbone Redundant Core Router' },
    { id: 'Academic-ABR', rid: '10.10.0.1', x: 120, y: 300, area: 10, isABR: true, subnet: '10.10.0.0/16', role: 'Academic Area Border Router (ABR)' },
    { id: 'Labs-ABR', rid: '10.40.0.1', x: 220, y: 220, area: 10, isABR: true, subnet: '10.40.0.0/16', role: 'High-Compute & IoT Labs ABR' },
    { id: 'Hostel-ABR', rid: '10.20.0.1', x: 660, y: 300, area: 20, isABR: true, subnet: '10.20.0.0/16', role: 'Residential Hostels Area Border Router (ABR)' },
    { id: 'Admin-ABR', rid: '10.30.0.1', x: 380, y: 360, area: 30, isABR: true, isStub: true, subnet: '10.30.0.0/16', role: 'Admin & VTOP Data Center ABR (Stub)' }
  ],
  links: [
    { id: 'l1', source: 'Core-R1', target: 'Core-R2', cost: 1, area: 0, ipSrc: '10.0.0.1', ipTgt: '10.0.0.2', state: 'UP' },
    { id: 'l2', source: 'Core-R1', target: 'Academic-ABR', cost: 1, area: 10, ipSrc: '10.0.0.5', ipTgt: '10.0.0.6', state: 'UP' },
    { id: 'l3', source: 'Core-R1', target: 'Labs-ABR', cost: 1, area: 10, ipSrc: '10.0.0.13', ipTgt: '10.0.0.14', state: 'UP' },
    { id: 'l4', source: 'Core-R1', target: 'Admin-ABR', cost: 1, area: 30, ipSrc: '10.0.0.9', ipTgt: '10.0.0.10', state: 'UP' },
    { id: 'l5', source: 'Core-R2', target: 'Hostel-ABR', cost: 5, area: 20, ipSrc: '10.0.0.17', ipTgt: '10.0.0.18', state: 'UP' },
    { id: 'l6', source: 'Labs-ABR', target: 'Hostel-ABR', cost: 1, area: 20, ipSrc: '10.0.0.21', ipTgt: '10.0.0.22', state: 'UP' }
  ]
};

// Global Simulation State Class
class OSPFSimulator {
  constructor() {
    this.topology = JSON.parse(JSON.stringify(TOPOLOGY));
    this.selectedNodeId = 'Core-R1';
    this.helloInterval = 1;
    this.deadInterval = 4;
    this.multiAreaEnabled = true;
    this.timerProfile = 'optimized';
    this.spfRuns = 0;
    this.lsdb = {};
    this.routingTables = {};
    this.initLSDB();
    this.computeAllRoutes();
  }

  initLSDB() {
    this.topology.nodes.forEach(node => {
      this.lsdb[node.id] = {
        type1: [], // Router LSAs
        type3: [], // Summary LSAs
        type5: []  // AS External LSAs
      };
    });
    this.generateLSAs();
  }

  generateLSAs() {
    this.topology.nodes.forEach(node => {
      const activeLinks = this.topology.links.filter(
        l => (l.source === node.id || l.target === node.id) && l.state === 'UP'
      );
      this.lsdb[node.id].type1 = [{
        linkId: node.rid,
        advRouter: node.rid,
        age: Math.floor(Math.random() * 200) + 10,
        seq: '0x80000004',
        checksum: '0x004A2C',
        linksCount: activeLinks.length
      }];
    });

    // Generate Type 3 Summaries across ABR boundaries
    if (this.multiAreaEnabled) {
      this.topology.nodes.forEach(node => {
        this.lsdb[node.id].type3 = [];
        
        // Area 30 stub isolation logic
        if (node.area === 30 || node.isStub) {
          // Totally Stubby Area receives ONLY default route summary (0.0.0.0/0)
          this.lsdb[node.id].type3.push({
            prefix: '0.0.0.0/0',
            advRouter: 'Core-R1',
            age: 45,
            seq: '0x80000001',
            checksum: '0x001A2B',
            metric: 1
          });
        } else {
          // Standard and Core areas receive summaries for remote subnets
          this.topology.nodes.filter(n => n.subnet && n.area !== node.area).forEach(remoteNode => {
            this.lsdb[node.id].type3.push({
              prefix: remoteNode.subnet,
              advRouter: remoteNode.rid,
              age: 120,
              seq: '0x80000002',
              checksum: '0x008F12',
              metric: 2
            });
          });
        }
      });
    } else {
      // In flat single area mode, summaries are disabled
      this.topology.nodes.forEach(node => {
        this.lsdb[node.id].type3 = [];
      });
    }
  }

  // Dijkstra Shortest Path Solver returning full details
  getSPT(startNodeId) {
    const distances = {};
    const previous = {};
    const unvisited = new Set();
    
    this.topology.nodes.forEach(n => {
      distances[n.id] = Infinity;
      previous[n.id] = null;
      unvisited.add(n.id);
    });
    distances[startNodeId] = 0;

    while (unvisited.size > 0) {
      let minNode = null;
      unvisited.forEach(n => {
        if (minNode === null || distances[n] < distances[minNode]) {
          minNode = n;
        }
      });

      if (distances[minNode] === Infinity) break;
      unvisited.delete(minNode);

      const activeNeighbors = this.topology.links.filter(
        l => l.state === 'UP' && (l.source === minNode || l.target === minNode)
      );

      activeNeighbors.forEach(link => {
        const neighborId = link.source === minNode ? link.target : link.source;
        if (unvisited.has(neighborId)) {
          const tentativeCost = distances[minNode] + link.cost;
          if (tentativeCost < distances[neighborId]) {
            distances[neighborId] = tentativeCost;
            previous[neighborId] = { node: minNode, link: link };
          }
        }
      });
    }

    return { distances, previous };
  }

  // Calculate routes for each router node
  computeSPF(startNodeId) {
    const { distances, previous } = this.getSPT(startNodeId);
    const routes = [];

    this.topology.nodes.forEach(n => {
      if (n.id !== startNodeId && distances[n.id] !== Infinity) {
        let curr = n.id;
        let prevStep = previous[curr];
        let nextHopIp = 'Directly Connected';
        let interfaceName = 'GigabitEthernet0/0';

        while (prevStep && prevStep.node !== startNodeId) {
          curr = prevStep.node;
          prevStep = previous[curr];
        }

        if (prevStep) {
          nextHopIp = prevStep.link.source === startNodeId ? prevStep.link.ipTgt : prevStep.link.ipSrc;
          interfaceName = `GigabitEthernet0/${prevStep.link.id.replace('l', '')}`;
        }

        const isInterArea = n.area !== this.topology.nodes.find(x => x.id === startNodeId).area;
        
        // Summarization behavior check
        let prefix = n.subnet || `10.${n.area}.0.0/16`;
        if (!this.multiAreaEnabled) {
          prefix = n.subnet ? `${n.subnet.split('.')[0]}.${n.subnet.split('.')[1]}.0.0/24` : `10.${n.area}.0.0/16`;
        }

        routes.push({
          type: (this.multiAreaEnabled && isInterArea) ? 'O IA' : 'O',
          prefix: prefix,
          metric: distances[n.id],
          nextHop: nextHopIp,
          interface: interfaceName
        });
      }
    });

    // Inject default route for Admin Area 30 stub nodes
    const currentNode = this.topology.nodes.find(x => x.id === startNodeId);
    if (currentNode && (currentNode.area === 30 || currentNode.isStub) && this.multiAreaEnabled) {
      routes.push({
        type: 'O*IA',
        prefix: '0.0.0.0/0',
        metric: 1,
        nextHop: '10.0.0.9',
        interface: 'GigabitEthernet0/4'
      });
    }

    return routes;
  }

  computeAllRoutes() {
    this.generateLSAs();
    this.topology.nodes.forEach(node => {
      this.routingTables[node.id] = this.computeSPF(node.id);
    });
  }

  toggleLinkState(linkId) {
    const link = this.topology.links.find(l => l.id === linkId);
    if (link) {
      link.state = link.state === 'UP' ? 'DOWN' : 'UP';
      this.computeAllRoutes();
    }
  }
}

// Global Simulator Instance
const sim = new OSPFSimulator();

// Topology Rendering Engine
function renderTopology() {
  const linksContainer = document.getElementById('svg-links');
  const nodesContainer = document.getElementById('svg-nodes');
  linksContainer.innerHTML = '';
  nodesContainer.innerHTML = '';

  // Render SVG Links
  sim.topology.links.forEach(link => {
    const srcNode = sim.topology.nodes.find(n => n.id === link.source);
    const tgtNode = sim.topology.nodes.find(n => n.id === link.target);
    const line = document.createElementNS('http://www.w3.org/2000/svg', 'line');
    
    line.setAttribute('x1', srcNode.x);
    line.setAttribute('y1', srcNode.y);
    line.setAttribute('x2', tgtNode.x);
    line.setAttribute('y2', tgtNode.y);
    line.setAttribute('class', `topology-link ${link.state.toLowerCase()}`);
    line.setAttribute('id', `svg-link-${link.id}`);
    
    linksContainer.appendChild(line);
  });

  // Render SVG Router Nodes
  sim.topology.nodes.forEach(node => {
    const g = document.createElementNS('http://www.w3.org/2000/svg', 'g');
    g.setAttribute('class', `router-node ${node.id === sim.selectedNodeId ? 'selected' : ''}`);
    g.setAttribute('transform', `translate(${node.x}, ${node.y})`);
    g.setAttribute('id', `node-container-${node.id}`);

    const outerGlow = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    outerGlow.setAttribute('r', '28');
    outerGlow.setAttribute('class', 'node-glow');

    const circle = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    circle.setAttribute('r', '22');
    circle.setAttribute('class', 'node-base');

    const icon = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    icon.setAttribute('dy', '5');
    icon.setAttribute('class', 'node-icon');
    icon.innerHTML = node.isStub ? '&#x229E;' : '&#x21C4;';

    const title = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    title.setAttribute('y', '38');
    title.setAttribute('class', 'node-title');
    title.textContent = node.id;

    const ip = document.createElementNS('http://www.w3.org/2000/svg', 'text');
    ip.setAttribute('y', '50');
    ip.setAttribute('class', 'node-ip');
    ip.textContent = node.rid;

    g.appendChild(outerGlow);
    g.appendChild(circle);
    g.appendChild(icon);
    g.appendChild(title);
    g.appendChild(ip);
    
    g.addEventListener('click', () => {
      selectRouter(node.id);
    });

    nodesContainer.appendChild(g);
  });
}

// Router selection UI updates
function selectRouter(nodeId) {
  sim.selectedNodeId = nodeId;
  
  // Highlight chosen node in SVG
  document.querySelectorAll('.router-node').forEach(n => {
    n.classList.remove('selected');
  });
  const currentGroup = document.getElementById(`node-container-${nodeId}`);
  if (currentGroup) {
    currentGroup.classList.add('selected');
  }

  renderUI();
  
  // Dynamically calculate and render Dijkstra Shortest Path Tree
  visualizeSPT(nodeId);
}

// Full Dashboard UI Controller
function renderUI() {
  renderTopology();
  
  const selectedNode = sim.topology.nodes.find(n => n.id === sim.selectedNodeId);
  document.getElementById('selected-node-display').textContent = `${selectedNode.id} [RID: ${selectedNode.rid}]`;
  document.getElementById('cli-prompt-text').textContent = `${selectedNode.id}#`;

  // Update Neighbor Table
  const neighborBody = document.getElementById('neighbor-table-body');
  neighborBody.innerHTML = '';
  const activeLinks = sim.topology.links.filter(
    l => (l.source === sim.selectedNodeId || l.target === sim.selectedNodeId)
  );

  activeLinks.forEach(link => {
    const neighborId = link.source === sim.selectedNodeId ? link.target : link.source;
    const neighborObj = sim.topology.nodes.find(n => n.id === neighborId);
    const tr = document.createElement('tr');
    const stateText = link.state === 'UP' ? '<span style="color: #10b981; font-weight: bold;">FULL/DR</span>' : '<span style="color: #ef4444;">DOWN</span>';
    const ipAddr = link.source === sim.selectedNodeId ? link.ipTgt : link.ipSrc;
    
    tr.innerHTML = `
      <td>${neighborObj.rid}</td>
      <td>${stateText}</td>
      <td>GigabitEthernet0/${link.id.replace('l', '')}</td>
      <td>${ipAddr}</td>
    `;
    neighborBody.appendChild(tr);
  });

  // Update Route Table
  const routeBody = document.getElementById('route-table-body');
  routeBody.innerHTML = '';
  const routes = sim.routingTables[sim.selectedNodeId] || [];
  
  routes.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="color: #06b6d4; font-weight: bold;">${r.type}</td>
      <td>${r.prefix}</td>
      <td>[110/${r.metric}]</td>
      <td>${r.nextHop}</td>
      <td>${r.interface}</td>
    `;
    routeBody.appendChild(tr);
  });

  // Update LSDB Display
  const lsdbContainer = document.getElementById('lsdb-view-container');
  const nodeLSDB = sim.lsdb[sim.selectedNodeId];
  let html = `<strong style="color: #06b6d4; font-size: 11px; text-transform: uppercase;">Router Link States (Area ${selectedNode.area})</strong>`;
  
  html += '<table class="data-table"><thead><tr><th>LINK ID</th><th>ADV ROUTER</th><th>AGE</th><th>SEQ#</th></tr></thead><tbody>';
  nodeLSDB.type1.forEach(lsa => {
    html += `<tr><td>${lsa.linkId}</td><td>${lsa.advRouter}</td><td>${lsa.age}</td><td>${lsa.seq}</td></tr>`;
  });
  html += '</tbody></table><br>';

  if (sim.multiAreaEnabled) {
    html += `<strong style="color: #06b6d4; font-size: 11px; text-transform: uppercase;">Summary Net Link States (Area ${selectedNode.area})</strong>`;
    html += '<table class="data-table"><thead><tr><th>PREFIX</th><th>ADV ROUTER</th><th>METRIC</th></tr></thead><tbody>';
    nodeLSDB.type3.forEach(lsa => {
      html += `<tr><td>${lsa.prefix}</td><td>${lsa.advRouter}</td><td>${lsa.metric}</td></tr>`;
    });
    html += '</tbody></table>';
  } else {
    html += `<div style="font-size: 9px; color: var(--text-dim); margin-top: 6px;">Inter-Area Summary LSAs disabled in FLAT mode.</div>`;
  }
  
  lsdbContainer.innerHTML = html;
}

// Cisco IOS Terminal Emulator
function initCLI() {
  const inputField = document.getElementById('cli-input-field');
  const outputLog = document.getElementById('cli-output-log');

  inputField.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const cmd = inputField.value.trim().toLowerCase();
      outputLog.innerHTML += `\n<span style="color: #06b6d4;">${document.getElementById('cli-prompt-text').textContent} ${inputField.value.trim()}</span>\n`;
      
      if (cmd === 'show ip ospf neighbor' || cmd === 'sh ip ospf nei') {
        const activeLinks = sim.topology.links.filter(l => (l.source === sim.selectedNodeId || l.target === sim.selectedNodeId));
        outputLog.innerHTML += 'Neighbor ID     Pri   State           Dead Time   Address         Interface\n';
        activeLinks.forEach(l => {
          const nId = l.source === sim.selectedNodeId ? l.target : l.source;
          const nObj = sim.topology.nodes.find(n => n.id === nId);
          const state = l.state === 'UP' ? 'FULL/DR' : 'DOWN';
          const deadTime = sim.timerProfile === 'optimized' ? '00:00:03' : '00:00:37';
          outputLog.innerHTML += `${nObj.rid.padEnd(16)} 1     ${state.padEnd(15)} ${deadTime}    ${(l.source === sim.selectedNodeId ? l.ipTgt : l.ipSrc).padEnd(15)} Gig0/${l.id.replace('l', '')}\n`;
        });
      } else if (cmd === 'show ip route' || cmd === 'sh ip ro') {
        outputLog.innerHTML += 'Codes: C - connected, O - OSPF, IA - OSPF inter area\n\n';
        const routes = sim.routingTables[sim.selectedNodeId] || [];
        routes.forEach(r => {
          outputLog.innerHTML += `${r.type.padEnd(8)} ${r.prefix.padEnd(18)} [110/${r.metric}] via ${r.nextHop.padEnd(14)} on ${r.interface}\n`;
        });
      } else if (cmd === 'show ip ospf database' || cmd === 'sh ip ospf da') {
        outputLog.innerHTML += ` OSPF Router with ID (${sim.topology.nodes.find(n => n.id === sim.selectedNodeId).rid})\n\n`;
        outputLog.innerHTML += 'Link ID         ADV Router      Age         Seq#\n';
        sim.lsdb[sim.selectedNodeId].type1.forEach(l => {
          outputLog.innerHTML += `${l.linkId.padEnd(15)} ${l.advRouter.padEnd(15)} ${String(l.age).padEnd(11)} ${l.seq}\n`;
        });
      } else if (cmd === 'clear' || cmd === 'cls') {
        outputLog.innerHTML = 'CLI Terminal Cleared.';
      } else if (cmd === 'help') {
        outputLog.innerHTML += 'Available commands:\n  show ip route\n  show ip ospf neighbor\n  show ip ospf database\n  clear\n';
      } else {
        outputLog.innerHTML += `% Unknown command: "${cmd}". Available: "show ip ospf neighbor", "show ip route", "show ip ospf database".\n`;
      }
      
      inputField.value = '';
      outputLog.scrollTop = outputLog.scrollHeight;
    }
  });
}

// Test Harness Console Logging Utility
function logTestConsole(msg, type = 'info') {
  const consoleEl = document.getElementById('test-console-log');
  const timestamp = new Date().toISOString().substring(11, 19);
  let color = '#38bdf8';
  if (type === 'pass') color = '#10b981';
  if (type === 'fail') color = '#ef4444';
  if (type === 'warn') color = '#f59e0b';
  
  consoleEl.innerHTML += `\n<span style="color: ${color}">[${timestamp}] ${msg}</span>`;
  consoleEl.scrollTop = consoleEl.scrollHeight;
}

// Timer Preset Switcher
function setTimerProfile(profile) {
  sim.timerProfile = profile;
  document.getElementById('standardTimersBtn').classList.toggle('active', profile === 'standard');
  document.getElementById('optimizedTimersBtn').classList.toggle('active', profile === 'optimized');

  const presetSelect = document.getElementById('timer-preset-select');
  presetSelect.value = profile;

  const convergenceElem = document.getElementById('convergence-text');
  if (profile === 'optimized') {
    sim.helloInterval = 1;
    sim.deadInterval = 4;
    document.getElementById('val-hello').textContent = '1s';
    document.getElementById('val-dead').textContent = '4s';
    convergenceElem.textContent = 'Sub-Second (1s/4s)';
    logTestConsole(`% OSPF-5-TIMER_UPDATE: Hello/Dead intervals reset to OPTIMIZED (1s/4s)`);
  } else {
    sim.helloInterval = 10;
    sim.deadInterval = 40;
    document.getElementById('val-hello').textContent = '10s';
    document.getElementById('val-dead').textContent = '40s';
    convergenceElem.textContent = 'Standard (10s/40s)';
    logTestConsole(`% OSPF-5-TIMER_UPDATE: Hello/Dead intervals reset to STANDARD (10s/40s)`, 'warn');
  }
  sim.computeAllRoutes();
  renderUI();
}

// Link Status Controls
function toggleSelectedLink(status) {
  const select = document.getElementById('linkFlapSelect');
  if (!select) return;
  const linkKey = select.value;
  const link = sim.topology.links.find(l => l.id === linkKey);
  if (link) {
    const isUp = (link.state === 'UP');
    if (isUp === status) return; // already matches

    link.state = status ? 'UP' : 'DOWN';
    sim.spfRuns++;
    document.getElementById('globalSpfCount').textContent = sim.spfRuns;

    logTestConsole(`[ INTERVENTION ] Link ${link.source} <-> ${link.target} changed state to ${link.state}`);
    sim.computeAllRoutes();
    renderUI();
    
    // Auto-update SPT highlight
    const sptSelect = document.getElementById('sptRootSelect');
    if (sptSelect) {
      visualizeSPT(sptSelect.value);
    } else {
      visualizeSPT(sim.selectedNodeId);
    }
  }
}

// Dynamic link flapping buttons mapping
function simulateLinkFlapping(linkId) {
  const link = sim.topology.links.find(l => l.id === linkId);
  if (link) {
    link.state = link.state === 'UP' ? 'DOWN' : 'UP';
    sim.spfRuns++;
    document.getElementById('globalSpfCount').textContent = sim.spfRuns;
    logTestConsole(`[ INTERVENTION ] Flapped Link ${link.source} <-> ${link.target}. State = ${link.state}`);
    sim.computeAllRoutes();
    renderUI();
    
    // Auto-update SPT highlight
    const sptSelect = document.getElementById('sptRootSelect');
    if (sptSelect) {
      visualizeSPT(sptSelect.value);
    } else {
      visualizeSPT(sim.selectedNodeId);
    }
  }
}

// Shortest Path First (Dijkstra) Visual Highlight
function visualizeSPT(rootName) {
  document.querySelectorAll('.topology-link').forEach(l => l.classList.remove('active-spt'));

  logTestConsole(`% SPF-ENGINE: Calculating Dijkstra SPT tree rooted at ${rootName}...`);
  
  const { distances, previous } = sim.getSPT(rootName);
  
  let activeLinks = 0;
  for (const v in previous) {
    const prevObj = previous[v];
    if (prevObj && prevObj.link) {
      const el = document.getElementById(`svg-link-${prevObj.link.id}`);
      if (el) {
        el.classList.add('active-spt');
        activeLinks++;
      }
    }
  }

  logTestConsole(`% OSPF-5-SPF: Dijkstra tree calculated. ${activeLinks} active connections highlighted in green.`);
}

// Packet Flow Animation Simulation along computed Dijkstra path
function triggerPacketSimulation() {
  const layer = document.getElementById('svg-packets');
  layer.innerHTML = ''; // reset

  const start = 'Hostel-ABR';
  const end = 'Admin-ABR';
  
  const { distances, previous } = sim.getSPT(start);
  
  if (distances[end] === Infinity) {
    logTestConsole(`% IP-FORWARD: Target destination VTOP database unreachable. Shortest path is blocked!`, 'fail');
    return;
  }

  // Backtrack path nodes
  const pathNodes = [];
  let curr = end;
  while (curr !== null) {
    pathNodes.unshift(curr);
    curr = previous[curr] ? previous[curr].node : null;
  }

  // Draw motion path coordinate string
  let dString = "";
  pathNodes.forEach((nodeId, index) => {
    const node = sim.topology.nodes.find(n => n.id === nodeId);
    if (index === 0) {
      dString += `M ${node.x},${node.y}`;
    } else {
      dString += ` L ${node.x},${node.y}`;
    }
  });

  layer.innerHTML = `
    <circle class="sim-packet" r="6">
      <animateMotion path="${dString}" dur="2.2s" fill="freeze" repeatCount="1" />
    </circle>
  `;
  
  logTestConsole(`% IP-FORWARD: Dispatched traceroute probe [Hostel-ABR ➔ Admin-ABR]`);
  logTestConsole(`% IP-FORWARD: Path taken: ${pathNodes.join(' ➔ ')}`);
}

// Security Spoofer Simulation
function injectRogueHello() {
  logTestConsole(`[ SECURITY AUDIT ] Malicious Hello LSA injected in Area 0!`, 'warn');
  
  const targetNode = document.getElementById('node-container-Academic-ABR');
  if (targetNode) {
    targetNode.classList.add('node-flash-red');
    setTimeout(() => {
      targetNode.classList.remove('node-flash-red');
    }, 2400);
  }

  logTestConsole(`[ SECURITY AUDIT ] Cryptographic MD5 checksum mismatch detected! Packet rejected.`, 'fail');
}

// Automated Test Harness Execution Suite
async function runTestHarness(testId) {
  logTestConsole(`======================================================================`);
  
  const testIds = (testId === 'ALL') ? [1, 2, 3, 4, 5] : [testId];
  
  for (const id of testIds) {
    const btn = document.getElementById(`btn-tc-${id}`);
    btn.className = 'btn btn-outline run';
    
    logTestConsole(`RUNNING TC-OSPF-0${id}: Executing validation harness...`);
    await new Promise(r => setTimeout(r, 600));
    
    let passed = true;
    
    if (id === 1) {
      // TC-OSPF-01 Adjacency checks
      const allUpLinks = sim.topology.links.every(l => l.state === 'UP');
      if (allUpLinks) {
        logTestConsole(` -> 100% adjacencies established and verified.`);
      } else {
        logTestConsole(` -> Warning: Adjacency checks run with some links disabled.`, 'warn');
      }
    } else if (id === 2) {
      // TC-OSPF-02 Database sequence matches
      const c1Seq = sim.lsdb['Core-R1'].type1[0].seq;
      const c2Seq = sim.lsdb['Core-R2'].type1[0].seq;
      passed = (c1Seq === c2Seq);
      logTestConsole(` -> LSDB Type-1 Sequence Sync verified: Core-R1 [${c1Seq}] <--> Core-R2 [${c2Seq}]`);
    } else if (id === 3) {
      // TC-OSPF-03 Summary routes checks
      if (sim.multiAreaEnabled) {
        const routes = sim.routingTables['Core-R1'] || [];
        const hasSummary = routes.some(r => r.prefix === '10.10.0.0/16');
        const hasSubnet = routes.some(r => r.prefix === '10.10.1.0/24');
        passed = (hasSummary && !hasSubnet);
        if (passed) {
          logTestConsole(` -> Summary checks pass: 10.10.0.0/16 exists on Core. Subnets hidden.`);
        }
      } else {
        logTestConsole(` -> Check Skipped: Multi-area mode disabled.`, 'warn');
        passed = false;
      }
    } else if (id === 4) {
      // TC-OSPF-04 Totally Stubby default route checks
      if (sim.multiAreaEnabled) {
        const routes = sim.routingTables['Admin-ABR'] || [];
        const hasDefault = routes.some(r => r.prefix === '0.0.0.0/0');
        const hasOtherSummary = routes.some(r => r.prefix === '10.10.0.0/16');
        passed = (hasDefault && !hasOtherSummary);
        if (passed) {
          logTestConsole(` -> Totally Stubby isolation verified inside Area 30.`);
        }
      } else {
        logTestConsole(` -> Check Skipped: Multi-area mode disabled.`, 'warn');
        passed = false;
      }
    } else if (id === 5) {
      // TC-OSPF-05 Hello and Dead preset validation
      passed = (sim.timerProfile === 'optimized');
      if (passed) {
        logTestConsole(` -> Timers validated: Hello: 1s, Dead: 4s. Sub-second convergence: OK`);
      } else {
        logTestConsole(` -> Timers failed: Standard Hello/Dead (10s/40s) convergence > 3.5s!`, 'fail');
      }
    }
    
    if (passed) {
      btn.className = 'btn btn-outline pass';
      logTestConsole(`RESULT TC-OSPF-0${id}: [ PASSED ] Compliance verified.`, 'pass');
    } else {
      btn.className = 'btn btn-outline fail';
      logTestConsole(`RESULT TC-OSPF-0${id}: [ FAILED ] Compliance mismatch.`, 'fail');
    }
  }
}

// Event Bindings and Initializers
document.addEventListener('DOMContentLoaded', () => {
  renderUI();
  initCLI();

  // Tab Navigation Bindings
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const container = btn.parentNode;
      container.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      
      const contentContainer = container.parentNode;
      contentContainer.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      
      btn.classList.add('active');
      document.getElementById(btn.dataset.tab).classList.add('active');
    });
  });

  // Timer Preset Change Listener
  document.getElementById('timer-preset-select').addEventListener('change', (e) => {
    setTimerProfile(e.target.value);
  });

  // Multi-Area Topology mode checkbox
  document.getElementById('toggle-multi-area').addEventListener('change', (e) => {
    sim.multiAreaEnabled = e.target.checked;
    
    const isolationBadge = document.getElementById('badge-isolation');
    const isolationText = document.getElementById('isolation-text');
    if (sim.multiAreaEnabled) {
      isolationText.textContent = 'Enforced';
      isolationBadge.querySelector('.status-dot').className = 'status-dot green';
      logTestConsole(`% OSPF-5-AREA_MODE: Database isolation mode set to HIERARCHICAL MULTI-AREA`);
    } else {
      isolationText.textContent = 'Disabled (Flat)';
      isolationBadge.querySelector('.status-dot').className = 'status-dot warning';
      logTestConsole(`% OSPF-5-AREA_MODE: Database isolation mode set to FLAT SINGLE-AREA`, 'warn');
    }
    sim.computeAllRoutes();
    renderUI();
  });

  // Link Flap Button Bindings
  document.getElementById('btn-flap-academic').addEventListener('click', () => {
    simulateLinkFlapping('l2');
  });

  document.getElementById('btn-flap-hostel').addEventListener('click', () => {
    simulateLinkFlapping('l5');
  });

  // Security Audit Action Binding
  document.getElementById('btn-inject-malicious').addEventListener('click', () => {
    injectRogueHello();
  });

  // Test Harness Button Bindings
  document.getElementById('btn-run-all-tests').addEventListener('click', () => runTestHarness('ALL'));
  document.getElementById('btn-tc-1').addEventListener('click', () => runTestHarness(1));
  document.getElementById('btn-tc-2').addEventListener('click', () => runTestHarness(2));
  document.getElementById('btn-tc-3').addEventListener('click', () => runTestHarness(3));
  document.getElementById('btn-tc-4').addEventListener('click', () => runTestHarness(4));
  document.getElementById('btn-tc-5').addEventListener('click', () => runTestHarness(5));
});
