const svg = document.getElementById("metroCanvas");
const viewport = document.getElementById("viewport");
const streamsLayer = document.getElementById("streamsLayer");
const nodesLayer = document.getElementById("nodesLayer");
const labelsLayer = document.getElementById("labelsLayer");
const connectionsLayer = document.getElementById("connectionsLayer");
const previewLayer = document.getElementById("previewLayer");
const handlesLayer = document.getElementById("handlesLayer");
const guidesLayer = document.getElementById("guidesLayer");
const calendarGuidesLayer = document.getElementById("calendarGuidesLayer");
const calendarBar = document.getElementById("calendarBar");
const canvasContextMenu = document.getElementById("canvasContextMenu");

const streamLegend = document.getElementById("streamLegend");
const selectionEditor = document.getElementById("selectionEditor");
const toolHint = document.getElementById("toolHint");

const newStreamNameInput = document.getElementById("newStreamName");
const newStreamColorInput = document.getElementById("newStreamColor");
const colorPalette = document.getElementById("colorPalette");
const colorHex = document.getElementById("colorHex");
const streamPresetSelect = document.getElementById("streamPreset");
const streamSetupHint = document.getElementById("streamSetupHint");

let selectedStreamColor = "#86BC25";

function initColorPalette() {
  newStreamColorInput.addEventListener("input", () => {
    selectedStreamColor = newStreamColorInput.value;
    colorHex.textContent = selectedStreamColor.toUpperCase();
  });
}

const toolButtons = document.getElementById("toolButtons");
const undoBtn = document.getElementById("undoBtn");
const redoBtn = document.getElementById("redoBtn");
const fitToViewBtn = document.getElementById("fitToView");
const goToTodayBtn = document.getElementById("goToToday");
const toggleGridBtn = document.getElementById("toggleGrid");
const themeToggle = document.getElementById("themeToggle");
const deleteSelectedBtn = document.getElementById("deleteSelected");
const exportPngBtn = document.getElementById("exportPng");
const exportSvgBtn = document.getElementById("exportSvg");
const exportLegendPngBtn = document.getElementById("exportLegendPng");
const exportLegendSvgBtn = document.getElementById("exportLegendSvg");
const exportViewerBtn = document.getElementById("exportViewer");
const saveMapBtn = document.getElementById("saveMap");
const loadMapBtn = document.getElementById("loadMap");
const loadMapFileInput = document.getElementById("loadMapFile");

const GRID_SIZE = 20;
let snapEnabled = true;
let labelFontSize = 14;

const STATUS_COLORS = {
  planned: "#ffffff",
  completed: "#92D050",
  ongoing: "#FFCD00",
  attention: "#ED8B00",
};

const DEFAULT_LABEL_MAX_WIDTH = 120;
const DEFAULT_MONTH_WIDTH = 1500;
const MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

const TOOL_HINTS = {
  select: "Select streams or stops. Edit details in the side panel.",
  "add-stream": "Click and drag to create a stream snapped to the grid.",
  "add-activity": "Click and drag to place an activity stop on the grid.",
  "add-deliverable": "Click and drag to place a deliverable stop on the grid.",
  connect: "Drag from one stop to another to create a connection.",
  "add-text-label": "Click to place a floating text label on the canvas.",
};

const state = {
  tool: "select",
  selected: [],
  selectionGrouped: false,
  streams: [],
  streamOrder: [],
  nodes: [],
  connections: [],
  guides: [],
  textLabels: [],
  calendar: {
    monthWidths: {},
    refYear: new Date().getFullYear(),
  },
  transform: { x: 0, y: 0, scale: 1 },
  pan: { active: false, startX: 0, startY: 0, baseX: 0, baseY: 0 },
  interaction: {
    mode: null,
    connectFromNodeId: null,
    connectHoverNodeId: null,
    dragMoved: false,
    historyCaptured: false,
    dragNodeId: null,
    nodeCreateStreamId: null,
    startPoint: null,
    previewPoint: null,
    dragStreamId: null,
    dragPointIndex: null,
    dragStreamOrigPoints: null,
    dragStreamOrigNodes: null,
    snapIndicator: null,
    dragGuideId: null,
    boxSelectStart: null,
    boxSelectEnd: null,
    multiDragOrigItems: null,
    dragTextLabelId: null,
    calendarDragKey: null,
    calendarDragStartX: null,
    calendarDragStartWidth: null,
  },
  history: {
    undo: [],
    redo: [],
    max: 120,
    isApplying: false,
  },
};

function uid(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function cloneData(value) {
  return JSON.parse(JSON.stringify(value));
}

function buildSnapshot() {
  return {
    streams: cloneData(state.streams),
    nodes: cloneData(state.nodes),
    connections: cloneData(state.connections),
    guides: cloneData(state.guides),
    textLabels: cloneData(state.textLabels),
    calendar: cloneData(state.calendar),
    selected: cloneData(state.selected),
    selectionGrouped: state.selectionGrouped,
  };
}

function updateUndoRedoButtons() {
  undoBtn.disabled = state.history.undo.length === 0;
  redoBtn.disabled = state.history.redo.length === 0;
}

function captureHistorySnapshot() {
  if (state.history.isApplying) {
    return;
  }

  state.history.undo.push(buildSnapshot());
  if (state.history.undo.length > state.history.max) {
    state.history.undo.shift();
  }
  state.history.redo = [];
  updateUndoRedoButtons();
}

function restoreFromSnapshot(snapshot) {
  state.streams = cloneData(snapshot.streams || []);
  state.nodes = cloneData(snapshot.nodes || []);
  state.connections = cloneData(snapshot.connections || []);
  state.guides = cloneData(snapshot.guides || []);
  state.textLabels = cloneData(snapshot.textLabels || []);
  if (snapshot.calendar) {
    state.calendar = cloneData(snapshot.calendar);
  }
  state.selected = cloneData(snapshot.selected || []);
  state.selectionGrouped = !!snapshot.selectionGrouped;


  renderSelectionEditor();
  render();
}

function undo() {
  if (state.history.undo.length === 0) {
    return;
  }

  const current = buildSnapshot();
  const previous = state.history.undo.pop();
  state.history.redo.push(current);
  state.history.isApplying = true;
  restoreFromSnapshot(previous);
  state.history.isApplying = false;
  updateUndoRedoButtons();
}

function redo() {
  if (state.history.redo.length === 0) {
    return;
  }

  const current = buildSnapshot();
  const next = state.history.redo.pop();
  state.history.undo.push(current);
  state.history.isApplying = true;
  restoreFromSnapshot(next);
  state.history.isApplying = false;
  updateUndoRedoButtons();
}

function createStream({ name, color, points, parentStreamId = null }) {
  captureHistorySnapshot();
  const stream = {
    id: uid("stream"),
    name: name || "New Stream",
    color: color || randomColor(),
    points: points || [],
    parentStreamId,
  };
  state.streams.push(stream);
  render();
  return stream;
}

function createNode({ streamId, type, x, y, label, date, status = "planned" }) {
  captureHistorySnapshot();
  const node = {
    id: uid("node"),
    streamId,
    type,
    x,
    y,
    label: label || (type === "deliverable" ? "Deliverable" : "Activity"),
    date: date || "",
    status,
    labelTextColor: type === "deliverable" ? "#ffffff" : "#000000",
    labelBgColor: type === "deliverable" ? "#000000" : null,
    labelMaxWidth: null,
    labelPosition: type === "deliverable" ? undefined : "down",
    labelDescription: type === "deliverable" ? "" : undefined,
    labelDescPosition: type === "deliverable" ? "right" : undefined,
    subLabels: type === "deliverable" ? [] : undefined,
  };
  state.nodes.push(node);
  render();
  return node;
}

function createConnection({ fromNodeId, toNodeId }) {
  captureHistorySnapshot();
  const connection = {
    id: uid("conn"),
    fromNodeId,
    toNodeId,
  };
  state.connections.push(connection);
  render();
  return connection;
}

function createTextLabel({ x, y, text, fontSize, maxWidth, align, textColor, bgColor }) {
  captureHistorySnapshot();
  const tl = {
    id: uid("tlabel"),
    x,
    y,
    text: text || "Text",
    fontSize: fontSize || 16,
    maxWidth: maxWidth || 160,
    align: align || "center",
    textColor: textColor || "#ffffff",
    bgColor: bgColor || null,
  };
  state.textLabels.push(tl);
  render();
  return tl;
}

function randomColor() {
  const palette = ["#54b7ff", "#ff8d66", "#8ccf5f", "#c88fff", "#f7d15a", "#66d6c1"];
  return palette[Math.floor(Math.random() * palette.length)];
}

function syncStreamPresetDropdown() {
  const current = streamPresetSelect.value;
  streamPresetSelect.innerHTML = "";

  const newOpt = document.createElement("option");
  newOpt.value = "__new__";
  newOpt.textContent = "\u2014 New Stream \u2014";
  streamPresetSelect.appendChild(newOpt);

  const seen = new Map();
  state.streams.forEach((stream) => {
    const key = normalizeColor(stream.color);
    if (!seen.has(key)) {
      seen.set(key, stream);
      const opt = document.createElement("option");
      opt.value = stream.id;
      opt.textContent = stream.name;
      streamPresetSelect.appendChild(opt);
    }
  });

  if ([...streamPresetSelect.options].some((o) => o.value === current)) {
    streamPresetSelect.value = current;
  } else {
    streamPresetSelect.value = "__new__";
  }
}

function applyStreamPreset() {
  const val = streamPresetSelect.value;
  if (val === "__new__") {
    newStreamNameInput.value = "";
    newStreamNameInput.disabled = false;
    colorPalette.style.pointerEvents = "";
    colorPalette.style.opacity = "";
    streamSetupHint.textContent = "";
    return;
  }
  const stream = getStreamById(val);
  if (stream) {
    newStreamNameInput.value = stream.name;
    newStreamNameInput.disabled = true;
    colorPalette.style.pointerEvents = "none";
    colorPalette.style.opacity = "0.5";
    selectSwatchByColor(stream.color);
    streamSetupHint.textContent = "Adding a segment to the existing \"" + stream.name + "\" stream.";
  }
}

function selectSwatchByColor(color) {
  selectedStreamColor = color;
  newStreamColorInput.value = color;
  colorHex.textContent = color.toUpperCase();
}

function isStreamNameTaken(name) {
  const lower = name.trim().toLowerCase();
  return state.streams.some((s) => s.name.trim().toLowerCase() === lower);
}

function isStreamColorTaken(color) {
  const c = normalizeColor(color);
  return state.streams.some((s) => normalizeColor(s.color) === c);
}



function setTool(tool) {
  state.tool = tool;
  state.interaction.mode = null;
  state.interaction.dragNodeId = null;
  state.interaction.nodeCreateStreamId = null;

  document.querySelectorAll("#toolButtons button[data-tool]").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tool === tool);
  });

  toolHint.textContent = TOOL_HINTS[tool] || "";
}

function clearSelection() {
  state.selected = [];
  state.selectionGrouped = false;
  renderSelectionEditor();
}

function selectItem(type, id) {
  state.selected = [{ type, id }];
  state.selectionGrouped = (type === "stream");
  renderSelectionEditor();
  render();
}

function toggleSelectItem(type, id) {
  const idx = state.selected.findIndex((s) => s.type === type && s.id === id);
  if (idx >= 0) {
    state.selected.splice(idx, 1);
  } else {
    state.selected.push({ type, id });
  }
  state.selectionGrouped = false;
  renderSelectionEditor();
  render();
}

function isItemSelected(type, id) {
  return state.selected.some((s) => s.type === type && s.id === id);
}

function isMultiSelection() {
  return state.selected.length > 1;
}

function buildMultiDragOrigItems() {
  const origStreams = [];
  const origNodes = [];
  const origTextLabels = [];
  const includedNodeIds = new Set();
  const selectedStreamIds = new Set();
  state.selected.forEach((sel) => {
    if (sel.type === "stream") {
      const s = getStreamById(sel.id);
      if (s) {
        origStreams.push({ streamId: s.id, points: cloneData(s.points) });
        selectedStreamIds.add(s.id);
      }
    } else if (sel.type === "node") {
      const n = state.nodes.find((nd) => nd.id === sel.id);
      if (n) {
        origNodes.push({ id: n.id, x: n.x, y: n.y });
        includedNodeIds.add(n.id);
      }
    } else if (sel.type === "text-label") {
      const tl = state.textLabels.find((t) => t.id === sel.id);
      if (tl) {
        origTextLabels.push({ id: tl.id, x: tl.x, y: tl.y });
      }
    }
  });
  // Also include nodes that belong to selected streams
  selectedStreamIds.forEach((sid) => {
    state.nodes.filter((n) => n.streamId === sid && !includedNodeIds.has(n.id)).forEach((n) => {
      origNodes.push({ id: n.id, x: n.x, y: n.y });
      includedNodeIds.add(n.id);
    });
  });
  return { streams: origStreams, nodes: origNodes, textLabels: origTextLabels };
}

function getSelectedEntity() {
  if (state.selected.length !== 1) {
    return null;
  }
  const sel = state.selected[0];

  if (sel.type === "stream") {
    const stream = state.streams.find((s) => s.id === sel.id);
    return stream ? { type: "stream", entity: stream } : null;
  }

  if (sel.type === "node") {
    const node = state.nodes.find((n) => n.id === sel.id);
    return node ? { type: "node", entity: node } : null;
  }

  if (sel.type === "connection") {
    const conn = state.connections.find((c) => c.id === sel.id);
    return conn ? { type: "connection", entity: conn } : null;
  }

  if (sel.type === "text-label") {
    const tl = state.textLabels.find((t) => t.id === sel.id);
    return tl ? { type: "text-label", entity: tl } : null;
  }

  return null;
}

function svgPointFromEvent(event) {
  const pt = svg.createSVGPoint();
  pt.x = event.clientX;
  pt.y = event.clientY;
  const svgPt = pt.matrixTransform(svg.getScreenCTM().inverse());
  return {
    x: (svgPt.x - state.transform.x) / state.transform.scale,
    y: (svgPt.y - state.transform.y) / state.transform.scale,
  };
}

function snap(value) {
  return snapEnabled ? Math.round(value / GRID_SIZE) * GRID_SIZE : value;
}

function snapPoint(point) {
  let result = snapEnabled
    ? { x: snap(point.x), y: snap(point.y) }
    : { x: point.x, y: point.y };

  // Snap to nearby guides (guide snap always active, overrides grid)
  const GUIDE_SNAP_DIST = 12;
  state.guides.forEach((g) => {
    if (g.axis === "v" && Math.abs(point.x - g.position) < GUIDE_SNAP_DIST) {
      result.x = g.position;
    }
    if (g.axis === "h" && Math.abs(point.y - g.position) < GUIDE_SNAP_DIST) {
      result.y = g.position;
    }
  });
  return result;
}

function nearestPointOnSegment(point, a, b) {
  const abx = b.x - a.x;
  const aby = b.y - a.y;
  const abLenSq = abx * abx + aby * aby;
  if (abLenSq === 0) {
    return { x: a.x, y: a.y, distanceSq: (point.x - a.x) ** 2 + (point.y - a.y) ** 2 };
  }

  const apx = point.x - a.x;
  const apy = point.y - a.y;
  const t = clamp((apx * abx + apy * aby) / abLenSq, 0, 1);
  const x = a.x + abx * t;
  const y = a.y + aby * t;
  return { x, y, distanceSq: (point.x - x) ** 2 + (point.y - y) ** 2 };
}

function nearestPointOnStream(point, stream) {
  if (!stream || stream.points.length === 0) {
    return null;
  }
  if (stream.points.length === 1) {
    return { ...stream.points[0] };
  }

  let best = null;
  for (let i = 0; i < stream.points.length - 1; i += 1) {
    const candidate = nearestPointOnSegment(point, stream.points[i], stream.points[i + 1]);
    if (!best || candidate.distanceSq < best.distanceSq) {
      best = candidate;
    }
  }

  return best ? { x: best.x, y: best.y } : { ...stream.points[0] };
}

function normalizeColor(color) {
  return String(color || "").trim().toLowerCase();
}

function pointsClose(a, b, tolerance = 1.5) {
  return (a.x - b.x) ** 2 + (a.y - b.y) ** 2 <= tolerance ** 2;
}

function streamSegments(stream) {
  const segments = [];
  for (let i = 0; i < stream.points.length - 1; i += 1) {
    segments.push([stream.points[i], stream.points[i + 1]]);
  }
  return segments;
}

function streamsGeometricallyConnected(streamA, streamB) {
  if (!streamA || !streamB || streamA.points.length < 2 || streamB.points.length < 2) {
    return false;
  }

  const endpointsA = [streamA.points[0], streamA.points[streamA.points.length - 1]];
  const endpointsB = [streamB.points[0], streamB.points[streamB.points.length - 1]];

  // Case 1: endpoint-to-endpoint
  for (const ea of endpointsA) {
    for (const eb of endpointsB) {
      if (pointsClose(ea, eb, 3)) return true;
    }
  }

  // Case 2: endpoint of A lies on a segment of B
  const segB = streamSegments(streamB);
  for (const ea of endpointsA) {
    for (const [b1, b2] of segB) {
      const proj = nearestPointOnSegment(ea, b1, b2);
      if (proj.distanceSq <= 9) return true;
    }
  }

  // Case 3: endpoint of B lies on a segment of A
  const segA = streamSegments(streamA);
  for (const eb of endpointsB) {
    for (const [a1, a2] of segA) {
      const proj = nearestPointOnSegment(eb, a1, a2);
      if (proj.distanceSq <= 9) return true;
    }
  }

  return false;
}

function getConnectedSameColorStreamIds(startStreamId) {
  const startStream = getStreamById(startStreamId);
  if (!startStream) {
    return new Set();
  }

  const targetColor = normalizeColor(startStream.color);
  const adjacency = new Map();

  state.streams.forEach((stream) => {
    if (normalizeColor(stream.color) === targetColor) {
      adjacency.set(stream.id, new Set());
    }
  });

  const sameColorStreams = state.streams.filter((stream) => normalizeColor(stream.color) === targetColor);
  for (let i = 0; i < sameColorStreams.length; i += 1) {
    for (let j = i + 1; j < sameColorStreams.length; j += 1) {
      const a = sameColorStreams[i];
      const b = sameColorStreams[j];
      if (streamsGeometricallyConnected(a, b)) {
        adjacency.get(a.id)?.add(b.id);
        adjacency.get(b.id)?.add(a.id);
      }
    }
  }

  const visited = new Set();
  const queue = [startStreamId];

  while (queue.length > 0) {
    const current = queue.shift();
    if (!current || visited.has(current)) {
      continue;
    }
    visited.add(current);

    const neighbors = adjacency.get(current);
    if (!neighbors) {
      continue;
    }

    neighbors.forEach((nextId) => {
      if (!visited.has(nextId)) {
        queue.push(nextId);
      }
    });
  }

  return visited;
}

function nearestPointOnConnectedStreams(point, startStreamId) {
  const candidateIds = getConnectedSameColorStreamIds(startStreamId);
  if (candidateIds.size === 0) {
    return null;
  }

  let best = null;

  candidateIds.forEach((streamId) => {
    const stream = getStreamById(streamId);
    if (!stream) {
      return;
    }
    const nearest = nearestPointOnStream(point, stream);
    if (!nearest) {
      return;
    }

    const distanceSq = (point.x - nearest.x) ** 2 + (point.y - nearest.y) ** 2;
    if (!best || distanceSq < best.distanceSq) {
      best = {
        streamId,
        x: nearest.x,
        y: nearest.y,
        distanceSq,
      };
    }
  });

  return best;
}

function getStreamById(streamId) {
  return state.streams.find((stream) => stream.id === streamId) || null;
}

const HANDLE_SNAP_DISTANCE = 15;

function getSnapTarget(point, excludeStreamId) {
  let best = null;
  let bestDist = (HANDLE_SNAP_DISTANCE + 5) * (HANDLE_SNAP_DISTANCE + 5);

  state.streams.forEach((stream) => {
    if (stream.id === excludeStreamId) {
      return;
    }
    // Check vertices (higher priority)
    stream.points.forEach((p) => {
      const dist = (point.x - p.x) ** 2 + (point.y - p.y) ** 2;
      if (dist < HANDLE_SNAP_DISTANCE * HANDLE_SNAP_DISTANCE && (!best || dist < bestDist)) {
        bestDist = dist;
        best = { x: p.x, y: p.y, isEndpoint: true };
      }
    });

    // Check projection onto segments
    if (stream.points.length >= 2) {
      for (let i = 0; i < stream.points.length - 1; i++) {
        const a = stream.points[i];
        const b = stream.points[i + 1];
        const proj = nearestPointOnSegment(point, a, b);
        if (proj.distanceSq < bestDist && !best?.isEndpoint) {
          bestDist = proj.distanceSq;
          best = { x: proj.x, y: proj.y, isEndpoint: false };
        }
      }
    }
  });

  return best;
}

function getLineSnapTarget(point, targetColor) {
  let best = null;
  let bestDist = (HANDLE_SNAP_DISTANCE + 5) * (HANDLE_SNAP_DISTANCE + 5);
  const normColor = normalizeColor(targetColor);

  state.streams.forEach((stream) => {
    if (normalizeColor(stream.color) !== normColor) return;
    if (stream.points.length < 2) return;

    // First check endpoints (higher priority)
    const endpoints = [stream.points[0], stream.points[stream.points.length - 1]];
    for (const ep of endpoints) {
      const dist = (point.x - ep.x) ** 2 + (point.y - ep.y) ** 2;
      if (dist < HANDLE_SNAP_DISTANCE * HANDLE_SNAP_DISTANCE && (!best || dist < bestDist)) {
        bestDist = dist;
        best = { x: ep.x, y: ep.y, isEndpoint: true, streamId: stream.id };
      }
    }

    // Then check projection onto segments (midpoint)
    for (let i = 0; i < stream.points.length - 1; i++) {
      const a = stream.points[i];
      const b = stream.points[i + 1];
      const proj = nearestPointOnSegment(point, a, b);
      if (proj.distanceSq < bestDist && !best?.isEndpoint) {
        bestDist = proj.distanceSq;
        best = { x: proj.x, y: proj.y, isEndpoint: false, streamId: stream.id, segA: a, segB: b };
      }
    }
  });

  return best;
}

function constrainActivityToStream(point, streamId) {
  const nearest = nearestPointOnConnectedStreams(point, streamId);
  if (!nearest) {
    return point;
  }
  return nearest;
}

function getSelectedStreamGroup() {
  if (!state.selectionGrouped || state.selected.length !== 1 || state.selected[0].type !== "stream") {
    return new Set();
  }
  return getConnectedSameColorStreamIds(state.selected[0].id);
}

function getNodeLabelStyle(node) {
  return {
    textColor: node.labelTextColor || (node.type === "deliverable" ? "#ffffff" : "#000000"),
    bgColor: node.labelBgColor === undefined
      ? node.type === "deliverable"
        ? "#000000"
        : null
      : node.labelBgColor,
  };
}

function applyViewportTransform() {
  const { x, y, scale } = state.transform;
  viewport.setAttribute("transform", `translate(${x}, ${y}) scale(${scale})`);
}

/* ── Calendar helpers ──────────────────────────────────────────── */

function calMonthKey(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function calMonthWidth(key) {
  return state.calendar.monthWidths[key] || DEFAULT_MONTH_WIDTH;
}

function calMonthFromIndex(idx) {
  const year = Math.floor(idx / 12);
  const month = (((idx % 12) + 12) % 12) + 1;
  return { year, month };
}

function calMonthIndex(year, month) {
  return year * 12 + (month - 1);
}

function calMonthSvgX(targetYear, targetMonth) {
  const refIdx = calMonthIndex(state.calendar.refYear, 1);
  const targetIdx = calMonthIndex(targetYear, targetMonth);
  let x = 0;
  if (targetIdx >= refIdx) {
    for (let i = refIdx; i < targetIdx; i++) {
      const m = calMonthFromIndex(i);
      x += calMonthWidth(calMonthKey(m.year, m.month));
    }
  } else {
    for (let i = refIdx - 1; i >= targetIdx; i--) {
      const m = calMonthFromIndex(i);
      x -= calMonthWidth(calMonthKey(m.year, m.month));
    }
  }
  return x;
}

function getSvgScreenMapping() {
  const cw = svg.clientWidth || 1;
  const ch = svg.clientHeight || 1;
  const svgScale = Math.min(cw / 1600, ch / 1000);
  const offsetX = (cw - 1600 * svgScale) / 2;
  return { svgScale, offsetX };
}

function svgXToPixelX(svgX) {
  const { svgScale, offsetX } = getSvgScreenMapping();
  const vbX = svgX * state.transform.scale + state.transform.x;
  return vbX * svgScale + offsetX;
}

function pixelXToSvgX(pixelX) {
  const { svgScale, offsetX } = getSvgScreenMapping();
  const vbX = (pixelX - offsetX) / svgScale;
  return (vbX - state.transform.x) / state.transform.scale;
}

function renderCalendar() {
  calendarBar.innerHTML = "";
  calendarGuidesLayer.innerHTML = "";

  const barWidth = calendarBar.clientWidth || 1;
  const leftSvgX = pixelXToSvgX(0);
  const rightSvgX = pixelXToSvgX(barWidth);

  const refIdx = calMonthIndex(state.calendar.refYear, 1);
  const now = new Date();
  const curKey = calMonthKey(now.getFullYear(), now.getMonth() + 1);

  // find first visible month
  let scanIdx = refIdx;
  let scanX = 0;
  while (scanX > leftSvgX) {
    scanIdx--;
    const m = calMonthFromIndex(scanIdx);
    scanX -= calMonthWidth(calMonthKey(m.year, m.month));
  }
  // go one more left for partial visibility
  scanIdx--;
  const m0 = calMonthFromIndex(scanIdx);
  scanX -= calMonthWidth(calMonthKey(m0.year, m0.month));

  const fragment = document.createDocumentFragment();
  const guideFragment = document.createDocumentFragment();
  const GUIDE_EXTENT = 10000;

  let idx = scanIdx;
  let x = scanX;

  while (x < rightSvgX + DEFAULT_MONTH_WIDTH) {
    const m = calMonthFromIndex(idx);
    const key = calMonthKey(m.year, m.month);
    const w = calMonthWidth(key);
    const pxLeft = svgXToPixelX(x);
    const pxRight = svgXToPixelX(x + w);
    const pxW = pxRight - pxLeft;

    if (pxRight > -10 && pxLeft < barWidth + 10) {
      // Month box
      const box = document.createElement("div");
      box.className = "cal-month" + (m.month === 1 ? " cal-jan" : "") + (key === curKey ? " cal-current" : "");
      box.style.left = pxLeft + "px";
      box.style.width = pxW + "px";
      if (pxW > 28) {
        box.textContent = m.month === 1 ? `${MONTH_NAMES[0]} ${m.year}` : MONTH_NAMES[m.month - 1];
      }
      fragment.appendChild(box);

      // Drag border (right edge)
      const border = document.createElement("div");
      border.className = "cal-border";
      border.style.left = (pxRight - 5) + "px";
      border.dataset.monthKey = key;
      fragment.appendChild(border);

      // SVG guide line at right edge (in SVG coords, inside viewport)
      const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
      line.setAttribute("x1", x + w);
      line.setAttribute("y1", -GUIDE_EXTENT);
      line.setAttribute("x2", x + w);
      line.setAttribute("y2", GUIDE_EXTENT);
      line.classList.add("calendar-guide-line");
      guideFragment.appendChild(line);
    }

    x += w;
    idx++;
  }

  calendarBar.appendChild(fragment);
  calendarGuidesLayer.appendChild(guideFragment);
}

function renderGuides() {
  guidesLayer.innerHTML = "";
  const EXTENT = 1e6;
  state.guides.forEach((guide) => {
    // Invisible wider hit area for grabbing (rendered first, behind)
    const hit = document.createElementNS("http://www.w3.org/2000/svg", "line");
    if (guide.axis === "v") {
      hit.setAttribute("x1", guide.position);
      hit.setAttribute("y1", -EXTENT);
      hit.setAttribute("x2", guide.position);
      hit.setAttribute("y2", EXTENT);
    } else {
      hit.setAttribute("x1", -EXTENT);
      hit.setAttribute("y1", guide.position);
      hit.setAttribute("x2", EXTENT);
      hit.setAttribute("y2", guide.position);
    }
    hit.classList.add("guide-hit");
    hit.dataset.guideId = guide.id;
    guidesLayer.appendChild(hit);

    // Visible dashed line (rendered on top, no pointer events)
    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    if (guide.axis === "v") {
      line.setAttribute("x1", guide.position);
      line.setAttribute("y1", -EXTENT);
      line.setAttribute("x2", guide.position);
      line.setAttribute("y2", EXTENT);
    } else {
      line.setAttribute("x1", -EXTENT);
      line.setAttribute("y1", guide.position);
      line.setAttribute("x2", EXTENT);
      line.setAttribute("y2", guide.position);
    }
    line.classList.add("guide-line");
    line.dataset.guideId = guide.id;
    guidesLayer.appendChild(line);
  });
}

function render() {
  renderGuides();
  renderStreams();
  renderConnections();
  renderNodes();
  renderHandles();
  renderLabels();
  renderPreview();
  renderLegend();
  applyViewportTransform();
  renderCalendar();
}

function renderPreview() {
  previewLayer.innerHTML = "";
  const mode = state.interaction.mode;
  const start = state.interaction.startPoint;
  const end = state.interaction.previewPoint;

  if (mode === "box-select") {
    const bs = state.interaction.boxSelectStart;
    const be = state.interaction.boxSelectEnd;
    if (bs && be) {
      const x = Math.min(bs.x, be.x);
      const y = Math.min(bs.y, be.y);
      const w = Math.abs(be.x - bs.x);
      const h = Math.abs(be.y - bs.y);
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.setAttribute("x", x);
      rect.setAttribute("y", y);
      rect.setAttribute("width", w);
      rect.setAttribute("height", h);
      rect.classList.add("box-select-rect");
      previewLayer.appendChild(rect);
    }
    return;
  }

  if (!mode || !start || !end) {
    return;
  }

  if (mode === "connect-drag") {
    const fromNode = state.nodes.find((n) => n.id === state.interaction.connectFromNodeId);
    if (!fromNode) return;

    const targetNode = state.interaction.connectHoverNodeId
      ? state.nodes.find((n) => n.id === state.interaction.connectHoverNodeId)
      : null;

    const endX = targetNode ? targetNode.x : end.x;
    const endY = targetNode ? targetNode.y : end.y;

    const c1x = fromNode.x + (endX - fromNode.x) * 0.35;
    const c2x = fromNode.x + (endX - fromNode.x) * 0.65;
    const d = `M ${fromNode.x} ${fromNode.y} C ${c1x} ${fromNode.y}, ${c2x} ${endY}, ${endX} ${endY}`;

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.setAttribute("d", d);
    path.setAttribute("fill", "none");
    path.setAttribute("stroke", targetNode ? "rgba(79,194,255,0.7)" : "rgba(255,255,255,0.45)");
    path.setAttribute("stroke-width", "3");
    path.setAttribute("stroke-dasharray", "7 7");
    previewLayer.appendChild(path);

    // Highlight source node
    const srcRing = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    srcRing.setAttribute("cx", fromNode.x);
    srcRing.setAttribute("cy", fromNode.y);
    srcRing.setAttribute("r", 18);
    srcRing.setAttribute("fill", "none");
    srcRing.setAttribute("stroke", "rgba(79,194,255,0.6)");
    srcRing.setAttribute("stroke-width", "2");
    previewLayer.appendChild(srcRing);

    // Highlight target node when hovering
    if (targetNode) {
      const tgtRing = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      tgtRing.setAttribute("cx", targetNode.x);
      tgtRing.setAttribute("cy", targetNode.y);
      tgtRing.setAttribute("r", 18);
      tgtRing.setAttribute("fill", "none");
      tgtRing.setAttribute("stroke", "rgba(79,194,255,0.6)");
      tgtRing.setAttribute("stroke-width", "2");
      previewLayer.appendChild(tgtRing);
    }
    return;
  }

  if (mode === "stream-create") {
    const color = selectedStreamColor || "rgba(255,255,255,0.8)";

    const anchorDot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    anchorDot.setAttribute("cx", start.x);
    anchorDot.setAttribute("cy", start.y);
    anchorDot.setAttribute("r", 10);
    anchorDot.setAttribute("fill", "none");
    anchorDot.setAttribute("stroke", color);
    anchorDot.setAttribute("stroke-width", "2");
    anchorDot.setAttribute("stroke-dasharray", "4 3");
    previewLayer.appendChild(anchorDot);

    const line = document.createElementNS("http://www.w3.org/2000/svg", "line");
    line.classList.add("preview-line");
    line.setAttribute("stroke", color);
    line.setAttribute("x1", start.x);
    line.setAttribute("y1", start.y);
    line.setAttribute("x2", end.x);
    line.setAttribute("y2", end.y);
    previewLayer.appendChild(line);

    const endDot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    endDot.setAttribute("cx", end.x);
    endDot.setAttribute("cy", end.y);
    endDot.setAttribute("r", 10);
    endDot.setAttribute("fill", "none");
    endDot.setAttribute("stroke", color);
    endDot.setAttribute("stroke-width", "2");
    endDot.setAttribute("stroke-dasharray", "4 3");
    previewLayer.appendChild(endDot);

    // Show snap indicator when projected onto same-color stream
    const snap = state.interaction.snapIndicator;
    if (snap && !snap.isEndpoint && snap.segA && snap.segB) {
      // Draw a small crosshair/projection marker on the target segment
      const projDot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      projDot.setAttribute("cx", snap.x);
      projDot.setAttribute("cy", snap.y);
      projDot.setAttribute("r", 6);
      projDot.setAttribute("fill", "rgba(79, 194, 255, 0.6)");
      projDot.setAttribute("stroke", "#fff");
      projDot.setAttribute("stroke-width", "2");
      previewLayer.appendChild(projDot);
    } else if (snap && snap.isEndpoint) {
      // Highlight the target endpoint
      const epRing = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      epRing.setAttribute("cx", snap.x);
      epRing.setAttribute("cy", snap.y);
      epRing.setAttribute("r", 14);
      epRing.setAttribute("fill", "none");
      epRing.setAttribute("stroke", "rgba(79, 194, 255, 0.8)");
      epRing.setAttribute("stroke-width", "2");
      previewLayer.appendChild(epRing);
    }
    return;
  }

  if (mode === "node-create") {
    const targetStreamId = state.interaction.nodeCreateStreamId || state.streams[0]?.id;
    const previewPoint = targetStreamId ? constrainActivityToStream(end, targetStreamId) : end;

    if (state.tool === "add-deliverable") {
      const diamond = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      diamond.setAttribute("x", previewPoint.x - 11);
      diamond.setAttribute("y", previewPoint.y - 11);
      diamond.setAttribute("width", 22);
      diamond.setAttribute("height", 22);
      diamond.setAttribute("transform", `rotate(45 ${previewPoint.x} ${previewPoint.y})`);
      diamond.classList.add("preview-node");
      previewLayer.appendChild(diamond);
      return;
    }

    const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
    circle.setAttribute("cx", previewPoint.x);
    circle.setAttribute("cy", previewPoint.y);
    circle.setAttribute("r", 12);
    circle.classList.add("preview-node");
    previewLayer.appendChild(circle);
  }
}

function renderStreams() {
  streamsLayer.innerHTML = "";
  const groupIds = getSelectedStreamGroup();

  state.streams.forEach((stream) => {
    if (stream.points.length < 2) {
      return;
    }

    const points = stream.points.map((p) => `${p.x},${p.y}`).join(" ");

    const path = document.createElementNS("http://www.w3.org/2000/svg", "polyline");

    path.setAttribute("points", points);
    path.setAttribute("stroke", stream.color);
    path.classList.add("stream-line");
    if (groupIds.has(stream.id) || isItemSelected("stream", stream.id)) {
      path.classList.add("selected");
    }
    path.dataset.streamId = stream.id;

    path.addEventListener("click", (event) => {
      event.stopPropagation();
      if (state.tool === "add-activity" || state.tool === "add-deliverable") {
        return;
      }
      if (event.ctrlKey || event.metaKey) {
        toggleSelectItem("stream", stream.id);
      } else {
        selectItem("stream", stream.id);
      }
    });

    if (state.tool === "add-stream") {
      path.addEventListener("mouseenter", () => {
        showSnapPreviewHandles(stream);
      });
      path.addEventListener("mouseleave", () => {
        clearSnapPreviewHandles();
      });
    }

    streamsLayer.appendChild(path);
  });
}

function renderHandles() {
  handlesLayer.innerHTML = "";

  if (state.selected.length !== 1 || state.selected[0].type !== "stream" || state.tool !== "select") {
    return;
  }

  const groupIds = getSelectedStreamGroup();
  const primaryId = state.selected[0].id;

  // Render non-primary streams first, primary stream last so its
  // handles sit on top at shared endpoints and receive clicks first.
  const sortedIds = [...groupIds].sort((a, b) =>
    a === primaryId ? 1 : b === primaryId ? -1 : 0
  );

  sortedIds.forEach((streamId) => {
    const stream = getStreamById(streamId);
    if (!stream) {
      return;
    }

    stream.points.forEach((p, i) => {
      const handle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      handle.setAttribute("cx", p.x);
      handle.setAttribute("cy", p.y);
      handle.setAttribute("r", 12);
      handle.classList.add("stream-handle");
      handle.dataset.streamId = stream.id;
      handle.dataset.pointIndex = i;
      handlesLayer.appendChild(handle);
    });
  });
}

function showSnapPreviewHandles(hoveredStream) {
  clearSnapPreviewHandles();
  const color = normalizeColor(hoveredStream.color);
  const sameColorStreams = state.streams.filter(
    (s) => normalizeColor(s.color) === color
  );
  sameColorStreams.forEach((stream) => {
    if (stream.points.length < 2) return;
    [stream.points[0], stream.points[stream.points.length - 1]].forEach((p) => {
      const dot = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      dot.setAttribute("cx", p.x);
      dot.setAttribute("cy", p.y);
      dot.setAttribute("r", 10);
      dot.classList.add("snap-preview-handle");
      handlesLayer.appendChild(dot);
    });
  });
}

function clearSnapPreviewHandles() {
  handlesLayer.querySelectorAll(".snap-preview-handle").forEach((el) => el.remove());
}

function renderConnections() {
  connectionsLayer.innerHTML = "";

  state.connections.forEach((conn) => {
    const a = state.nodes.find((n) => n.id === conn.fromNodeId);
    const b = state.nodes.find((n) => n.id === conn.toNodeId);
    if (!a || !b) {
      return;
    }

    const c1x = a.x + (b.x - a.x) * 0.35;
    const c2x = a.x + (b.x - a.x) * 0.65;
    const d = `M ${a.x} ${a.y} C ${c1x} ${a.y}, ${c2x} ${b.y}, ${b.x} ${b.y}`;

    const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
    path.classList.add("connection");
    if (isItemSelected("connection", conn.id)) {
      path.classList.add("selected");
    }
    path.setAttribute("d", d);
    path.dataset.connectionId = conn.id;
    connectionsLayer.appendChild(path);

    // Wider invisible hit area for clicking
    const hit = document.createElementNS("http://www.w3.org/2000/svg", "path");
    hit.classList.add("connection-hit");
    hit.setAttribute("d", d);
    hit.dataset.connectionId = conn.id;
    hit.addEventListener("click", (event) => {
      event.stopPropagation();
      if (state.tool === "select") {
        if (event.ctrlKey || event.metaKey) {
          toggleSelectItem("connection", conn.id);
        } else {
          selectItem("connection", conn.id);
        }
      }
    });
    connectionsLayer.appendChild(hit);
  });
}

function renderNodes() {
  nodesLayer.innerHTML = "";

  state.nodes.forEach((node) => {
    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.dataset.nodeId = node.id;

    const fill = STATUS_COLORS[node.status] || STATUS_COLORS.planned;

    if (node.type === "activity") {
      const circle = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      circle.setAttribute("cx", node.x);
      circle.setAttribute("cy", node.y);
      circle.setAttribute("r", 12);
      circle.setAttribute("fill", fill);
      circle.classList.add("node");
      if (isItemSelected("node", node.id)) {
        circle.classList.add("selected");
      }
      g.appendChild(circle);
    } else {
      const diamond = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      diamond.setAttribute("x", node.x - 11);
      diamond.setAttribute("y", node.y - 11);
      diamond.setAttribute("width", 22);
      diamond.setAttribute("height", 22);
      diamond.setAttribute("fill", fill);
      diamond.setAttribute("transform", `rotate(45 ${node.x} ${node.y})`);
      diamond.setAttribute("stroke", "#53565A");
      diamond.setAttribute("stroke-width", "3");
      diamond.classList.add("node");
      if (isItemSelected("node", node.id)) {
        diamond.classList.add("selected");
      }
      g.appendChild(diamond);
    }

    g.addEventListener("click", (event) => {
      event.stopPropagation();
      if (state.tool === "connect") return;
      if (event.ctrlKey || event.metaKey) {
        toggleSelectItem("node", node.id);
      } else {
        selectItem("node", node.id);
      }
    });

    g.addEventListener("mouseenter", (e) => showNodeTooltip(e, node));
    g.addEventListener("mouseleave", hideNodeTooltip);

    nodesLayer.appendChild(g);
  });
}

function formatDatePretty(dateStr) {
  if (!dateStr) return null;
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString("en-GB", { day: "numeric", month: "short", year: "numeric" });
}

function showNodeTooltip(event, node) {
  let tip = document.getElementById("nodeTooltip");
  if (!tip) {
    tip = document.createElement("div");
    tip.id = "nodeTooltip";
    tip.className = "node-tooltip";
    document.body.appendChild(tip);
  }
  const date = formatDatePretty(node.date);
  if (!date) { tip.classList.remove("visible"); return; }
  tip.innerHTML = `<span class="node-tooltip-date">${date}</span>`;
  tip.style.left = event.clientX + "px";
  tip.style.top = (event.clientY - 10) + "px";
  tip.classList.add("visible");
}

function hideNodeTooltip() {
  const tip = document.getElementById("nodeTooltip");
  if (tip) tip.classList.remove("visible");
}

function renderLabels() {
  labelsLayer.innerHTML = "";

  state.nodes.forEach((node) => {
    const mainStyle = getNodeLabelStyle(node);
    const mw = node.labelMaxWidth || DEFAULT_LABEL_MAX_WIDTH;
    const fs = labelFontSize;
    const rawSubs = node.subLabels || [];
    const subObjs = rawSubs.map((s) => typeof s === "string" ? { text: s, rightText: "" } : s);
    const isDeliv = node.type === "deliverable";
    const pos = node.labelPosition || (isDeliv ? "down" : "down");
    const allLabels = [
      { text: node.label, rightText: node.labelDescription || "", textColor: mainStyle.textColor, bgColor: mainStyle.bgColor, descPosition: node.labelDescPosition || "right" },
      ...subObjs.map((s) => ({
        text: s.text,
        rightText: s.rightText || "",
        textColor: s.textColor || mainStyle.textColor,
        bgColor: s.bgColor !== undefined ? s.bgColor : mainStyle.bgColor,
        descPosition: s.descPosition || "right",
      })),
    ];
    const labelGap = 8;

    // Compute starting position based on label position
    let curY, anchorX, textAlign;
    if (pos === "up") {
      // Render upward: we need to measure first, then shift up
      textAlign = "center";
      anchorX = node.x - mw / 2;
    } else if (pos === "left") {
      textAlign = "right";
      anchorX = node.x - 22;
      curY = node.y - fs * 0.7;
    } else if (pos === "right") {
      textAlign = "left";
      anchorX = node.x + 22;
      curY = node.y - fs * 0.7;
    } else {
      // "down" (default)
      textAlign = "center";
      anchorX = node.x - mw / 2;
      curY = node.y + 22;
    }

    const labelEntries = [];

    if (pos === "up") {
      // Render upward: measure all first, then place bottom-to-top
      const measured = [];
      allLabels.forEach((entry) => {
        const fo = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
        fo.setAttribute("x", anchorX);
        fo.setAttribute("y", 0);
        fo.setAttribute("width", mw);
        fo.setAttribute("height", 200);
        fo.setAttribute("class", "label-fo");
        const div = document.createElement("div");
        div.className = "label-wrap" + (isDeliv ? " label-deliverable" : "");
        div.style.cssText = `color:${entry.textColor};font-size:${fs}px;max-width:${mw}px;text-align:${textAlign};`;
        div.textContent = entry.text;
        fo.appendChild(div);
        labelsLayer.appendChild(fo);
        const actualH = div.offsetHeight || fs * 1.4;
        fo.setAttribute("height", actualH + 4);
        measured.push({ fo, h: actualH, entry });
      });
      // Stack from bottom (just above node) going up
      let bottomY = node.y - 22;
      for (let i = measured.length - 1; i >= 0; i--) {
        const m = measured[i];
        const y = bottomY - m.h;
        m.fo.setAttribute("y", y);
        labelEntries.push({ fo: m.fo, y, h: m.h, rightText: m.entry.rightText, bgColor: m.entry.bgColor, descPosition: m.entry.descPosition });
        bottomY = y - 6 - labelGap;
      }
    } else if (pos === "left" || pos === "right") {
      // Horizontal: measure first, then vertically center on node
      const fo = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
      if (pos === "left") {
        fo.setAttribute("x", anchorX - mw);
      } else {
        fo.setAttribute("x", anchorX);
      }
      fo.setAttribute("y", 0);
      fo.setAttribute("width", mw);
      fo.setAttribute("height", 200);
      fo.setAttribute("class", "label-fo");
      const div = document.createElement("div");
      div.className = "label-wrap" + (isDeliv ? " label-deliverable" : "");
      div.style.cssText = `color:${allLabels[0].textColor};font-size:${fs}px;max-width:${mw}px;text-align:${textAlign};`;
      div.textContent = allLabels[0].text;
      fo.appendChild(div);
      labelsLayer.appendChild(fo);
      const actualH = div.offsetHeight || fs * 1.4;
      fo.setAttribute("height", actualH + 4);
      // Vertically center on node
      const centeredY = node.y - actualH / 2;
      fo.setAttribute("y", centeredY);
      labelEntries.push({ fo, y: centeredY, h: actualH, rightText: allLabels[0].rightText, bgColor: allLabels[0].bgColor, descPosition: allLabels[0].descPosition });

      // For deliverables with stacked labels, remaining labels still go downward
      if (isDeliv && allLabels.length > 1) {
        let stackY = node.y + 22;
        for (let i = 1; i < allLabels.length; i++) {
          const entry = allLabels[i];
          const sfo = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
          sfo.setAttribute("x", node.x - mw / 2);
          sfo.setAttribute("y", stackY);
          sfo.setAttribute("width", mw);
          sfo.setAttribute("height", 200);
          sfo.setAttribute("class", "label-fo");
          const sdiv = document.createElement("div");
          sdiv.className = "label-wrap label-deliverable";
          sdiv.style.cssText = `color:${entry.textColor};font-size:${fs}px;max-width:${mw}px;text-align:center;`;
          sdiv.textContent = entry.text;
          sfo.appendChild(sdiv);
          labelsLayer.appendChild(sfo);
          const sH = sdiv.offsetHeight || fs * 1.4;
          sfo.setAttribute("height", sH + 4);
          labelEntries.push({ fo: sfo, y: stackY, h: sH, rightText: entry.rightText, bgColor: entry.bgColor, descPosition: entry.descPosition });
          stackY += sH + 6 + labelGap;
        }
      }
    } else {
      // "down" (default) — original stacking behavior
      curY = node.y + 22;
      allLabels.forEach((entry) => {
        const fo = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
        fo.setAttribute("x", node.x - mw / 2);
        fo.setAttribute("y", curY);
        fo.setAttribute("width", mw);
        fo.setAttribute("height", 200);
        fo.setAttribute("class", "label-fo");
        const div = document.createElement("div");
        div.className = "label-wrap" + (isDeliv ? " label-deliverable" : "");
        div.style.cssText = `color:${entry.textColor};font-size:${fs}px;max-width:${mw}px;text-align:center;`;
        div.textContent = entry.text;
        fo.appendChild(div);
        labelsLayer.appendChild(fo);
        const actualH = div.offsetHeight || fs * 1.4;
        fo.setAttribute("height", actualH + 4);
        labelEntries.push({ fo, y: curY, h: actualH, rightText: entry.rightText, bgColor: entry.bgColor, descPosition: entry.descPosition });
        curY += actualH + 6 + labelGap;
      });
    }

    // Background rectangles
    labelEntries.forEach((entry) => {
      if (!entry.bgColor) return;
      const foX = parseFloat(entry.fo.getAttribute("x"));
      const foW = parseFloat(entry.fo.getAttribute("width"));
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.classList.add("label-bg");
      rect.setAttribute("x", foX - 6);
      rect.setAttribute("y", entry.y - 3);
      rect.setAttribute("width", foW + 12);
      rect.setAttribute("height", entry.h + 10);
      rect.setAttribute("fill", entry.bgColor);
      labelsLayer.insertBefore(rect, entry.fo);
    });

    // Render description text based on descPosition per label
    labelEntries.forEach((entry) => {
      if (!entry.rightText) return;
      const dp = entry.descPosition || "right";
      const rfo = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
      rfo.setAttribute("class", "label-fo");
      const rdiv = document.createElement("div");
      rdiv.className = "label-wrap label-deliverable label-right-text";

      const foX = parseFloat(entry.fo.getAttribute("x"));
      const foW = parseFloat(entry.fo.getAttribute("width"));

      if (dp === "left") {
        rfo.setAttribute("x", foX - 2010);
        rfo.setAttribute("y", entry.y);
        rfo.setAttribute("width", 2000);
        rfo.setAttribute("height", entry.h + 4);
        rdiv.style.cssText = `color:#000000;font-size:${fs}px;text-align:right;`;
      } else {
        // "right" (default)
        rfo.setAttribute("x", foX + foW + 10);
        rfo.setAttribute("y", entry.y);
        rfo.setAttribute("width", 2000);
        rfo.setAttribute("height", entry.h + 4);
        rdiv.style.cssText = `color:#000000;font-size:${fs}px;text-align:left;`;
      }
      rdiv.textContent = entry.rightText;
      rfo.appendChild(rdiv);
      labelsLayer.appendChild(rfo);
    });
  });

  // Render floating text labels
  state.textLabels.forEach((tl) => {
    const mw = tl.maxWidth || 160;
    const fs = tl.fontSize || 16;
    const align = tl.align || "center";

    const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
    g.dataset.textLabelId = tl.id;
    g.classList.add("text-label-group");
    if (isItemSelected("text-label", tl.id)) {
      g.classList.add("selected");
    }

    const fo = document.createElementNS("http://www.w3.org/2000/svg", "foreignObject");
    fo.setAttribute("x", tl.x - mw / 2);
    fo.setAttribute("y", tl.y);
    fo.setAttribute("width", mw);
    fo.setAttribute("height", 400);
    fo.setAttribute("class", "label-fo");

    const div = document.createElement("div");
    div.className = "text-label-content";
    div.style.cssText = `color:${tl.textColor || "#ffffff"};font-size:${fs}px;max-width:${mw}px;text-align:${align};`;
    div.textContent = tl.text;
    fo.appendChild(div);
    g.appendChild(fo);
    labelsLayer.appendChild(g);

    const actualH = div.offsetHeight || fs * 1.4;
    fo.setAttribute("height", actualH + 4);

    if (tl.bgColor) {
      const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
      rect.classList.add("label-bg");
      rect.setAttribute("x", tl.x - mw / 2 - 6);
      rect.setAttribute("y", tl.y - 3);
      rect.setAttribute("width", mw + 12);
      rect.setAttribute("height", actualH + 10);
      rect.setAttribute("fill", tl.bgColor);
      g.insertBefore(rect, fo);
    }

    // Hit area for selection / drag
    const hitRect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
    hitRect.setAttribute("x", tl.x - mw / 2);
    hitRect.setAttribute("y", tl.y);
    hitRect.setAttribute("width", mw);
    hitRect.setAttribute("height", Math.max(actualH + 4, 20));
    hitRect.setAttribute("fill", "transparent");
    hitRect.classList.add("text-label-hit");
    g.appendChild(hitRect);

    g.addEventListener("click", (event) => {
      event.stopPropagation();
      if (event.ctrlKey || event.metaKey) {
        toggleSelectItem("text-label", tl.id);
      } else {
        selectItem("text-label", tl.id);
      }
    });

    labelsLayer.appendChild(g);
  });
}

function getOrderedStreamGroups() {
  const groupedByColor = new Map();
  state.streams.forEach((stream) => {
    const colorKey = normalizeColor(stream.color);
    const existing = groupedByColor.get(colorKey);
    if (!existing) {
      groupedByColor.set(colorKey, { color: stream.color, name: stream.name, colorKey });
    }
  });

  // Build ordered list: known order first, then any new colors
  const ordered = [];
  const used = new Set();
  (state.streamOrder || []).forEach((key) => {
    const entry = groupedByColor.get(key);
    if (entry) { ordered.push(entry); used.add(key); }
  });
  groupedByColor.forEach((entry, key) => {
    if (!used.has(key)) ordered.push(entry);
  });

  // Sync streamOrder to match current reality
  state.streamOrder = ordered.map((e) => e.colorKey);
  return ordered;
}

function renderLegend() {
  streamLegend.innerHTML = "";
  const ordered = getOrderedStreamGroups();

  ordered.forEach((entry, idx) => {
    const li = document.createElement("li");
    li.setAttribute("draggable", "true");
    li.dataset.colorKey = entry.colorKey;
    li.dataset.idx = idx;

    const grip = document.createElement("span");
    grip.className = "legend-drag-grip";
    grip.textContent = "⠿";
    li.appendChild(grip);

    const swatch = document.createElement("span");
    swatch.className = "legend-stream-color";
    swatch.style.background = entry.color;
    li.appendChild(swatch);
    li.appendChild(document.createTextNode(entry.name));

    // Drag events
    li.addEventListener("dragstart", (e) => {
      li.classList.add("legend-dragging");
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", String(idx));
    });
    li.addEventListener("dragend", () => {
      li.classList.remove("legend-dragging");
      streamLegend.querySelectorAll(".legend-drag-over").forEach((el) => el.classList.remove("legend-drag-over"));
    });
    li.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";
      li.classList.add("legend-drag-over");
    });
    li.addEventListener("dragleave", () => {
      li.classList.remove("legend-drag-over");
    });
    li.addEventListener("drop", (e) => {
      e.preventDefault();
      li.classList.remove("legend-drag-over");
      const fromIdx = parseInt(e.dataTransfer.getData("text/plain"), 10);
      const toIdx = idx;
      if (fromIdx === toIdx || isNaN(fromIdx)) return;
      const order = [...state.streamOrder];
      const [moved] = order.splice(fromIdx, 1);
      order.splice(toIdx, 0, moved);
      state.streamOrder = order;
      renderLegend();
    });

    streamLegend.appendChild(li);
  });
}

function renderSelectionEditor() {
  if (state.selected.length === 0) {
    selectionEditor.className = "selection-editor muted";
    selectionEditor.textContent = "Nothing selected.";
    return;
  }

  selectionEditor.className = "selection-editor";

  if (state.selected.length > 1) {
    selectionEditor.innerHTML = `
      <p><strong>${state.selected.length} items selected</strong></p>
      <p class="muted">Drag to move all. Press Delete to remove.</p>
    `;
    return;
  }

  const selected = getSelectedEntity();
  if (!selected) {
    selectionEditor.className = "selection-editor muted";
    selectionEditor.textContent = "Nothing selected.";
    return;
  }

  if (selected.type === "connection") {
    const conn = selected.entity;
    const fromNode = state.nodes.find((n) => n.id === conn.fromNodeId);
    const toNode = state.nodes.find((n) => n.id === conn.toNodeId);
    const fromLabel = fromNode ? fromNode.label : "?";
    const toLabel = toNode ? toNode.label : "?";
    selectionEditor.innerHTML = `
      <p><strong>Connection</strong></p>
      <p class="muted">${escapeHtml(fromLabel)} → ${escapeHtml(toLabel)}</p>
      <p class="muted">Press Delete to remove.</p>
    `;
    return;
  }

  if (selected.type === "stream") {
    const stream = selected.entity;
    const groupIds = getSelectedStreamGroup();
    const groupStreams = state.streams.filter((s) => groupIds.has(s.id));
    const segmentCount = groupStreams.length;
    selectionEditor.innerHTML = `
      <label>Stream Name <input id="editStreamName" value="${escapeHtml(stream.name)}" /></label>
      <label>Stream Color <input id="editStreamColor" type="color" value="${stream.color}" /></label>
      <p class="muted">${segmentCount > 1 ? `Group of ${segmentCount} connected segments. Edits apply to all.` : "Tip: Snap endpoints to same-color streams to form a group."}</p>
    `;

    document.getElementById("editStreamName").addEventListener("input", (event) => {
      if (stream.name === event.target.value) {
        return;
      }
      captureHistorySnapshot();
      groupStreams.forEach((s) => { s.name = event.target.value; });
      render();
      syncStreamPresetDropdown();
    });

    document.getElementById("editStreamColor").addEventListener("input", (event) => {
      if (stream.color === event.target.value) {
        return;
      }
      captureHistorySnapshot();
      groupStreams.forEach((s) => { s.color = event.target.value; });
      render();
      syncStreamPresetDropdown();
    });
    return;
  }

  if (selected.type === "text-label") {
    const tl = selected.entity;
    const hasBg = !!tl.bgColor;
    selectionEditor.innerHTML = `
      <label>Text <textarea id="editTLText" rows="2">${escapeHtml(tl.text)}</textarea></label>
      <label class="range-label">Font Size <span id="editTLFontSizeVal">${tl.fontSize}px</span>
        <input id="editTLFontSize" type="range" min="8" max="60" step="1" value="${tl.fontSize}" />
      </label>
      <label class="range-label">Max Width <span id="editTLMaxWidthVal">${tl.maxWidth}px</span>
        <input id="editTLMaxWidth" type="range" min="40" max="600" step="10" value="${tl.maxWidth}" />
      </label>
      <label>Alignment</label>
      <div class="align-btn-row">
        <button class="align-btn${tl.align === "left" ? " active" : ""}" data-align="left" title="Left">
          <svg width="16" height="16" viewBox="0 0 16 16"><rect x="1" y="2" width="10" height="2" fill="currentColor"/><rect x="1" y="6" width="14" height="2" fill="currentColor"/><rect x="1" y="10" width="8" height="2" fill="currentColor"/><rect x="1" y="14" width="12" height="2" fill="currentColor"/></svg>
        </button>
        <button class="align-btn${tl.align === "center" ? " active" : ""}" data-align="center" title="Center">
          <svg width="16" height="16" viewBox="0 0 16 16"><rect x="3" y="2" width="10" height="2" fill="currentColor"/><rect x="1" y="6" width="14" height="2" fill="currentColor"/><rect x="4" y="10" width="8" height="2" fill="currentColor"/><rect x="2" y="14" width="12" height="2" fill="currentColor"/></svg>
        </button>
        <button class="align-btn${tl.align === "right" ? " active" : ""}" data-align="right" title="Right">
          <svg width="16" height="16" viewBox="0 0 16 16"><rect x="5" y="2" width="10" height="2" fill="currentColor"/><rect x="1" y="6" width="14" height="2" fill="currentColor"/><rect x="7" y="10" width="8" height="2" fill="currentColor"/><rect x="3" y="14" width="12" height="2" fill="currentColor"/></svg>
        </button>
      </div>
      <label>Text Color <input id="editTLTextColor" type="color" value="${tl.textColor || "#ffffff"}" /></label>
      <div class="editor-row">
        <label class="editor-row-field">Background <input id="editTLBgColor" type="color" value="${tl.bgColor || "#000000"}" ${hasBg ? "" : "disabled"} /></label>
        <button id="editTLBgEnabled" class="toggle-btn${hasBg ? " active" : ""}">${hasBg ? "On" : "Off"}</button>
      </div>
    `;

    document.getElementById("editTLText").addEventListener("input", (e) => {
      captureHistorySnapshot(); tl.text = e.target.value; render();
    });
    document.getElementById("editTLFontSize").addEventListener("input", (e) => {
      tl.fontSize = Number(e.target.value);
      document.getElementById("editTLFontSizeVal").textContent = tl.fontSize + "px";
      render();
    });
    document.getElementById("editTLFontSize").addEventListener("change", () => { captureHistorySnapshot(); });
    document.getElementById("editTLMaxWidth").addEventListener("input", (e) => {
      tl.maxWidth = Number(e.target.value);
      document.getElementById("editTLMaxWidthVal").textContent = tl.maxWidth + "px";
      render();
    });
    document.getElementById("editTLMaxWidth").addEventListener("change", () => { captureHistorySnapshot(); });
    document.querySelectorAll(".align-btn-row .align-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        captureHistorySnapshot();
        tl.align = btn.dataset.align;
        renderSelectionEditor();
        render();
      });
    });
    document.getElementById("editTLTextColor").addEventListener("input", (e) => {
      captureHistorySnapshot(); tl.textColor = e.target.value; render();
    });
    document.getElementById("editTLBgColor").addEventListener("input", (e) => {
      if (!tl.bgColor) return;
      captureHistorySnapshot(); tl.bgColor = e.target.value; render();
    });
    document.getElementById("editTLBgEnabled").addEventListener("click", () => {
      const isOn = !tl.bgColor;
      captureHistorySnapshot();
      tl.bgColor = isOn ? (tl.bgColor || "#000000") : null;
      renderSelectionEditor();
      render();
    });
    return;
  }

  const node = selected.entity;
  const nodeMaxW = node.labelMaxWidth || DEFAULT_LABEL_MAX_WIDTH;
  const isDeliverable = node.type === "deliverable";

  // Build unified labels array: [main label, ...subLabels]
  // Each entry: { text, rightText, textColor, bgColor }
  function getAllLabels() {
    const mainStyle = getNodeLabelStyle(node);
    const main = {
      text: node.label,
      rightText: node.labelDescription || "",
      textColor: mainStyle.textColor,
      bgColor: mainStyle.bgColor,
      descPosition: node.labelDescPosition || "right",
    };
    const subs = (node.subLabels || []).map((s) => {
      if (typeof s === "string") s = { text: s, rightText: "" };
      return {
        text: s.text,
        rightText: s.rightText || "",
        textColor: s.textColor || (node.type === "deliverable" ? "#ffffff" : "#000000"),
        bgColor: s.bgColor !== undefined ? s.bgColor : (node.type === "deliverable" ? "#000000" : null),
        descPosition: s.descPosition || "right",
      };
    });
    return [main, ...subs];
  }

  function writeAllLabels(arr) {
    // arr[0] is the main label, rest are subLabels
    const m = arr[0];
    node.label = m.text;
    node.labelDescription = m.rightText;
    node.labelTextColor = m.textColor;
    node.labelBgColor = m.bgColor;
    node.labelDescPosition = m.descPosition || "right";
    node.subLabels = arr.slice(1).map((s) => ({
      text: s.text,
      rightText: s.rightText || "",
      textColor: s.textColor,
      bgColor: s.bgColor,
      descPosition: s.descPosition || "right",
    }));
  }

  const allLabels = getAllLabels();

  let labelsHtml = "";
  if (isDeliverable) {
    const labelCards = allLabels.map((lbl, i) => {
      const hasBg = !!lbl.bgColor;
      const canDelete = allLabels.length > 1;
      return `
      <div class="sub-label-stack" data-label-idx="${i}">
        <div class="sub-label-drag-header">
          <span class="sub-label-grip">⠿</span>
          <span class="sub-label-title">Label ${i + 1}</span>
          ${canDelete ? `<button class="sub-label-remove danger" data-label-idx="${i}" title="Remove">&times;</button>` : ""}
        </div>
        <label class="sub-label-field">Deliverable ID <input class="label-text-input" value="${escapeHtml(lbl.text)}" data-label-idx="${i}" /></label>
        <label class="sub-label-field">Description <input class="label-right-input" value="${escapeHtml(lbl.rightText)}" data-label-idx="${i}" /></label>
        <div class="sub-label-pos-row">
          <span class="sub-label-pos-label">Label Position</span>
          <div class="pos-btn-row pos-btn-full">
            <button class="pos-btn${lbl.descPosition === "left" ? " active" : ""}" data-label-idx="${i}" data-pos="left">Left</button>
            <button class="pos-btn${lbl.descPosition === "right" || !lbl.descPosition ? " active" : ""}" data-label-idx="${i}" data-pos="right">Right</button>
          </div>
        </div>
        <div class="sub-label-color-section">
          <span class="sub-label-color-heading">Primary Color</span>
          <input type="color" class="label-text-color sub-label-color-input" value="${lbl.textColor}" data-label-idx="${i}" />
        </div>
        <div class="sub-label-color-section">
          <span class="sub-label-color-heading">Secondary Color <button type="button" class="label-bg-toggle" data-label-idx="${i}" title="${hasBg ? "Disable" : "Enable"}"><span class="bg-toggle-track${hasBg ? " on" : ""}"><span class="bg-toggle-thumb"></span></span></button></span>
          <input type="color" class="label-bg-color sub-label-color-input${hasBg ? "" : " color-disabled"}" value="${lbl.bgColor || (isDeliverable ? "#000000" : "#ffffff")}" data-label-idx="${i}" ${hasBg ? "" : "disabled"} />
        </div>
      </div>`;
    }).join("");
    labelsHtml = `
      <div class="sub-labels-section" id="labelsContainer">
        ${labelCards}
        <button id="addSubLabel" class="sub-label-add">+ Add Label</button>
      </div>`;
  }

  const actLabelPos = node.labelPosition || "down";

  selectionEditor.innerHTML = `
    ${isDeliverable ? "" : `<label>Label <input id="editNodeLabel" value="${escapeHtml(node.label)}" /></label>`}
    ${isDeliverable ? labelsHtml : ""}
    ${isDeliverable ? "" : `
    <label>Label Position</label>
    <div class="pos-btn-row pos-btn-full" id="labelPosRow">
      <button class="pos-btn${actLabelPos === "up" ? " active" : ""}" data-pos="up">Up</button>
      <button class="pos-btn${actLabelPos === "down" ? " active" : ""}" data-pos="down">Down</button>
      <button class="pos-btn${actLabelPos === "left" ? " active" : ""}" data-pos="left">Left</button>
      <button class="pos-btn${actLabelPos === "right" ? " active" : ""}" data-pos="right">Right</button>
    </div>`}
    <label>Date <input id="editNodeDate" type="date" value="${node.date}" /></label>
    <label class="range-label">Label Max Width <span id="editNodeMaxWidthVal">${nodeMaxW}px</span>
      <input id="editNodeMaxWidth" type="range" min="40" max="400" step="10" value="${nodeMaxW}" />
    </label>
    ${isDeliverable ? "" : `
    <div class="sub-label-color-section">
      <span class="sub-label-color-heading">Primary Color</span>
      <input id="editNodeTextColor" type="color" class="sub-label-color-input" value="${getNodeLabelStyle(node).textColor}" />
    </div>
    <div class="sub-label-color-section">
      <span class="sub-label-color-heading">Secondary Color <button type="button" id="editNodeLabelBgEnabled" class="label-bg-toggle" title="${getNodeLabelStyle(node).bgColor ? "Disable" : "Enable"}"><span class="bg-toggle-track${getNodeLabelStyle(node).bgColor ? " on" : ""}"><span class="bg-toggle-thumb"></span></span></button></span>
      <input id="editNodeLabelBgColor" type="color" class="sub-label-color-input${getNodeLabelStyle(node).bgColor ? "" : " color-disabled"}" value="${(getNodeLabelStyle(node).bgColor || "#ffffff")}" ${getNodeLabelStyle(node).bgColor ? "" : "disabled"} />
    </div>`}
    <label>Status
      <select id="editNodeStatus">
        <option value="planned">Planned Activity</option>
        <option value="completed">Completed</option>
        <option value="ongoing">On-going</option>
        <option value="attention">Needs Attention</option>
      </select>
    </label>
    <label>Stream
      <select id="editNodeStream"></select>
    </label>
  `;

  // --- Activity label input ---
  if (!isDeliverable) {
    document.getElementById("editNodeLabel").addEventListener("input", (event) => {
      if (node.label === event.target.value) return;
      captureHistorySnapshot();
      node.label = event.target.value;
      render();
    });
    // Activity label position selector
    document.querySelectorAll("#labelPosRow .pos-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const newPos = btn.dataset.pos;
        if (node.labelPosition === newPos) return;
        captureHistorySnapshot();
        node.labelPosition = newPos;
        renderSelectionEditor();
        render();
      });
    });
  }

  // --- Deliverable label editing ---
  if (isDeliverable) {
    // Normalize subLabels to objects
    if (node.subLabels) {
      node.subLabels = node.subLabels.map((s) => typeof s === "string" ? { text: s, rightText: "" } : s);
    }

    // Text inputs
    document.querySelectorAll(".label-text-input").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const idx = Number(e.target.dataset.labelIdx);
        const labels = getAllLabels();
        if (labels[idx].text === e.target.value) return;
        captureHistorySnapshot();
        labels[idx].text = e.target.value;
        writeAllLabels(labels);
        render();
      });
    });

    // Right-text (description) inputs
    document.querySelectorAll(".label-right-input").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const idx = Number(e.target.dataset.labelIdx);
        const labels = getAllLabels();
        if (labels[idx].rightText === e.target.value) return;
        captureHistorySnapshot();
        labels[idx].rightText = e.target.value;
        writeAllLabels(labels);
        render();
      });
    });

    // Per-label text color
    document.querySelectorAll(".label-text-color").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const idx = Number(e.target.dataset.labelIdx);
        const labels = getAllLabels();
        captureHistorySnapshot();
        labels[idx].textColor = e.target.value;
        writeAllLabels(labels);
        render();
      });
    });

    // Per-label background color
    document.querySelectorAll(".label-bg-color").forEach((inp) => {
      inp.addEventListener("input", (e) => {
        const idx = Number(e.target.dataset.labelIdx);
        const labels = getAllLabels();
        if (!labels[idx].bgColor) return;
        captureHistorySnapshot();
        labels[idx].bgColor = e.target.value;
        writeAllLabels(labels);
        render();
      });
    });

    // Per-label background toggle
    document.querySelectorAll(".label-bg-toggle").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.labelIdx);
        const labels = getAllLabels();
        const isOn = !labels[idx].bgColor;
        captureHistorySnapshot();
        labels[idx].bgColor = isOn ? (isDeliverable ? "#000000" : "#ffffff") : null;
        writeAllLabels(labels);
        renderSelectionEditor();
        render();
      });
    });

    // Per-label description position
    document.querySelectorAll(".sub-label-pos-row .pos-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.labelIdx);
        const newPos = btn.dataset.pos;
        const labels = getAllLabels();
        if (labels[idx].descPosition === newPos) return;
        captureHistorySnapshot();
        labels[idx].descPosition = newPos;
        writeAllLabels(labels);
        renderSelectionEditor();
        render();
      });
    });

    // Remove label
    document.querySelectorAll(".sub-label-remove").forEach((btn) => {
      btn.addEventListener("click", () => {
        const idx = Number(btn.dataset.labelIdx);
        const labels = getAllLabels();
        if (labels.length <= 1) return;
        captureHistorySnapshot();
        labels.splice(idx, 1);
        writeAllLabels(labels);
        renderSelectionEditor();
        render();
      });
    });

    // Add label
    const addBtn = document.getElementById("addSubLabel");
    if (addBtn) {
      addBtn.addEventListener("click", () => {
        captureHistorySnapshot();
        if (!node.subLabels) node.subLabels = [];
        node.subLabels.push({ text: "Deliverable", rightText: "", textColor: "#ffffff", bgColor: "#000000", descPosition: "right" });
        renderSelectionEditor();
        render();
      });
    }

    // Drag to reorder labels — only from grip handle
    const container = document.getElementById("labelsContainer");
    if (container) {
      let dragIdx = null;
      // Enable draggable on the card only while grip is held
      container.querySelectorAll(".sub-label-grip").forEach((grip) => {
        grip.addEventListener("mousedown", () => {
          const card = grip.closest(".sub-label-stack");
          if (card) card.setAttribute("draggable", "true");
        });
      });
      // Remove draggable after drag ends or mouse releases elsewhere
      const disableDrag = () => {
        container.querySelectorAll(".sub-label-stack").forEach((c) => c.removeAttribute("draggable"));
      };
      container.addEventListener("dragstart", (e) => {
        const card = e.target.closest(".sub-label-stack[data-label-idx]");
        if (!card) return;
        dragIdx = Number(card.dataset.labelIdx);
        card.classList.add("label-dragging");
        e.dataTransfer.effectAllowed = "move";
      });
      container.addEventListener("dragover", (e) => {
        e.preventDefault();
        e.dataTransfer.dropEffect = "move";
        const card = e.target.closest(".sub-label-stack[data-label-idx]");
        container.querySelectorAll(".sub-label-stack").forEach((c) => c.classList.remove("label-drag-over"));
        if (card) card.classList.add("label-drag-over");
      });
      container.addEventListener("dragleave", (e) => {
        const card = e.target.closest(".sub-label-stack[data-label-idx]");
        if (card) card.classList.remove("label-drag-over");
      });
      container.addEventListener("drop", (e) => {
        e.preventDefault();
        container.querySelectorAll(".sub-label-stack").forEach((c) => { c.classList.remove("label-drag-over"); c.classList.remove("label-dragging"); });
        const card = e.target.closest(".sub-label-stack[data-label-idx]");
        if (!card || dragIdx === null) return;
        const dropIdx = Number(card.dataset.labelIdx);
        if (dropIdx === dragIdx) return;
        captureHistorySnapshot();
        const labels = getAllLabels();
        const [moved] = labels.splice(dragIdx, 1);
        labels.splice(dropIdx, 0, moved);
        writeAllLabels(labels);
        renderSelectionEditor();
        render();
      });
      container.addEventListener("dragend", () => {
        container.querySelectorAll(".sub-label-stack").forEach((c) => { c.classList.remove("label-drag-over"); c.classList.remove("label-dragging"); });
        disableDrag();
        dragIdx = null;
      });
      document.addEventListener("mouseup", disableDrag);
    }
  }

  document.getElementById("editNodeDate").addEventListener("input", (event) => {
    if (node.date === event.target.value) {
      return;
    }
    captureHistorySnapshot();
    node.date = event.target.value;
    render();
  });
  const maxWidthSlider = document.getElementById("editNodeMaxWidth");
  let maxWidthSnapshotTaken = false;
  maxWidthSlider.addEventListener("pointerdown", () => { if (!maxWidthSnapshotTaken) { captureHistorySnapshot(); maxWidthSnapshotTaken = true; } });
  maxWidthSlider.addEventListener("input", (event) => {
    node.labelMaxWidth = Number(event.target.value);
    document.getElementById("editNodeMaxWidthVal").textContent = node.labelMaxWidth + "px";
    render();
  });
  maxWidthSlider.addEventListener("pointerup", () => { maxWidthSnapshotTaken = false; });
  const textColorEl = document.getElementById("editNodeTextColor");
  if (textColorEl) {
    textColorEl.addEventListener("input", (event) => {
      if ((node.labelTextColor || (node.type === "deliverable" ? "#ffffff" : "#000000")) === event.target.value) {
        return;
      }
      captureHistorySnapshot();
      node.labelTextColor = event.target.value;
      render();
    });
  }
  const bgColorEl = document.getElementById("editNodeLabelBgColor");
  if (bgColorEl) {
    bgColorEl.addEventListener("input", (event) => {
      if (!node.labelBgColor) {
        return;
      }
      if (node.labelBgColor === event.target.value) {
        return;
      }
      captureHistorySnapshot();
      node.labelBgColor = event.target.value;
      render();
    });
  }
  const bgEnabledEl = document.getElementById("editNodeLabelBgEnabled");
  if (bgEnabledEl) {
    bgEnabledEl.addEventListener("click", (event) => {
      const isOn = !getNodeLabelStyle(node).bgColor;
      const nextValue = isOn
        ? node.labelBgColor || "#000000"
        : null;
      if ((node.labelBgColor || null) === nextValue) {
        return;
      }
      captureHistorySnapshot();
      node.labelBgColor = nextValue;
      renderSelectionEditor();
      render();
    });
  }
  document.getElementById("editNodeStatus").value = node.status;
  document.getElementById("editNodeStatus").addEventListener("change", (event) => {
    if (node.status === event.target.value) {
      return;
    }
    captureHistorySnapshot();
    node.status = event.target.value;
    render();
  });

  const streamSelect = document.getElementById("editNodeStream");
  state.streams.forEach((stream) => {
    const option = document.createElement("option");
    option.value = stream.id;
    option.textContent = stream.name;
    streamSelect.appendChild(option);
  });
  streamSelect.value = node.streamId;
  streamSelect.addEventListener("change", (event) => {
    if (node.streamId === event.target.value) {
      return;
    }
    captureHistorySnapshot();
    node.streamId = event.target.value;
    if (node.streamId) {
      const constrained = constrainActivityToStream({ x: node.x, y: node.y }, node.streamId);
      node.x = constrained.x;
      node.y = constrained.y;
    }
    render();
  });
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function isInteractiveTarget(target) {
  return Boolean(target.closest(".node") || target.closest(".stream-line") || target.closest(".stream-handle") || target.closest(".guide-hit") || target.closest(".connection-hit") || target.closest(".text-label-group"));
}

function startPan(event) {
  state.pan.active = true;
  state.pan.startX = event.clientX;
  state.pan.startY = event.clientY;
  state.pan.baseX = state.transform.x;
  state.pan.baseY = state.transform.y;
  state.interaction.mode = "pan";
}

function handlePointerDown(event) {
  if (event.button !== 0 && event.button !== 1) {
    return;
  }

  const point = snapPoint(svgPointFromEvent(event));
  state.interaction.dragMoved = false;

  const handle = event.target.closest(".stream-handle");
  if (state.tool === "select" && handle && event.button === 0) {
    state.interaction.mode = "handle-drag";
    state.interaction.dragStreamId = handle.dataset.streamId;
    state.interaction.dragPointIndex = Number(handle.dataset.pointIndex);
    state.interaction.historyCaptured = false;
    state.interaction.startPoint = svgPointFromEvent(event);
    // Precompute shift-constrain direction from original segment
    const hStream = getStreamById(handle.dataset.streamId);
    const hIdx = state.interaction.dragPointIndex;
    if (hStream && hStream.points.length >= 2) {
      const hAdjIdx = hIdx === 0 ? 1 : hIdx - 1;
      const hAnchor = hStream.points[hAdjIdx];
      const hDragged = hStream.points[hIdx];
      const odx = hDragged.x - hAnchor.x;
      const ody = hDragged.y - hAnchor.y;
      const olen = Math.sqrt(odx * odx + ody * ody);
      if (olen > 0) {
        state.interaction.shiftDir = { x: odx / olen, y: ody / olen };
        state.interaction.shiftAnchor = { x: hAnchor.x, y: hAnchor.y };
      }
    }
    return;
  }

  const guideHit = event.target.closest(".guide-hit");
  if (state.tool === "select" && guideHit && event.button === 0) {
    const guide = state.guides.find((g) => g.id === guideHit.dataset.guideId);
    if (guide) {
      state.interaction.mode = "guide-drag";
      state.interaction.dragGuideId = guide.id;
      state.interaction.historyCaptured = false;
      state.interaction.startPoint = svgPointFromEvent(event);
      return;
    }
  }

  const nodeGroup = event.target.closest("g[data-node-id]");
  if (state.tool === "connect" && nodeGroup && event.button === 0) {
    const fromNode = state.nodes.find((n) => n.id === nodeGroup.dataset.nodeId);
    if (fromNode) {
      state.interaction.mode = "connect-drag";
      state.interaction.connectFromNodeId = fromNode.id;
      state.interaction.connectHoverNodeId = null;
      state.interaction.startPoint = { x: fromNode.x, y: fromNode.y };
      state.interaction.previewPoint = { x: fromNode.x, y: fromNode.y };
      return;
    }
  }

  if (state.tool === "select" && nodeGroup && event.button === 0) {
    const nodeId = nodeGroup.dataset.nodeId;
    // Ctrl+click: let the click handler toggle selection
    if (event.ctrlKey || event.metaKey) {
      return;
    }
    // If this node is part of a multi-selection, start multi-drag
    if (isMultiSelection() && isItemSelected("node", nodeId)) {
      const origItems = buildMultiDragOrigItems();
      state.interaction.mode = "multi-drag";
      state.interaction.historyCaptured = false;
      state.interaction.startPoint = svgPointFromEvent(event);
      state.interaction.multiDragOrigItems = origItems;
      return;
    }
    state.interaction.mode = "node-drag";
    state.interaction.historyCaptured = false;
    state.interaction.dragNodeId = nodeId;
    state.interaction.startPoint = point;
    state.interaction.previewPoint = point;
    return;
  }

  const textLabelGroup = event.target.closest("g[data-text-label-id]");
  if (state.tool === "select" && textLabelGroup && event.button === 0) {
    const tlId = textLabelGroup.dataset.textLabelId;
    if (event.ctrlKey || event.metaKey) {
      return;
    }
    if (isMultiSelection() && isItemSelected("text-label", tlId)) {
      const origItems = buildMultiDragOrigItems();
      state.interaction.mode = "multi-drag";
      state.interaction.historyCaptured = false;
      state.interaction.startPoint = svgPointFromEvent(event);
      state.interaction.multiDragOrigItems = origItems;
      return;
    }
    state.interaction.mode = "text-label-drag";
    state.interaction.historyCaptured = false;
    state.interaction.dragTextLabelId = tlId;
    state.interaction.startPoint = point;
    state.interaction.previewPoint = point;
    return;
  }

  const streamLine = event.target.closest(".stream-line");
  if (state.tool === "select" && streamLine && event.button === 0) {
    const streamId = streamLine.dataset.streamId;
    // Ctrl+click: let the click handler toggle selection
    if (event.ctrlKey || event.metaKey) {
      return;
    }
    // If this stream is part of a multi-selection, start multi-drag
    if (isMultiSelection() && isItemSelected("stream", streamId)) {
      const origItems = buildMultiDragOrigItems();
      state.interaction.mode = "multi-drag";
      state.interaction.historyCaptured = false;
      state.interaction.startPoint = svgPointFromEvent(event);
      state.interaction.multiDragOrigItems = origItems;
      return;
    }
    // If this is a selected but ungrouped stream (e.g. ctrl+clicked), start multi-drag
    if (state.selected.length === 1 && isItemSelected("stream", streamId) && !state.selectionGrouped) {
      const origItems = buildMultiDragOrigItems();
      state.interaction.mode = "multi-drag";
      state.interaction.historyCaptured = false;
      state.interaction.startPoint = svgPointFromEvent(event);
      state.interaction.multiDragOrigItems = origItems;
      return;
    }
    const groupIds = getSelectedStreamGroup();
    if (groupIds.has(streamId)) {
      const allOrigPoints = [];
      const allOrigNodes = [];
      groupIds.forEach((gid) => {
        const s = getStreamById(gid);
        if (s) {
          allOrigPoints.push({ streamId: gid, points: cloneData(s.points) });
        }
        state.nodes.filter((n) => n.streamId === gid).forEach((n) => {
          allOrigNodes.push({ id: n.id, x: n.x, y: n.y });
        });
      });
      state.interaction.mode = "stream-drag";
      state.interaction.dragStreamId = streamId;
      state.interaction.historyCaptured = false;
      state.interaction.startPoint = svgPointFromEvent(event);
      state.interaction.dragStreamOrigPoints = allOrigPoints;
      state.interaction.dragStreamOrigNodes = allOrigNodes;
      return;
    }
  }

  if (state.tool === "add-stream" && event.button === 0) {
    const clickedStream = event.target.closest(".stream-line");
    let startPt;
    if (clickedStream) {
      const stream = getStreamById(clickedStream.dataset.streamId);
      if (stream) {
        const raw = svgPointFromEvent(event);
        const endpoints = [stream.points[0], stream.points[stream.points.length - 1]];
        let snappedToEndpoint = false;
        for (const ep of endpoints) {
          if ((raw.x - ep.x) ** 2 + (raw.y - ep.y) ** 2 <= HANDLE_SNAP_DISTANCE * HANDLE_SNAP_DISTANCE) {
            startPt = { x: ep.x, y: ep.y };
            snappedToEndpoint = true;
            break;
          }
        }
        if (!snappedToEndpoint) {
          const nearest = nearestPointOnStream(raw, stream);
          startPt = nearest || snapPoint(raw);
        }
      } else {
        startPt = point;
      }
    } else {
      startPt = point;
    }
    state.interaction.mode = "stream-create";
    state.interaction.startPoint = startPt;
    state.interaction.previewPoint = startPt;
    render();
    return;
  }

  if (
    (state.tool === "add-activity" || state.tool === "add-deliverable") &&
    event.button === 0
  ) {
    const clickedStreamId = event.target.closest(".stream-line")?.dataset.streamId;
    if (!clickedStreamId) {
      if (!state.streams.length) {
        showFloatingError("Create at least one stream first.");
      } else {
        showFloatingError("Click on a stream to place the node.");
      }
      return;
    }
    const constrainedStart = constrainActivityToStream(svgPointFromEvent(event), clickedStreamId);
    state.interaction.mode = "node-create";
    state.interaction.nodeCreateStreamId = clickedStreamId;
    state.interaction.startPoint = constrainedStart;
    state.interaction.previewPoint = constrainedStart;
    render();
    return;
  }

  if (state.tool === "add-text-label" && event.button === 0) {
    const rawPoint = svgPointFromEvent(event);
    const tl = createTextLabel({ x: rawPoint.x, y: rawPoint.y });
    selectItem("text-label", tl.id);
    setTool("select");
    return;
  }

  if (!isInteractiveTarget(event.target)) {
    if (state.tool === "select" && event.button === 0) {
      // Start box-select
      const rawPoint = svgPointFromEvent(event);
      if (!(event.ctrlKey || event.metaKey)) {
        clearSelection();
      }
      state.interaction.mode = "box-select";
      state.interaction.boxSelectStart = rawPoint;
      state.interaction.boxSelectEnd = rawPoint;
      state.interaction.startPoint = rawPoint;
      state.interaction.previewPoint = rawPoint;
      // Also set up pan data so we can fall back if needed
      state.pan.startX = event.clientX;
      state.pan.startY = event.clientY;
      return;
    }
    startPan(event);
  }
}

function trackPointer(event) { lastPointerEvent = event; }

function handlePointerMove(event) {
  const mode = state.interaction.mode;
  if (!mode) {
    return;
  }

  const rawCurrent = svgPointFromEvent(event);
  const snappedCurrent = snapPoint(rawCurrent);

  if (
    Math.abs(event.clientX - state.pan.startX) > 1 ||
    Math.abs(event.clientY - state.pan.startY) > 1 ||
    (state.interaction.startPoint &&
      (snappedCurrent.x !== state.interaction.startPoint.x ||
        snappedCurrent.y !== state.interaction.startPoint.y))
  ) {
    state.interaction.dragMoved = true;
  }

  if (mode === "pan") {
    const dx = event.clientX - state.pan.startX;
    const dy = event.clientY - state.pan.startY;
    state.transform.x = state.pan.baseX + dx;
    state.transform.y = state.pan.baseY + dy;
    applyViewportTransform();
    renderCalendar();
    return;
  }

  if (mode === "calendar-border-drag") {
    if (!state.interaction.historyCaptured) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    const { svgScale } = getSvgScreenMapping();
    const deltaScreenX = event.clientX - state.interaction.calendarDragStartX;
    const deltaVbX = deltaScreenX / svgScale;
    const deltaSvgX = deltaVbX / state.transform.scale;
    const newWidth = Math.max(60, Math.round(state.interaction.calendarDragStartWidth + deltaSvgX));
    state.calendar.monthWidths[state.interaction.calendarDragKey] = newWidth;
    renderCalendar();
    return;
  }

  if (mode === "box-select") {
    state.interaction.boxSelectEnd = rawCurrent;
    state.interaction.previewPoint = rawCurrent;
    render();
    return;
  }

  if (mode === "multi-drag") {
    if (!state.interaction.historyCaptured && state.interaction.dragMoved) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    const startSnapped = snapPoint(state.interaction.startPoint);
    const dx = snappedCurrent.x - startSnapped.x;
    const dy = snappedCurrent.y - startSnapped.y;

    const orig = state.interaction.multiDragOrigItems;
    if (orig) {
      orig.streams.forEach(({ streamId, points }) => {
        const stream = getStreamById(streamId);
        if (stream) {
          stream.points = points.map((p) => ({ x: p.x + dx, y: p.y + dy }));
        }
      });
      orig.nodes.forEach(({ id, x, y }) => {
        const node = state.nodes.find((n) => n.id === id);
        if (node) {
          if (node.streamId) {
            const constrained = constrainActivityToStream({ x: x + dx, y: y + dy }, node.streamId);
            node.x = constrained.x;
            node.y = constrained.y;
            if (constrained.streamId) node.streamId = constrained.streamId;
          } else {
            node.x = x + dx;
            node.y = y + dy;
          }
        }
      });
      if (orig.textLabels) {
        orig.textLabels.forEach(({ id, x, y }) => {
          const tl = state.textLabels.find((t) => t.id === id);
          if (tl) {
            tl.x = x + dx;
            tl.y = y + dy;
          }
        });
      }
    }
    render();
    return;
  }

  if (mode === "handle-drag") {
    const stream = getStreamById(state.interaction.dragStreamId);
    if (!stream) {
      return;
    }
    if (!state.interaction.historyCaptured && state.interaction.dragMoved) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    let target = snappedCurrent;
    const snapTarget = getSnapTarget(rawCurrent, stream.id);
    if (snapTarget) {
      target = snapTarget;
    }

    // Shift: constrain to existing segment direction (only change length)
    if (event.shiftKey && state.interaction.shiftDir && state.interaction.shiftAnchor) {
      const dir = state.interaction.shiftDir;
      const anchor = state.interaction.shiftAnchor;
      const dx = target.x - anchor.x;
      const dy = target.y - anchor.y;
      const proj = dx * dir.x + dy * dir.y;
      target = snapPoint({ x: anchor.x + dir.x * proj, y: anchor.y + dir.y * proj });
    }

    stream.points[state.interaction.dragPointIndex] = target;

    const groupIds = getSelectedStreamGroup();
    groupIds.forEach((sid) => {
      state.nodes.filter((n) => n.streamId === sid).forEach((node) => {
        const constrained = nearestPointOnStream({ x: node.x, y: node.y }, getStreamById(sid));
        if (constrained) {
          node.x = constrained.x;
          node.y = constrained.y;
        }
      });
    });

    render();
    return;
  }

  if (mode === "connect-drag") {
    state.interaction.previewPoint = rawCurrent;
    const HIT_DIST = 24;
    let hoverNodeId = null;
    for (const node of state.nodes) {
      if (node.id === state.interaction.connectFromNodeId) continue;
      const dx = rawCurrent.x - node.x;
      const dy = rawCurrent.y - node.y;
      if (dx * dx + dy * dy <= HIT_DIST * HIT_DIST) {
        hoverNodeId = node.id;
        break;
      }
    }
    state.interaction.connectHoverNodeId = hoverNodeId;
    render();
    return;
  }

  if (mode === "guide-drag") {
    const guide = state.guides.find((g) => g.id === state.interaction.dragGuideId);
    if (!guide) return;
    if (!state.interaction.historyCaptured && state.interaction.dragMoved) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    guide.position = guide.axis === "v" ? snappedCurrent.x : snappedCurrent.y;
    render();
    return;
  }

  if (mode === "stream-drag") {
    if (!state.interaction.historyCaptured && state.interaction.dragMoved) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    const startSnapped = snapPoint(state.interaction.startPoint);
    const dx = snappedCurrent.x - startSnapped.x;
    const dy = snappedCurrent.y - startSnapped.y;

    const origEntries = state.interaction.dragStreamOrigPoints;
    origEntries.forEach(({ streamId, points: orig }) => {
      const stream = getStreamById(streamId);
      if (stream) {
        stream.points = orig.map((p) => ({ x: p.x + dx, y: p.y + dy }));
      }
    });

    const origNodes = state.interaction.dragStreamOrigNodes;
    if (origNodes) {
      origNodes.forEach(({ id, x, y }) => {
        const node = state.nodes.find((n) => n.id === id);
        if (node) {
          node.x = x + dx;
          node.y = y + dy;
        }
      });
    }
    render();
    return;
  }

  if (mode === "node-drag") {
    const node = state.nodes.find((n) => n.id === state.interaction.dragNodeId);
    if (!node) {
      return;
    }
    if (!state.interaction.historyCaptured && state.interaction.dragMoved) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    if (node.streamId) {
      const constrained = constrainActivityToStream(rawCurrent, node.streamId);
      node.x = constrained.x;
      node.y = constrained.y;
      if (constrained.streamId) {
        node.streamId = constrained.streamId;
      }
    } else {
      node.x = snappedCurrent.x;
      node.y = snappedCurrent.y;
    }
    render();
    return;
  }

  if (mode === "text-label-drag") {
    const tl = state.textLabels.find((t) => t.id === state.interaction.dragTextLabelId);
    if (!tl) return;
    if (!state.interaction.historyCaptured && state.interaction.dragMoved) {
      captureHistorySnapshot();
      state.interaction.historyCaptured = true;
    }
    tl.x = snappedCurrent.x;
    tl.y = snappedCurrent.y;
    render();
    return;
  }

  if (mode === "stream-create" || mode === "node-create") {
    if (mode === "node-create") {
      const targetStreamId = state.interaction.nodeCreateStreamId || state.streams[0]?.id;
      state.interaction.previewPoint = targetStreamId
        ? constrainActivityToStream(rawCurrent, targetStreamId)
        : snappedCurrent;
    } else {
      // stream-create: check for line snap onto same-color streams
      const presetStreamId = streamPresetSelect.value;
      const isExisting = presetStreamId !== "__new__";
      const targetColor = isExisting
        ? (getStreamById(presetStreamId)?.color || selectedStreamColor)
        : selectedStreamColor;
      const lineSnap = getLineSnapTarget(rawCurrent, targetColor);
      if (lineSnap) {
        state.interaction.previewPoint = { x: lineSnap.x, y: lineSnap.y };
        state.interaction.snapIndicator = lineSnap;
      } else {
        state.interaction.previewPoint = snappedCurrent;
        state.interaction.snapIndicator = null;
      }
    }
    render();
  }
}

let lastPointerEvent = null;

function showFloatingError(message) {
  const existing = document.querySelector('.floating-error');
  if (existing) existing.remove();

  const el = document.createElement('div');
  el.className = 'floating-error';
  el.textContent = message;
  document.body.appendChild(el);

  const px = lastPointerEvent ? lastPointerEvent.clientX : window.innerWidth / 2;
  const py = lastPointerEvent ? lastPointerEvent.clientY : window.innerHeight / 2;
  el.style.left = px + 'px';
  el.style.top = (py - 40) + 'px';

  requestAnimationFrame(() => el.classList.add('show'));
  setTimeout(() => {
    el.classList.remove('show');
    el.classList.add('hide');
    el.addEventListener('transitionend', () => el.remove());
  }, 2000);
}

function resetInteractionState() {
  state.pan.active = false;
  state.interaction.mode = null;
  state.interaction.historyCaptured = false;
  state.interaction.dragNodeId = null;
  state.interaction.nodeCreateStreamId = null;
  state.interaction.startPoint = null;
  state.interaction.previewPoint = null;
  state.interaction.dragMoved = false;
  state.interaction.dragStreamId = null;
  state.interaction.dragPointIndex = null;
  state.interaction.dragStreamOrigPoints = null;
  state.interaction.dragStreamOrigNodes = null;
  state.interaction.snapIndicator = null;
  state.interaction.dragGuideId = null;
  state.interaction.connectFromNodeId = null;
  state.interaction.connectHoverNodeId = null;
  state.interaction.boxSelectStart = null;
  state.interaction.boxSelectEnd = null;
  state.interaction.multiDragOrigItems = null;
  state.interaction.dragTextLabelId = null;
  state.interaction.calendarDragKey = null;
  state.interaction.calendarDragStartX = null;
  state.interaction.calendarDragStartWidth = null;
  state.interaction.shiftDir = null;
  state.interaction.shiftAnchor = null;
}

function abortInteraction(errorMsg) {
  resetInteractionState();
  render();
  if (errorMsg) showFloatingError(errorMsg);
}

function finishInteraction() {
  const mode = state.interaction.mode;
  if (!mode) {
    return;
  }

  if (mode === "stream-create") {
    const start = state.interaction.startPoint;
    const end = state.interaction.previewPoint;
    if (start && end && (start.x !== end.x || start.y !== end.y)) {
      const isExisting = streamPresetSelect.value !== "__new__";
      let name, color;

      if (isExisting) {
        const existing = getStreamById(streamPresetSelect.value);
        name = existing ? existing.name : "New Stream";
        color = existing ? existing.color : selectedStreamColor;
      } else {
        name = newStreamNameInput.value.trim();
        color = selectedStreamColor;

        if (!name) {
          abortInteraction("Please enter a stream name.");
          return;
        }
        if (isStreamNameTaken(name)) {
          abortInteraction('Stream name "' + name + '" already exists.');
          return;
        }
        if (isStreamColorTaken(color)) {
          abortInteraction("This color is already used by another stream.");
          return;
        }
      }

      const stream = createStream({ name, color, points: [start, end] });
      syncStreamPresetDropdown();
      streamPresetSelect.value = stream.id;
      applyStreamPreset();
    }
  }

  if (mode === "node-create") {
    const targetStreamId = state.interaction.nodeCreateStreamId || state.streams[0]?.id;
    const point = state.interaction.previewPoint;
    if (targetStreamId && point) {
      const type = state.tool === "add-deliverable" ? "deliverable" : "activity";
      const nodePoint = constrainActivityToStream(point, targetStreamId);
      const node = createNode({
        streamId: nodePoint.streamId || targetStreamId,
        type,
        x: nodePoint.x,
        y: nodePoint.y,
        label: type === "deliverable" ? "Deliverable" : "Activity",
        status: "planned",
        date: "",
      });
      selectItem("node", node.id);
      setTool("select");
    }
  }

  if (mode === "connect-drag") {
    const fromId = state.interaction.connectFromNodeId;
    const toId = state.interaction.connectHoverNodeId;
    if (fromId && toId && fromId !== toId) {
      createConnection({ fromNodeId: fromId, toNodeId: toId });
    }
  }

  if (mode === "box-select") {
    const bs = state.interaction.boxSelectStart;
    const be = state.interaction.boxSelectEnd;
    if (bs && be && state.interaction.dragMoved) {
      const minX = Math.min(bs.x, be.x);
      const maxX = Math.max(bs.x, be.x);
      const minY = Math.min(bs.y, be.y);
      const maxY = Math.max(bs.y, be.y);

      const inRect = (px, py) => px >= minX && px <= maxX && py >= minY && py <= maxY;

      // Select individual stream segments fully inside rect
      state.streams.forEach((stream) => {
        if (stream.points.every((p) => inRect(p.x, p.y))) {
          if (!isItemSelected("stream", stream.id)) {
            state.selected.push({ type: "stream", id: stream.id });
          }
        }
      });

      // Select nodes within rect
      state.nodes.forEach((node) => {
        if (inRect(node.x, node.y)) {
          if (!isItemSelected("node", node.id)) {
            state.selected.push({ type: "node", id: node.id });
          }
        }
      });

      // Select connections where both endpoint nodes are in rect
      state.connections.forEach((conn) => {
        const a = state.nodes.find((n) => n.id === conn.fromNodeId);
        const b = state.nodes.find((n) => n.id === conn.toNodeId);
        if (a && inRect(a.x, a.y) && b && inRect(b.x, b.y)) {
          if (!isItemSelected("connection", conn.id)) {
            state.selected.push({ type: "connection", id: conn.id });
          }
        }
      });

      // Select text labels within rect
      state.textLabels.forEach((tl) => {
        if (inRect(tl.x, tl.y)) {
          if (!isItemSelected("text-label", tl.id)) {
            state.selected.push({ type: "text-label", id: tl.id });
          }
        }
      });
    }
  }

  if (mode === "node-drag" && state.interaction.dragNodeId && !state.interaction.dragMoved) {
    selectItem("node", state.interaction.dragNodeId);
  }

  if (mode === "text-label-drag" && state.interaction.dragTextLabelId && !state.interaction.dragMoved) {
    selectItem("text-label", state.interaction.dragTextLabelId);
  }

  resetInteractionState();
  render();
}

function arrangeElement(type, id, action) {
  let arr;
  if (type === "stream") arr = state.streams;
  else if (type === "connection") arr = state.connections;
  else return;

  const idx = arr.findIndex((el) => el.id === id);
  if (idx === -1) return;

  captureHistorySnapshot();

  if (action === "bring-to-front") {
    const item = arr.splice(idx, 1)[0];
    arr.push(item);
  } else if (action === "send-to-back") {
    const item = arr.splice(idx, 1)[0];
    arr.unshift(item);
  } else if (action === "bring-forward") {
    if (idx < arr.length - 1) {
      [arr[idx], arr[idx + 1]] = [arr[idx + 1], arr[idx]];
    }
  } else if (action === "send-backward") {
    if (idx > 0) {
      [arr[idx], arr[idx - 1]] = [arr[idx - 1], arr[idx]];
    }
  }

  render();
}

function removeSelected() {
  if (state.selected.length === 0) return;

  captureHistorySnapshot();

  if (isMultiSelection()) {
    // Multi-selection: delete each item individually (streams as single segments, not grouped)
    const streamIdsToRemove = new Set();
    const nodeIdsToRemove = new Set();
    const connIdsToRemove = new Set();
    const textLabelIdsToRemove = new Set();

    state.selected.forEach((sel) => {
      if (sel.type === "stream") streamIdsToRemove.add(sel.id);
      else if (sel.type === "node") nodeIdsToRemove.add(sel.id);
      else if (sel.type === "connection") connIdsToRemove.add(sel.id);
      else if (sel.type === "text-label") textLabelIdsToRemove.add(sel.id);
    });

    // Also collect nodes that belong to removed streams
    streamIdsToRemove.forEach((sid) => {
      state.nodes.filter((n) => n.streamId === sid).forEach((n) => nodeIdsToRemove.add(n.id));
    });

    // Also remove connections attached to removed nodes
    state.connections.forEach((c) => {
      if (nodeIdsToRemove.has(c.fromNodeId) || nodeIdsToRemove.has(c.toNodeId)) {
        connIdsToRemove.add(c.id);
      }
    });

    state.streams = state.streams.filter((s) => !streamIdsToRemove.has(s.id));
    state.nodes = state.nodes.filter((n) => !nodeIdsToRemove.has(n.id));
    state.connections = state.connections.filter((c) => !connIdsToRemove.has(c.id));
    state.textLabels = state.textLabels.filter((t) => !textLabelIdsToRemove.has(t.id));
  } else {
    const selected = getSelectedEntity();
    if (!selected) return;

    if (selected.type === "node") {
      state.nodes = state.nodes.filter((n) => n.id !== selected.entity.id);
      state.connections = state.connections.filter(
        (c) => c.fromNodeId !== selected.entity.id && c.toNodeId !== selected.entity.id,
      );
    }

    if (selected.type === "stream") {
      const idsToRemove = state.selectionGrouped
        ? getConnectedSameColorStreamIds(selected.entity.id)
        : new Set([selected.entity.id]);
      const allNodeIds = new Set();
      idsToRemove.forEach((sid) => {
        state.nodes.filter((n) => n.streamId === sid).forEach((n) => allNodeIds.add(n.id));
      });
      state.streams = state.streams.filter((s) => !idsToRemove.has(s.id));
      state.nodes = state.nodes.filter((n) => !allNodeIds.has(n.id));
      state.connections = state.connections.filter(
        (c) => !allNodeIds.has(c.fromNodeId) && !allNodeIds.has(c.toNodeId),
      );
    }

    if (selected.type === "connection") {
      state.connections = state.connections.filter((c) => c.id !== selected.entity.id);
    }

    if (selected.type === "text-label") {
      state.textLabels = state.textLabels.filter((t) => t.id !== selected.entity.id);
    }
  }

  clearSelection();
  syncStreamPresetDropdown();
  applyStreamPreset();
  render();
}

function wirePanAndZoom() {
  svg.addEventListener("mousedown", handlePointerDown);
  window.addEventListener("mousemove", handlePointerMove);
  window.addEventListener("mousemove", trackPointer);
  window.addEventListener("mouseup", finishInteraction);

  svg.addEventListener(
    "wheel",
    (event) => {
      event.preventDefault();
      const zoomFactor = event.deltaY < 0 ? 1.08 : 1 / 1.08;
      const oldScale = state.transform.scale;
      const nextScale = clamp(oldScale * zoomFactor, 0.005, 5);
      // Zoom relative to the center of the visible canvas (viewBox center)
      const cx = 800;
      const cy = 500;
      // The SVG point under the viewBox center: svgX = (cx - tx) / scale
      const svgCenterX = (cx - state.transform.x) / oldScale;
      const svgCenterY = (cy - state.transform.y) / oldScale;
      state.transform.scale = nextScale;
      state.transform.x = cx - svgCenterX * nextScale;
      state.transform.y = cy - svgCenterY * nextScale;
      applyViewportTransform();
      renderCalendar();
    },
    { passive: false },
  );
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function goToToday() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const day = now.getDate();
  const daysInMonth = new Date(year, month, 0).getDate();
  const monthX = calMonthSvgX(year, month);
  const w = calMonthWidth(calMonthKey(year, month));
  const todayX = monthX + w * (day / daysInMonth);
  state.transform.x = 800 - todayX * state.transform.scale;
  applyViewportTransform();
  renderCalendar();
}

function fitToView() {
  if (state.streams.length === 0 && state.nodes.length === 0 && state.textLabels.length === 0) {
    state.transform = { x: 0, y: 0, scale: 1 };
    applyViewportTransform();
    renderCalendar();
    return;
  }

  const points = [];
  state.streams.forEach((s) => s.points.forEach((p) => points.push(p)));
  state.nodes.forEach((n) => points.push({ x: n.x, y: n.y }));
  state.textLabels.forEach((tl) => points.push({ x: tl.x, y: tl.y }));

  const minX = Math.min(...points.map((p) => p.x));
  const minY = Math.min(...points.map((p) => p.y));
  const maxX = Math.max(...points.map((p) => p.x));
  const maxY = Math.max(...points.map((p) => p.y));

  const mapWidth = maxX - minX + 160;
  const mapHeight = maxY - minY + 160;
  const scaleX = 1600 / mapWidth;
  const scaleY = 1000 / mapHeight;
  const scale = clamp(Math.min(scaleX, scaleY), 0.005, 2.2);

  state.transform.scale = scale;
  state.transform.x = 800 - ((minX + maxX) / 2) * scale;
  state.transform.y = 500 - ((minY + maxY) / 2) * scale;
  applyViewportTransform();
  renderCalendar();
}

function saveMap() {
  const data = {
    version: 1,
    streams: cloneData(state.streams),
    streamOrder: cloneData(state.streamOrder),
    nodes: cloneData(state.nodes),
    connections: cloneData(state.connections),
    guides: cloneData(state.guides),
    textLabels: cloneData(state.textLabels),
    calendar: cloneData(state.calendar),
    labelFontSize: labelFontSize,
  };
  downloadBlob(
    new Blob([JSON.stringify(data, null, 2)], { type: "application/json" }),
    "metro-map.json",
  );
}

function loadMap(file) {
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      if (!data || !Array.isArray(data.streams) || !Array.isArray(data.nodes)) {
        alert("Invalid map file.");
        return;
      }
      captureHistorySnapshot();
      state.streams = data.streams;
      state.streamOrder = Array.isArray(data.streamOrder) ? data.streamOrder : [];
      state.nodes = data.nodes;
      state.connections = Array.isArray(data.connections) ? data.connections : [];
      state.guides = Array.isArray(data.guides) ? data.guides : [];
      state.textLabels = Array.isArray(data.textLabels) ? data.textLabels : [];
      if (data.calendar && typeof data.calendar === "object") {
        state.calendar.monthWidths = data.calendar.monthWidths || {};
        if (typeof data.calendar.refYear === "number") state.calendar.refYear = data.calendar.refYear;
      }
      labelFontSize = typeof data.labelFontSize === "number" ? data.labelFontSize : 14;
      document.getElementById("labelFontSize").value = labelFontSize;
      document.getElementById("labelFontSizeValue").textContent = labelFontSize + "px";
      clearSelection();
      render();
      syncStreamPresetDropdown();
      fitToView();
    } catch {
      alert("Could not read map file.");
    }
  };
  reader.readAsText(file);
}

function getContentBounds() {
  const points = [];
  state.streams.forEach((s) => s.points.forEach((p) => points.push(p)));
  state.nodes.forEach((n) => {
    points.push({ x: n.x, y: n.y });
    // Account for label width
    const mw = n.labelMaxWidth || DEFAULT_LABEL_MAX_WIDTH;
    points.push({ x: n.x + mw / 2, y: n.y });
    points.push({ x: n.x - mw / 2, y: n.y });
    // Account for right-side description text
    const desc = n.labelDescription || "";
    const subs = (n.subLabels || []);
    const hasRightText = desc || subs.some((s) => (typeof s === "string" ? "" : (s.rightText || "")));
    if (hasRightText) {
      // Estimate description text width: ~7px per character at default font size
      const allTexts = [desc, ...subs.map((s) => typeof s === "string" ? "" : (s.rightText || ""))].filter(Boolean);
      const longestLen = Math.max(...allTexts.map((t) => t.length), 0);
      const estWidth = longestLen * (labelFontSize * 0.6);
      points.push({ x: n.x + mw / 2 + 10 + estWidth, y: n.y });
    }
    // Account for vertical extent of stacked labels below the node
    const labelCount = 1 + subs.length;
    const estLineH = labelFontSize * 1.4;
    const labelGap = 8;
    // Each label: text height (~estLineH) + 6px spacing + gap; starts 22px below node
    // For wrapped text, estimate ~2 lines per label as conservative bound
    const estLabelH = 22 + labelCount * (estLineH * 2 + 6 + labelGap);
    points.push({ x: n.x, y: n.y + estLabelH });
  });
  state.textLabels.forEach((tl) => {
    const mw = tl.maxWidth || 160;
    const fs = tl.fontSize || 16;
    points.push({ x: tl.x - mw / 2, y: tl.y });
    points.push({ x: tl.x + mw / 2, y: tl.y });
    // Estimate text label height: wrap text and add height
    const text = tl.text || "";
    const charsPerLine = Math.max(1, Math.floor(mw / (fs * 0.6)));
    const estLines = Math.max(1, Math.ceil(text.length / charsPerLine));
    points.push({ x: tl.x, y: tl.y + estLines * fs * 1.4 });
  });
  if (points.length === 0) return { x: 0, y: 0, w: 1600, h: 1000 };

  const minX = Math.min(...points.map((p) => p.x));
  const minY = Math.min(...points.map((p) => p.y));
  const maxX = Math.max(...points.map((p) => p.x));
  const maxY = Math.max(...points.map((p) => p.y));

  const pad = 100;
  return { x: minX - pad, y: minY - pad, w: maxX - minX + pad * 2, h: maxY - minY + pad * 2 };
}

function buildExportSvgString() {
  const ns = "http://www.w3.org/2000/svg";
  const clone = svg.cloneNode(true);

  // Remove UI-only elements
  clone.querySelectorAll(".stream-handle, .preview-line, .preview-node, .guide-line, .guide-hit, .connection-hit, .text-label-hit").forEach((el) => el.remove());
  clone.querySelector("#previewLayer")?.replaceChildren();
  clone.querySelector("#handlesLayer")?.replaceChildren();
  clone.querySelector("#guidesLayer")?.replaceChildren();
  clone.querySelector("#calendarGuidesLayer")?.replaceChildren();

  clone.querySelectorAll(".stream-line.selected").forEach((el) => el.classList.remove("selected"));
  clone.querySelectorAll(".node.selected").forEach((el) => el.classList.remove("selected"));
  clone.querySelectorAll(".connection.selected").forEach((el) => el.classList.remove("selected"));
  clone.querySelectorAll(".text-label-group.selected").forEach((el) => el.classList.remove("selected"));

  // Remove background and grid for transparent export
  clone.querySelectorAll(".canvas-bg, .canvas-grid").forEach((el) => el.remove());

  // Reset viewport transform so we export world coordinates
  const vp = clone.querySelector("#viewport");
  if (vp) vp.setAttribute("transform", "translate(0,0) scale(1)");

  // Convert foreignObject labels to native SVG text for PowerPoint compatibility
  clone.querySelectorAll("foreignObject.label-fo").forEach((fo) => {
    const div = fo.querySelector("div");
    if (!div) { fo.remove(); return; }

    const text = (div.textContent || "").trim();
    if (!text) { fo.remove(); return; }

    const foX = parseFloat(fo.getAttribute("x")) || 0;
    const foY = parseFloat(fo.getAttribute("y")) || 0;
    const foW = parseFloat(fo.getAttribute("width")) || 120;
    const foH = parseFloat(fo.getAttribute("height")) || 20;

    // Extract styling from inline style
    const style = div.style;
    const color = style.color || "#000000";
    const fontSize = parseFloat(style.fontSize) || 14;
    const textAlign = style.textAlign || "center";
    const maxWidth = parseFloat(style.maxWidth) || foW;
    const isDeliverable = div.classList.contains("label-deliverable");
    const isRightText = div.classList.contains("label-right-text");
    const fontWeight = isDeliverable ? "700" : "400";

    // Word-wrap into lines
    const effectiveW = isRightText ? 9999 : maxWidth;
    const lines = wrapTextToLines(text, effectiveW, fontSize, fontWeight);

    const textEl = document.createElementNS(ns, "text");
    textEl.setAttribute("fill", color);
    textEl.setAttribute("font-family", "'Open Sans', 'Segoe UI', sans-serif");
    textEl.setAttribute("font-size", fontSize);
    textEl.setAttribute("font-weight", fontWeight);

    // Compute x position based on alignment
    let anchorX;
    let textAnchor;
    if (textAlign === "left" || isRightText) {
      anchorX = foX;
      textAnchor = "start";
    } else if (textAlign === "right") {
      anchorX = foX + foW;
      textAnchor = "end";
    } else {
      anchorX = foX + foW / 2;
      textAnchor = "middle";
    }
    textEl.setAttribute("text-anchor", textAnchor);

    const lineH = fontSize * 1.3;
    lines.forEach((line, i) => {
      const tspan = document.createElementNS(ns, "tspan");
      tspan.setAttribute("x", anchorX);
      tspan.setAttribute("y", foY + fontSize + i * lineH);
      tspan.textContent = line;
      textEl.appendChild(tspan);
    });

    fo.parentNode.insertBefore(textEl, fo);
    fo.remove();
  });

  // Fit viewBox to content
  const bounds = getContentBounds();
  clone.setAttribute("viewBox", `${bounds.x} ${bounds.y} ${bounds.w} ${bounds.h}`);

  const styleEl = document.createElementNS(ns, "style");
  styleEl.textContent = `
    .stream-line { fill: none; stroke-width: 20; stroke-linecap: round; stroke-linejoin: round; }
    .connection { fill: none; stroke: rgba(0,0,0,0.4); stroke-width: 3; stroke-dasharray: 7 7; }
    .node { stroke: #53565A; stroke-width: 2; }
  `;

  const defs = clone.querySelector("defs");
  if (defs) {
    defs.appendChild(styleEl);
  } else {
    clone.insertBefore(styleEl, clone.firstChild);
  }

  return new XMLSerializer().serializeToString(clone);
}

// Simple word-wrap: split text into lines that fit within maxWidth
function wrapTextToLines(text, maxWidth, fontSize, fontWeight) {
  // Use a temporary SVG text element to measure
  const ns = "http://www.w3.org/2000/svg";
  const tmpSvg = document.createElementNS(ns, "svg");
  tmpSvg.style.cssText = "position:absolute;left:-9999px;top:-9999px;visibility:hidden;";
  document.body.appendChild(tmpSvg);

  const measure = document.createElementNS(ns, "text");
  measure.setAttribute("font-family", "'Open Sans', 'Segoe UI', sans-serif");
  measure.setAttribute("font-size", fontSize);
  measure.setAttribute("font-weight", fontWeight);
  tmpSvg.appendChild(measure);

  const words = text.split(/\s+/);
  const lines = [];
  let currentLine = "";

  words.forEach((word) => {
    const test = currentLine ? currentLine + " " + word : word;
    measure.textContent = test;
    const w = measure.getComputedTextLength();
    if (w > maxWidth && currentLine) {
      lines.push(currentLine);
      currentLine = word;
    } else {
      currentLine = test;
    }
  });
  if (currentLine) lines.push(currentLine);

  document.body.removeChild(tmpSvg);
  return lines.length ? lines : [text];
}

function exportSvg() {
  const svgString = buildExportSvgString();
  downloadBlob(
    new Blob([svgString], { type: "image/svg+xml;charset=utf-8" }),
    "metro-map.svg",
  );
}

function exportPng() {
  const svgString = buildExportSvgString();
  const svgBlob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(svgBlob);

  const bounds = getContentBounds();
  const MAX_DIM = 16384;
  const MAX_AREA = 268435456; // 256M pixels
  let scale = 2;

  // Clamp to browser canvas limits
  if (bounds.w * scale > MAX_DIM) scale = MAX_DIM / bounds.w;
  if (bounds.h * scale > MAX_DIM) scale = MAX_DIM / bounds.h;
  if (bounds.w * scale * bounds.h * scale > MAX_AREA) {
    scale = Math.sqrt(MAX_AREA / (bounds.w * bounds.h));
  }
  scale = Math.max(scale, 0.25); // never go below 0.25x

  const img = new Image();
  img.onload = () => {
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(bounds.w * scale);
    canvas.height = Math.ceil(bounds.h * scale);
    const ctx = canvas.getContext("2d");
    if (!ctx) {
      showFloatingError("Canvas too large to export as PNG. Try SVG export instead.");
      URL.revokeObjectURL(url);
      return;
    }
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

    canvas.toBlob((blob) => {
      if (blob) {
        downloadBlob(blob, "metro-map.png");
      } else {
        showFloatingError("PNG export failed. Try SVG export for very large maps.");
      }
      URL.revokeObjectURL(url);
    }, "image/png");
  };
  img.onerror = () => {
    showFloatingError("Failed to render SVG for PNG export.");
    URL.revokeObjectURL(url);
  };
  img.src = url;
}

function buildLegendSvgString() {
  const ns = "http://www.w3.org/2000/svg";

  const statusItems = [
    { label: "Planned Activity", color: "#ffffff", shape: "circle" },
    { label: "Completed", color: "#3dd86f", shape: "circle" },
    { label: "On-going", color: "#f0a13d", shape: "circle" },
    { label: "Needs Attention", color: "#c26700", shape: "circle" },
    { label: "Deliverable", color: "#f4d87a", shape: "diamond" },
  ];

  const streamGroups = getOrderedStreamGroups();
  const streamItems = streamGroups.map((g) => ({ label: g.name, color: g.color }));

  const lineH = 24;
  const pad = 20;
  const dotR = 7;
  const textX = pad + dotR * 2 + 12;
  const sectionGap = 16;
  const headerH = 26;
  const fontSize = 13;
  const headerFontSize = 11;

  let totalH = pad + headerH + statusItems.length * lineH;
  if (streamItems.length > 0) totalH += sectionGap + headerH + streamItems.length * lineH;
  totalH += pad;

  const textW = 180;
  const totalW = textX + textW + pad;

  const root = document.createElementNS(ns, "svg");
  root.setAttribute("xmlns", ns);
  root.setAttribute("viewBox", `0 0 ${totalW} ${totalH}`);
  root.setAttribute("width", totalW);
  root.setAttribute("height", totalH);

  const styleEl = document.createElementNS(ns, "style");
  styleEl.textContent = `@import url('https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;600;700&display=swap');`;
  root.appendChild(styleEl);

  let cy = pad;

  // Header: Stop Status
  const h1 = document.createElementNS(ns, "text");
  h1.setAttribute("x", pad);
  h1.setAttribute("y", cy);
  h1.setAttribute("fill", "#8fa6be");
  h1.setAttribute("font-family", "'Open Sans', sans-serif");
  h1.setAttribute("font-size", headerFontSize);
  h1.setAttribute("font-weight", "600");
  h1.setAttribute("dominant-baseline", "hanging");
  h1.textContent = "STOP STATUS";
  root.appendChild(h1);
  cy += headerH;

  statusItems.forEach((item) => {
    const iy = cy + lineH / 2;
    if (item.shape === "circle") {
      const c = document.createElementNS(ns, "circle");
      c.setAttribute("cx", pad + dotR);
      c.setAttribute("cy", iy);
      c.setAttribute("r", dotR);
      c.setAttribute("fill", item.color);
      c.setAttribute("stroke", "rgba(0,0,0,0.25)");
      c.setAttribute("stroke-width", "1");
      root.appendChild(c);
    } else {
      const d = document.createElementNS(ns, "rect");
      const s = dotR * 1.2;
      d.setAttribute("x", pad + dotR - s / 2);
      d.setAttribute("y", iy - s / 2);
      d.setAttribute("width", s);
      d.setAttribute("height", s);
      d.setAttribute("fill", item.color);
      d.setAttribute("stroke", "rgba(0,0,0,0.25)");
      d.setAttribute("stroke-width", "1");
      d.setAttribute("transform", `rotate(45 ${pad + dotR} ${iy})`);
      root.appendChild(d);
    }
    const t = document.createElementNS(ns, "text");
    t.setAttribute("x", textX);
    t.setAttribute("y", iy);
    t.setAttribute("fill", "#1a1a1a");
    t.setAttribute("font-family", "'Open Sans', sans-serif");
    t.setAttribute("font-size", fontSize);
    t.setAttribute("dominant-baseline", "central");
    t.textContent = item.label;
    root.appendChild(t);
    cy += lineH;
  });

  if (streamItems.length > 0) {
    cy += sectionGap;
    const h2 = document.createElementNS(ns, "text");
    h2.setAttribute("x", pad);
    h2.setAttribute("y", cy);
    h2.setAttribute("fill", "#8fa6be");
    h2.setAttribute("font-family", "'Open Sans', sans-serif");
    h2.setAttribute("font-size", headerFontSize);
    h2.setAttribute("font-weight", "600");
    h2.setAttribute("dominant-baseline", "hanging");
    h2.textContent = "STREAM COLORS";
    root.appendChild(h2);
    cy += headerH;

    streamItems.forEach((item) => {
      const iy = cy + lineH / 2;
      const c = document.createElementNS(ns, "circle");
      c.setAttribute("cx", pad + dotR);
      c.setAttribute("cy", iy);
      c.setAttribute("r", dotR);
      c.setAttribute("fill", item.color);
      c.setAttribute("stroke", "rgba(0,0,0,0.25)");
      c.setAttribute("stroke-width", "1");
      root.appendChild(c);

      const t = document.createElementNS(ns, "text");
      t.setAttribute("x", textX);
      t.setAttribute("y", iy);
      t.setAttribute("fill", "#1a1a1a");
      t.setAttribute("font-family", "'Open Sans', sans-serif");
      t.setAttribute("font-size", fontSize);
      t.setAttribute("dominant-baseline", "central");
      t.textContent = item.label;
      root.appendChild(t);
      cy += lineH;
    });
  }

  return new XMLSerializer().serializeToString(root);
}

function exportLegendSvg() {
  const svgString = buildLegendSvgString();
  downloadBlob(
    new Blob([svgString], { type: "image/svg+xml;charset=utf-8" }),
    "metro-legend.svg",
  );
}

function exportLegendPng() {
  const svgString = buildLegendSvgString();
  const svgBlob = new Blob([svgString], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(svgBlob);

  const parser = new DOMParser();
  const doc = parser.parseFromString(svgString, "image/svg+xml");
  const svgEl = doc.documentElement;
  const w = parseFloat(svgEl.getAttribute("width"));
  const h = parseFloat(svgEl.getAttribute("height"));
  const scale = 2;

  const img = new Image();
  img.onload = () => {
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(w * scale);
    canvas.height = Math.ceil(h * scale);
    const ctx = canvas.getContext("2d");
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    canvas.toBlob((blob) => {
      if (blob) downloadBlob(blob, "metro-legend.png");
      URL.revokeObjectURL(url);
    }, "image/png");
  };
  img.src = url;
}

function exportViewer() {
  const mapData = {
    streams: cloneData(state.streams),
    streamOrder: cloneData(state.streamOrder),
    nodes: cloneData(state.nodes),
    connections: cloneData(state.connections),
    textLabels: cloneData(state.textLabels),
    calendar: cloneData(state.calendar),
    labelFontSize: labelFontSize,
  };

  const orderedGroups = getOrderedStreamGroups();
  const legendStreams = orderedGroups.map((g) => ({ name: g.name, color: g.color }));

  const isDark = document.documentElement.getAttribute("data-theme") === "dark";

  const html = buildViewerHtml(mapData, legendStreams, isDark);
  downloadBlob(
    new Blob([html], { type: "text/html;charset=utf-8" }),
    "metro-map-viewer.html",
  );
}

function buildViewerHtml(mapData, legendStreams, isDark) {
  const dataJson = JSON.stringify(mapData).replace(/<\//g, "<\\/");
  const legendJson = JSON.stringify(legendStreams).replace(/<\//g, "<\\/");

  return `<!DOCTYPE html>
<html lang="en" data-theme="${isDark ? "dark" : "light"}">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Metro Map Viewer</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin/>
<link href="https://fonts.googleapis.com/css2?family=Open+Sans:wght@400;700&family=Space+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap" rel="stylesheet"/>
<style>
:root {
  --bg-0:#f0f2f5;--bg-1:#ffffff;--bg-2:#e8ecf1;
  --panel:rgba(255,255,255,0.92);--text-strong:#1a2332;--text-muted:#6b7a8d;
  --line-faint:rgba(0,0,0,0.08);--accent:#0078d4;
  --canvas-bg:#f8f9fb;--canvas-wrap-bg:rgba(248,249,251,0.85);
  --calendar-bg:rgba(255,255,255,0.92);--calendar-border:rgba(0,120,212,0.12);
  --calendar-text:rgba(44,62,80,0.50);--calendar-jan-text:rgba(44,62,80,0.75);
  --calendar-jan-bg:rgba(0,120,212,0.04);--calendar-current-bg:rgba(0,120,212,0.08);
  --calendar-current-text:rgba(0,120,212,0.70);--calendar-guide:rgba(0,120,212,0.06);
  --legend-bg:rgba(255,255,255,0.88);
  --tooltip-bg:rgba(255,255,255,0.96);--tooltip-border:#b0bac5;--tooltip-text:#2c3e50;
  --node-stroke:#53565A;--connection-stroke:rgba(0,0,0,0.32);
  --grid-stroke:rgba(0,0,0,0.06);
  --heading-color:#3a4d63;--planned:#ffffff;--completed:#92D050;
  --ongoing:#FFCD00;--attention:#ED8B00;
  --body-gradient:linear-gradient(160deg,#f0f2f5,#ffffff 44%,#f5f7fa 100%);
  --body-radial-1:radial-gradient(circle at 14% 9%,rgba(0,120,212,0.08),transparent 34%);
  --body-radial-2:radial-gradient(circle at 84% 86%,rgba(255,145,77,0.06),transparent 28%);
}
[data-theme="dark"] {
  --bg-0:#0a1018;--bg-1:#121b28;--bg-2:#192538;
  --panel:rgba(13,20,31,0.86);--text-strong:#e6eef8;--text-muted:#8ea1b5;
  --line-faint:rgba(255,255,255,0.08);--accent:#4fc2ff;
  --canvas-bg:#0d1725;--canvas-wrap-bg:rgba(10,16,24,0.75);
  --calendar-bg:rgba(13,23,37,0.88);--calendar-border:rgba(79,194,255,0.10);
  --calendar-text:rgba(200,214,230,0.40);--calendar-jan-text:rgba(200,214,230,0.65);
  --calendar-jan-bg:rgba(79,194,255,0.03);--calendar-current-bg:rgba(79,194,255,0.06);
  --calendar-current-text:rgba(79,194,255,0.55);--calendar-guide:rgba(79,194,255,0.035);
  --legend-bg:rgba(10,16,24,0.82);
  --tooltip-bg:rgba(15,24,37,0.94);--tooltip-border:#3e5b7f;--tooltip-text:#d5e5f7;
  --node-stroke:#53565A;--connection-stroke:rgba(255,255,255,0.52);
  --grid-stroke:rgba(255,255,255,0.06);
  --heading-color:#c7d6e7;
  --body-gradient:linear-gradient(160deg,#0a1018,#121b28 44%,#0f1724 100%);
  --body-radial-1:radial-gradient(circle at 14% 9%,rgba(79,194,255,0.23),transparent 34%);
  --body-radial-2:radial-gradient(circle at 84% 86%,rgba(255,145,77,0.2),transparent 28%);
}
*{box-sizing:border-box;margin:0;padding:0;}
html,body{width:100%;height:100%;overflow:hidden;
  font-family:"Space Grotesk","Segoe UI",sans-serif;
  background:var(--body-radial-1),var(--body-radial-2),var(--body-gradient);
  color:var(--text-strong);
}
.viewer-wrap{position:relative;width:100%;height:100%;}
svg{width:100%;height:100%;display:block;}
.canvas-bg{fill:var(--canvas-bg);}
.canvas-grid{pointer-events:none;}
.stream-line{fill:none;stroke-width:20;stroke-linecap:round;stroke-linejoin:round;}
.connection{fill:none;stroke:var(--connection-stroke);stroke-width:3;stroke-dasharray:7 7;pointer-events:none;}
.node{stroke:var(--node-stroke);stroke-width:2;cursor:pointer;}
.label-fo{overflow:visible;pointer-events:none;}
.label-wrap{font-family:'Open Sans','Segoe UI',sans-serif;font-weight:400;line-height:1.3;word-wrap:break-word;overflow-wrap:break-word;white-space:normal;}
.label-wrap.label-deliverable{font-weight:700;}
.label-right-text{white-space:nowrap;word-wrap:normal;overflow-wrap:normal;}
.text-label-content{font-family:'Open Sans','Segoe UI',sans-serif;font-weight:400;line-height:1.35;word-wrap:break-word;overflow-wrap:break-word;white-space:pre-wrap;}
.label-bg{rx:4;ry:4;}
.calendar-guide-line{stroke:var(--calendar-guide);stroke-width:0.5;pointer-events:none;}

/* Calendar bar */
.calendar-bar{position:absolute;top:0;left:0;right:0;height:36px;z-index:2;overflow:hidden;
  background:var(--calendar-bg);border-bottom:1px solid var(--calendar-border);
  backdrop-filter:blur(8px);pointer-events:none;user-select:none;}
.cal-month{position:absolute;top:0;height:100%;display:flex;align-items:center;justify-content:center;
  font-family:'Space Grotesk',sans-serif;font-size:0.72rem;font-weight:500;letter-spacing:0.03em;
  color:var(--calendar-text);border-right:1px solid var(--calendar-border);
  box-sizing:border-box;user-select:none;overflow:hidden;white-space:nowrap;pointer-events:none;}
.cal-month.cal-jan{color:var(--calendar-jan-text);font-weight:600;letter-spacing:0.06em;background:var(--calendar-jan-bg);}
.cal-month.cal-current{background:var(--calendar-current-bg);color:var(--calendar-current-text);}

/* Legend */
.legend-card{position:absolute;right:14px;bottom:14px;width:230px;border-radius:14px;
  border:1px solid var(--line-faint);background:var(--legend-bg);backdrop-filter:blur(8px);padding:12px;
  z-index:3;max-height:60vh;overflow-y:auto;}
.legend-card h3{margin:0 0 9px;font-size:0.95rem;}
.legend-title{margin:8px 0 4px;font-size:0.78rem;text-transform:uppercase;letter-spacing:0.07em;color:var(--text-muted);}
.legend-card ul{list-style:none;margin:0;padding:0;display:grid;gap:6px;}
.legend-card li{display:flex;align-items:center;gap:8px;font-size:0.78rem;color:var(--text-strong);}
.dot,.legend-stream-color{width:14px;height:14px;border-radius:50%;border:1px solid rgba(255,255,255,0.5);flex-shrink:0;}
.dot.planned{background:var(--planned);}
.dot.completed{background:var(--completed);}
.dot.ongoing{background:var(--ongoing);}
.dot.attention{background:var(--attention);}
.diamond{width:12px;height:12px;transform:rotate(45deg);border:2px solid #53565A;background:rgba(83,86,90,0.2);flex-shrink:0;}

/* Tooltip */
.node-tooltip{position:fixed;transform:translate(-50%,-100%);padding:6px 14px;
  background:var(--tooltip-bg);border:1px solid var(--tooltip-border);border-radius:8px;
  color:var(--tooltip-text);font-family:"Space Grotesk",sans-serif;font-size:12px;
  pointer-events:none;white-space:nowrap;z-index:9999;opacity:0;transition:opacity 0.15s ease;
  display:flex;align-items:center;gap:8px;}
.node-tooltip.visible{opacity:1;}
.node-tooltip-date{color:var(--accent);font-family:"JetBrains Mono",monospace;font-weight:600;}

/* Viewer toolbar */
.viewer-toolbar{position:absolute;top:46px;left:50%;transform:translateX(-50%);z-index:4;
  display:flex;gap:6px;padding:6px 10px;border-radius:12px;
  background:var(--legend-bg);border:1px solid var(--line-faint);backdrop-filter:blur(8px);}
.viewer-toolbar button{background:var(--legend-bg);color:var(--text-strong);border:1px solid var(--line-faint);
  border-radius:8px;padding:6px 14px;font-family:"Space Grotesk",sans-serif;font-size:0.8rem;font-weight:500;
  cursor:pointer;transition:background 0.15s,border-color 0.15s;}
.viewer-toolbar button:hover{background:var(--line-faint);border-color:var(--accent);}

/* Minimap */
.minimap{position:absolute;left:14px;bottom:14px;width:220px;height:140px;border-radius:12px;
  border:1px solid var(--line-faint);background:var(--legend-bg);backdrop-filter:blur(8px);
  z-index:4;overflow:hidden;cursor:pointer;box-shadow:0 2px 12px rgba(0,0,0,0.12);}
.minimap canvas{width:100%;height:100%;display:block;}
.minimap-viewport{position:absolute;border:2px solid var(--accent);border-radius:2px;
  background:rgba(0,120,212,0.08);pointer-events:none;transition:none;}
[data-theme="dark"] .minimap-viewport{background:rgba(79,194,255,0.10);}
</style>
</head>
<body>
<div class="viewer-wrap">
  <div class="calendar-bar" id="calendarBar"></div>
  <div class="viewer-toolbar">
    <button id="fitViewBtn">Fit to View</button>
    <button id="todayBtn">Today</button>
  </div>
  <svg id="metroCanvas" viewBox="0 0 1600 1000" xmlns="http://www.w3.org/2000/svg">
    <defs>
      <pattern id="gridPattern" width="20" height="20" patternUnits="userSpaceOnUse">
        <path d="M 20 0 L 0 0 0 20" fill="none" stroke="var(--grid-stroke)" stroke-width="1"/>
      </pattern>
    </defs>
    <rect class="canvas-bg" x="-5000" y="-5000" width="16000" height="16000"></rect>
    <rect class="canvas-grid" x="-5000" y="-5000" width="16000" height="16000" fill="url(#gridPattern)"></rect>
    <g id="viewport">
      <g id="calendarGuidesLayer"></g>
      <g id="connectionsLayer"></g>
      <g id="streamsLayer"></g>
      <g id="nodesLayer"></g>
      <g id="labelsLayer"></g>
    </g>
  </svg>

  <div class="minimap" id="minimap">
    <canvas id="minimapCanvas"></canvas>
    <div class="minimap-viewport" id="minimapViewport"></div>
  </div>

  <div class="legend-card">
    <h3>Legend</h3>
    <p class="legend-title">Stop Status</p>
    <ul>
      <li><span class="dot planned"></span> Planned Activity</li>
      <li><span class="dot completed"></span> Completed</li>
      <li><span class="dot ongoing"></span> On-going</li>
      <li><span class="dot attention"></span> Needs Attention</li>
      <li><span class="diamond"></span> Deliverable</li>
    </ul>
    <p class="legend-title">Stream Colors</p>
    <ul id="streamLegend"></ul>
  </div>
</div>

<script>
(function(){
"use strict";
var DATA = ${dataJson};
var LEGEND_STREAMS = ${legendJson};

var STATUS_COLORS = {planned:"#ffffff",completed:"#92D050",ongoing:"#FFCD00",attention:"#ED8B00"};
var DEFAULT_LABEL_MAX_WIDTH = 120;
var DEFAULT_MONTH_WIDTH = 1500;
var MONTH_NAMES = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

var svg = document.getElementById("metroCanvas");
var viewport = document.getElementById("viewport");
var calendarBar = document.getElementById("calendarBar");
var calendarGuidesLayer = document.getElementById("calendarGuidesLayer");
var streamsLayer = document.getElementById("streamsLayer");
var nodesLayer = document.getElementById("nodesLayer");
var labelsLayer = document.getElementById("labelsLayer");
var connectionsLayer = document.getElementById("connectionsLayer");
var streamLegendEl = document.getElementById("streamLegend");

var labelFontSize = DATA.labelFontSize || 14;
var calendar = DATA.calendar || {monthWidths:{},refYear:new Date().getFullYear()};

/* ── Transform state ──────────────────────────────── */
var tx = 0, ty = 0, scale = 1;

/* ── Smooth animation state ── */
var animTarget = {tx:0,ty:0,scale:1};
var animCurrent = {tx:0,ty:0,scale:1};
var animating = false;
var LERP_FACTOR = 0.18;
var SNAP_THRESHOLD = 0.0005;

function scheduleAnim(){
  if(!animating){animating=true;requestAnimationFrame(animLoop);}
}
function animLoop(){
  var dx = animTarget.tx - animCurrent.tx;
  var dy = animTarget.ty - animCurrent.ty;
  var ds = animTarget.scale - animCurrent.scale;
  animCurrent.tx += dx * LERP_FACTOR;
  animCurrent.ty += dy * LERP_FACTOR;
  animCurrent.scale += ds * LERP_FACTOR;

  var settled = Math.abs(dx)<0.5 && Math.abs(dy)<0.5 && Math.abs(ds/animCurrent.scale)<SNAP_THRESHOLD;
  if(settled){
    animCurrent.tx = animTarget.tx;
    animCurrent.ty = animTarget.ty;
    animCurrent.scale = animTarget.scale;
    animating = false;
  }
  tx = animCurrent.tx; ty = animCurrent.ty; scale = animCurrent.scale;
  applyTransform();
  renderCalendar();
  if(animating) requestAnimationFrame(animLoop);
}

function setTransformAnimated(ntx,nty,ns){
  animTarget.tx=ntx; animTarget.ty=nty; animTarget.scale=ns;
  scheduleAnim();
}
function setTransformImmediate(ntx,nty,ns){
  tx=ntx;ty=nty;scale=ns;
  animTarget.tx=ntx;animTarget.ty=nty;animTarget.scale=ns;
  animCurrent.tx=ntx;animCurrent.ty=nty;animCurrent.scale=ns;
  applyTransform();
  renderCalendar();
}

function applyTransform(){
  viewport.setAttribute("transform","translate("+tx+","+ty+") scale("+scale+")");
}

/* ── Calendar helpers ─────────────────────────────── */
function calMonthKey(y,m){return y+"-"+String(m).padStart(2,"0");}
function calMonthWidth(key){return calendar.monthWidths[key]||DEFAULT_MONTH_WIDTH;}
function calMonthFromIndex(idx){var y=Math.floor(idx/12);var m=(((idx%12)+12)%12)+1;return{year:y,month:m};}
function calMonthIndex(y,m){return y*12+(m-1);}
function calMonthSvgX(tY,tM){
  var refIdx=calMonthIndex(calendar.refYear,1);var targetIdx=calMonthIndex(tY,tM);var x=0;
  if(targetIdx>=refIdx){for(var i=refIdx;i<targetIdx;i++){var m=calMonthFromIndex(i);x+=calMonthWidth(calMonthKey(m.year,m.month));}}
  else{for(var i=refIdx-1;i>=targetIdx;i--){var m=calMonthFromIndex(i);x-=calMonthWidth(calMonthKey(m.year,m.month));}}
  return x;
}

function getSvgScreenMapping(){
  var cw=svg.clientWidth||1;var ch=svg.clientHeight||1;
  var svgScale=Math.min(cw/1600,ch/1000);var offsetX=(cw-1600*svgScale)/2;
  return{svgScale:svgScale,offsetX:offsetX};
}
function svgXToPixelX(svgX){
  var m=getSvgScreenMapping();var vbX=svgX*scale+tx;return vbX*m.svgScale+m.offsetX;
}
function pixelXToSvgX(pixelX){
  var m=getSvgScreenMapping();var vbX=(pixelX-m.offsetX)/m.svgScale;return(vbX-tx)/scale;
}

function renderCalendar(){
  calendarBar.innerHTML="";
  calendarGuidesLayer.innerHTML="";
  var barWidth=calendarBar.clientWidth||1;
  var leftSvgX=pixelXToSvgX(0);var rightSvgX=pixelXToSvgX(barWidth);
  var refIdx=calMonthIndex(calendar.refYear,1);
  var now=new Date();var curKey=calMonthKey(now.getFullYear(),now.getMonth()+1);
  var scanIdx=refIdx;var scanX=0;
  while(scanX>leftSvgX){scanIdx--;var mm=calMonthFromIndex(scanIdx);scanX-=calMonthWidth(calMonthKey(mm.year,mm.month));}
  scanIdx--;var m0=calMonthFromIndex(scanIdx);scanX-=calMonthWidth(calMonthKey(m0.year,m0.month));
  var frag=document.createDocumentFragment();
  var gFrag=document.createDocumentFragment();
  var EXTENT=10000;var idx=scanIdx;var x=scanX;
  while(x<rightSvgX+DEFAULT_MONTH_WIDTH){
    var mm=calMonthFromIndex(idx);var key=calMonthKey(mm.year,mm.month);var w=calMonthWidth(key);
    var pxL=svgXToPixelX(x);var pxR=svgXToPixelX(x+w);var pxW=pxR-pxL;
    if(pxR>-10&&pxL<barWidth+10){
      var box=document.createElement("div");
      box.className="cal-month"+(mm.month===1?" cal-jan":"")+(key===curKey?" cal-current":"");
      box.style.left=pxL+"px";box.style.width=pxW+"px";
      if(pxW>28) box.textContent=mm.month===1?MONTH_NAMES[0]+" "+mm.year:MONTH_NAMES[mm.month-1];
      frag.appendChild(box);
      var line=document.createElementNS("http://www.w3.org/2000/svg","line");
      line.setAttribute("x1",x+w);line.setAttribute("y1",-EXTENT);
      line.setAttribute("x2",x+w);line.setAttribute("y2",EXTENT);
      line.classList.add("calendar-guide-line");
      gFrag.appendChild(line);
    }
    x+=w;idx++;
  }
  calendarBar.appendChild(frag);
  calendarGuidesLayer.appendChild(gFrag);
}

/* ── Render map ────────────────────────────────────── */
function renderStreams(){
  streamsLayer.innerHTML="";
  DATA.streams.forEach(function(s){
    if(s.points.length<2)return;
    var pts=s.points.map(function(p){return p.x+","+p.y;}).join(" ");
    var pl=document.createElementNS("http://www.w3.org/2000/svg","polyline");
    pl.setAttribute("points",pts);pl.setAttribute("stroke",s.color);
    pl.classList.add("stream-line");streamsLayer.appendChild(pl);
  });
}

function renderConnections(){
  connectionsLayer.innerHTML="";
  DATA.connections.forEach(function(conn){
    var a=DATA.nodes.find(function(n){return n.id===conn.fromNodeId;});
    var b=DATA.nodes.find(function(n){return n.id===conn.toNodeId;});
    if(!a||!b)return;
    var c1x=a.x+(b.x-a.x)*0.35;var c2x=a.x+(b.x-a.x)*0.65;
    var d="M "+a.x+" "+a.y+" C "+c1x+" "+a.y+", "+c2x+" "+b.y+", "+b.x+" "+b.y;
    var path=document.createElementNS("http://www.w3.org/2000/svg","path");
    path.classList.add("connection");path.setAttribute("d",d);
    connectionsLayer.appendChild(path);
  });
}

function renderNodes(){
  nodesLayer.innerHTML="";
  DATA.nodes.forEach(function(node){
    var g=document.createElementNS("http://www.w3.org/2000/svg","g");
    var fill=STATUS_COLORS[node.status]||STATUS_COLORS.planned;
    if(node.type==="activity"){
      var c=document.createElementNS("http://www.w3.org/2000/svg","circle");
      c.setAttribute("cx",node.x);c.setAttribute("cy",node.y);c.setAttribute("r",12);
      c.setAttribute("fill",fill);c.classList.add("node");g.appendChild(c);
    }else{
      var d=document.createElementNS("http://www.w3.org/2000/svg","rect");
      d.setAttribute("x",node.x-11);d.setAttribute("y",node.y-11);
      d.setAttribute("width",22);d.setAttribute("height",22);
      d.setAttribute("fill",fill);
      d.setAttribute("transform","rotate(45 "+node.x+" "+node.y+")");
      d.setAttribute("stroke","#53565A");d.setAttribute("stroke-width","3");
      d.classList.add("node");g.appendChild(d);
    }
    g.addEventListener("mouseenter",function(e){showTooltip(e,node);});
    g.addEventListener("mousemove",function(e){moveTooltip(e);});
    g.addEventListener("mouseleave",hideTooltip);
    nodesLayer.appendChild(g);
  });
}

function getNodeLabelStyle(node){
  return{
    textColor:node.labelTextColor||(node.type==="deliverable"?"#ffffff":"#000000"),
    bgColor:node.labelBgColor===undefined?(node.type==="deliverable"?"#000000":null):node.labelBgColor
  };
}

function renderLabels(){
  labelsLayer.innerHTML="";
  var fs=labelFontSize;
  DATA.nodes.forEach(function(node){
    var mainStyle=getNodeLabelStyle(node);
    var mw=node.labelMaxWidth||DEFAULT_LABEL_MAX_WIDTH;
    var rawSubs=node.subLabels||[];
    var subObjs=rawSubs.map(function(s){return typeof s==="string"?{text:s,rightText:""}:s;});
    var isDeliv=node.type==="deliverable";
    var pos=node.labelPosition||(isDeliv?"down":"down");
    var allLabels=[{text:node.label,rightText:node.labelDescription||"",textColor:mainStyle.textColor,bgColor:mainStyle.bgColor,descPosition:node.labelDescPosition||"right"}].concat(subObjs.map(function(s){
      return{text:s.text,rightText:s.rightText||"",textColor:s.textColor||mainStyle.textColor,bgColor:s.bgColor!==undefined?s.bgColor:mainStyle.bgColor,descPosition:s.descPosition||"right"};
    }));
    var labelGap=8;var entries=[];
    if(pos==="up"){
      var measured=[];
      allLabels.forEach(function(entry){
        var fo=document.createElementNS("http://www.w3.org/2000/svg","foreignObject");
        fo.setAttribute("x",node.x-mw/2);fo.setAttribute("y",0);fo.setAttribute("width",mw);fo.setAttribute("height",200);fo.setAttribute("class","label-fo");
        var div=document.createElement("div");div.className="label-wrap"+(isDeliv?" label-deliverable":"");
        div.style.cssText="color:"+entry.textColor+";font-size:"+fs+"px;max-width:"+mw+"px;text-align:center;";
        div.textContent=entry.text;fo.appendChild(div);labelsLayer.appendChild(fo);
        var actualH=div.offsetHeight||fs*1.4;fo.setAttribute("height",actualH+4);
        measured.push({fo:fo,h:actualH,entry:entry});
      });
      var bottomY=node.y-22;
      for(var i=measured.length-1;i>=0;i--){
        var m=measured[i];var y=bottomY-m.h;m.fo.setAttribute("y",y);
        entries.push({fo:m.fo,y:y,h:m.h,rightText:m.entry.rightText,bgColor:m.entry.bgColor,descPosition:m.entry.descPosition});
        bottomY=y-6-labelGap;
      }
    }else if(pos==="left"||pos==="right"){
      var tAlign=pos==="left"?"right":"left";
      var fo=document.createElementNS("http://www.w3.org/2000/svg","foreignObject");
      var aX=pos==="left"?node.x-22-mw:node.x+22;
      fo.setAttribute("x",aX);fo.setAttribute("y",0);fo.setAttribute("width",mw);fo.setAttribute("height",200);fo.setAttribute("class","label-fo");
      var div=document.createElement("div");div.className="label-wrap"+(isDeliv?" label-deliverable":"");
      div.style.cssText="color:"+allLabels[0].textColor+";font-size:"+fs+"px;max-width:"+mw+"px;text-align:"+tAlign+";";
      div.textContent=allLabels[0].text;fo.appendChild(div);labelsLayer.appendChild(fo);
      var actualH=div.offsetHeight||fs*1.4;fo.setAttribute("height",actualH+4);
      var centeredY=node.y-actualH/2;
      fo.setAttribute("y",centeredY);
      entries.push({fo:fo,y:centeredY,h:actualH,rightText:allLabels[0].rightText,bgColor:allLabels[0].bgColor,descPosition:allLabels[0].descPosition});
      if(isDeliv&&allLabels.length>1){
        var stackY=node.y+22;
        for(var i=1;i<allLabels.length;i++){
          var entry=allLabels[i];
          var sfo=document.createElementNS("http://www.w3.org/2000/svg","foreignObject");
          sfo.setAttribute("x",node.x-mw/2);sfo.setAttribute("y",stackY);sfo.setAttribute("width",mw);sfo.setAttribute("height",200);sfo.setAttribute("class","label-fo");
          var sdiv=document.createElement("div");sdiv.className="label-wrap label-deliverable";
          sdiv.style.cssText="color:"+entry.textColor+";font-size:"+fs+"px;max-width:"+mw+"px;text-align:center;";
          sdiv.textContent=entry.text;sfo.appendChild(sdiv);labelsLayer.appendChild(sfo);
          var sH=sdiv.offsetHeight||fs*1.4;sfo.setAttribute("height",sH+4);
          entries.push({fo:sfo,y:stackY,h:sH,rightText:entry.rightText,bgColor:entry.bgColor,descPosition:entry.descPosition});
          stackY+=sH+6+labelGap;
        }
      }
    }else{
      var curY=node.y+22;
      allLabels.forEach(function(entry){
        var fo=document.createElementNS("http://www.w3.org/2000/svg","foreignObject");
        fo.setAttribute("x",node.x-mw/2);fo.setAttribute("y",curY);fo.setAttribute("width",mw);fo.setAttribute("height",200);fo.setAttribute("class","label-fo");
        var div=document.createElement("div");div.className="label-wrap"+(isDeliv?" label-deliverable":"");
        div.style.cssText="color:"+entry.textColor+";font-size:"+fs+"px;max-width:"+mw+"px;text-align:center;";
        div.textContent=entry.text;fo.appendChild(div);labelsLayer.appendChild(fo);
        var actualH=div.offsetHeight||fs*1.4;fo.setAttribute("height",actualH+4);
        entries.push({fo:fo,y:curY,h:actualH,rightText:entry.rightText,bgColor:entry.bgColor,descPosition:entry.descPosition});
        curY+=actualH+6+labelGap;
      });
    }
    entries.forEach(function(entry){
      if(!entry.bgColor)return;
      var foX=parseFloat(entry.fo.getAttribute("x"));var foW=parseFloat(entry.fo.getAttribute("width"));
      var rect=document.createElementNS("http://www.w3.org/2000/svg","rect");
      rect.classList.add("label-bg");rect.setAttribute("x",foX-6);rect.setAttribute("y",entry.y-3);
      rect.setAttribute("width",foW+12);rect.setAttribute("height",entry.h+10);rect.setAttribute("fill",entry.bgColor);
      labelsLayer.insertBefore(rect,entry.fo);
    });
    entries.forEach(function(entry){
      if(!entry.rightText)return;
      var dp=entry.descPosition||"right";
      var rfo=document.createElementNS("http://www.w3.org/2000/svg","foreignObject");
      rfo.setAttribute("class","label-fo");
      var rdiv=document.createElement("div");rdiv.className="label-wrap label-deliverable label-right-text";
      var foX=parseFloat(entry.fo.getAttribute("x"));var foW=parseFloat(entry.fo.getAttribute("width"));
      if(dp==="left"){
        rfo.setAttribute("x",foX-2010);rfo.setAttribute("y",entry.y);rfo.setAttribute("width",2000);rfo.setAttribute("height",entry.h+4);
        rdiv.style.cssText="color:#000000;font-size:"+fs+"px;text-align:right;";
      }else{
        rfo.setAttribute("x",foX+foW+10);rfo.setAttribute("y",entry.y);rfo.setAttribute("width",2000);rfo.setAttribute("height",entry.h+4);
        rdiv.style.cssText="color:#000000;font-size:"+fs+"px;text-align:left;";
      }
      rdiv.textContent=entry.rightText;rfo.appendChild(rdiv);labelsLayer.appendChild(rfo);
    });
  });
  DATA.textLabels.forEach(function(tl){
    var mw=tl.maxWidth||160;var fs2=tl.fontSize||16;var align=tl.align||"center";
    var g=document.createElementNS("http://www.w3.org/2000/svg","g");
    var fo=document.createElementNS("http://www.w3.org/2000/svg","foreignObject");
    fo.setAttribute("x",tl.x-mw/2);fo.setAttribute("y",tl.y);
    fo.setAttribute("width",mw);fo.setAttribute("height",400);fo.setAttribute("class","label-fo");
    var div=document.createElement("div");
    div.className="text-label-content";
    div.style.cssText="color:"+(tl.textColor||"#ffffff")+";font-size:"+fs2+"px;max-width:"+mw+"px;text-align:"+align+";";
    div.textContent=tl.text;fo.appendChild(div);g.appendChild(fo);labelsLayer.appendChild(g);
    var actualH=div.offsetHeight||fs2*1.4;fo.setAttribute("height",actualH+4);
    if(tl.bgColor){
      var rect=document.createElementNS("http://www.w3.org/2000/svg","rect");
      rect.classList.add("label-bg");
      rect.setAttribute("x",tl.x-mw/2-6);rect.setAttribute("y",tl.y-3);
      rect.setAttribute("width",mw+12);rect.setAttribute("height",actualH+10);
      rect.setAttribute("fill",tl.bgColor);g.insertBefore(rect,fo);
    }
  });
}

function renderLegend(){
  streamLegendEl.innerHTML="";
  LEGEND_STREAMS.forEach(function(entry){
    var li=document.createElement("li");
    var swatch=document.createElement("span");
    swatch.className="legend-stream-color";
    swatch.style.background=entry.color;
    li.appendChild(swatch);
    li.appendChild(document.createTextNode(entry.name));
    streamLegendEl.appendChild(li);
  });
}

/* ── Tooltip ───────────────────────────────────────── */
var tooltipEl=null;
function ensureTooltip(){
  if(!tooltipEl){tooltipEl=document.createElement("div");tooltipEl.className="node-tooltip";document.body.appendChild(tooltipEl);}
  return tooltipEl;
}
function formatDatePretty(dateStr){
  if(!dateStr)return null;
  var d=new Date(dateStr+"T00:00:00");if(isNaN(d))return dateStr;
  return d.toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"});
}
function showTooltip(e,node){
  var tip=ensureTooltip();
  var date=formatDatePretty(node.date);
  if(!date){tip.classList.remove("visible");return;}
  tip.innerHTML='<span class="node-tooltip-date">'+date+"</span>";
  tip.style.left=e.clientX+"px";tip.style.top=(e.clientY-10)+"px";
  tip.classList.add("visible");
}
function moveTooltip(e){
  if(!tooltipEl)return;
  tooltipEl.style.left=e.clientX+"px";tooltipEl.style.top=(e.clientY-10)+"px";
}
function hideTooltip(){if(tooltipEl)tooltipEl.classList.remove("visible");}

/* ── Pan & Zoom (smooth) ──────────────────────────── */
var panState={active:false,startX:0,startY:0,baseTx:0,baseTy:0};

svg.addEventListener("mousedown",function(e){
  if(e.button!==0)return;
  panState.active=true;
  panState.startX=e.clientX;panState.startY=e.clientY;
  panState.baseTx=animTarget.tx;panState.baseTy=animTarget.ty;
  svg.style.cursor="grabbing";
  e.preventDefault();
});
window.addEventListener("mousemove",function(e){
  if(!panState.active)return;
  var mapping=getSvgScreenMapping();
  var dx=(e.clientX-panState.startX)/mapping.svgScale;
  var dy=(e.clientY-panState.startY)/mapping.svgScale;
  setTransformAnimated(panState.baseTx+dx,panState.baseTy+dy,animTarget.scale);
});
window.addEventListener("mouseup",function(){
  if(panState.active){panState.active=false;svg.style.cursor="";}
});

/* Touch support for panning */
var touchState={active:false,startX:0,startY:0,baseTx:0,baseTy:0,
  pinch:false,startDist:0,startScale:1,startMidX:0,startMidY:0};
function touchDist(t){var dx=t[0].clientX-t[1].clientX;var dy=t[0].clientY-t[1].clientY;return Math.sqrt(dx*dx+dy*dy);}
svg.addEventListener("touchstart",function(e){
  if(e.touches.length===1){
    touchState.active=true;touchState.pinch=false;
    touchState.startX=e.touches[0].clientX;touchState.startY=e.touches[0].clientY;
    touchState.baseTx=animTarget.tx;touchState.baseTy=animTarget.ty;
    e.preventDefault();
  }else if(e.touches.length===2){
    touchState.active=false;touchState.pinch=true;
    touchState.startDist=touchDist(e.touches);touchState.startScale=animTarget.scale;
    touchState.startMidX=(e.touches[0].clientX+e.touches[1].clientX)/2;
    touchState.startMidY=(e.touches[0].clientY+e.touches[1].clientY)/2;
    touchState.baseTx=animTarget.tx;touchState.baseTy=animTarget.ty;
    e.preventDefault();
  }
},{passive:false});
window.addEventListener("touchmove",function(e){
  if(touchState.active&&e.touches.length===1){
    var mapping=getSvgScreenMapping();
    var dx=(e.touches[0].clientX-touchState.startX)/mapping.svgScale;
    var dy=(e.touches[0].clientY-touchState.startY)/mapping.svgScale;
    setTransformAnimated(touchState.baseTx+dx,touchState.baseTy+dy,animTarget.scale);
    e.preventDefault();
  }else if(touchState.pinch&&e.touches.length===2){
    var dist=touchDist(e.touches);var ratio=dist/touchState.startDist;
    var ns=clamp(touchState.startScale*ratio,0.005,5);
    var cx=800;var cy2=500;
    var svgCX=(cx-touchState.baseTx)/touchState.startScale;
    var svgCY=(cy2-touchState.baseTy)/touchState.startScale;
    var ntx=cx-svgCX*ns;var nty=cy2-svgCY*ns;
    setTransformAnimated(ntx,nty,ns);
    e.preventDefault();
  }
},{passive:false});
window.addEventListener("touchend",function(){touchState.active=false;touchState.pinch=false;});

svg.addEventListener("wheel",function(e){
  e.preventDefault();
  var zoomFactor=e.deltaY<0?1.08:1/1.08;
  var oldScale=animTarget.scale;
  var nextScale=clamp(oldScale*zoomFactor,0.005,5);
  var cx=800;var cy2=500;
  var svgCX=(cx-animTarget.tx)/oldScale;
  var svgCY=(cy2-animTarget.ty)/oldScale;
  var ntx=cx-svgCX*nextScale;var nty=cy2-svgCY*nextScale;
  setTransformAnimated(ntx,nty,nextScale);
},{passive:false});

function clamp(v,mn,mx){return Math.max(mn,Math.min(mx,v));}

/* ── Fit to view ───────────────────────────────────── */
function fitToView(){
  var pts=[];
  DATA.streams.forEach(function(s){s.points.forEach(function(p){pts.push(p);});});
  DATA.nodes.forEach(function(n){pts.push({x:n.x,y:n.y});});
  DATA.textLabels.forEach(function(tl){pts.push({x:tl.x,y:tl.y});});
  if(pts.length===0){setTransformImmediate(0,0,1);return;}
  var minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity;
  pts.forEach(function(p){if(p.x<minX)minX=p.x;if(p.y<minY)minY=p.y;if(p.x>maxX)maxX=p.x;if(p.y>maxY)maxY=p.y;});
  var mw=maxX-minX+160;var mh=maxY-minY+160;
  var sX=1600/mw;var sY=1000/mh;
  var s=clamp(Math.min(sX,sY),0.005,2.2);
  var ntx=800-((minX+maxX)/2)*s;var nty=500-((minY+maxY)/2)*s;
  setTransformImmediate(ntx,nty,s);
}

function goToToday(){
  var now=new Date();var year=now.getFullYear();var month=now.getMonth()+1;
  var day=now.getDate();var daysInMonth=new Date(year,month,0).getDate();
  var monthX=calMonthSvgX(year,month);
  var w=calMonthWidth(calMonthKey(year,month));
  var todayX=monthX+w*(day/daysInMonth);
  setTransformAnimated(800-todayX*animTarget.scale,animTarget.ty,animTarget.scale);
}

/* ── Minimap ────────────────────────────────────────── */
var minimapEl=document.getElementById("minimap");
var minimapCanvas=document.getElementById("minimapCanvas");
var minimapVp=document.getElementById("minimapViewport");
var mmCtx=minimapCanvas.getContext("2d");
var contentBounds=null;

function getContentBounds(){
  var pts=[];
  DATA.streams.forEach(function(s){s.points.forEach(function(p){pts.push(p);});});
  DATA.nodes.forEach(function(n){
    pts.push({x:n.x,y:n.y});
    var mw=n.labelMaxWidth||DEFAULT_LABEL_MAX_WIDTH;
    pts.push({x:n.x-mw/2,y:n.y});
    pts.push({x:n.x+mw/2,y:n.y});
    var subs=n.subLabels||[];
    var labelCount=1+subs.length;
    var estLineH=labelFontSize*1.4;
    var estLabelH=22+labelCount*(estLineH*2+6+8);
    pts.push({x:n.x,y:n.y+estLabelH});
  });
  DATA.textLabels.forEach(function(tl){
    var mw=tl.maxWidth||160;var fs=tl.fontSize||16;
    pts.push({x:tl.x-mw/2,y:tl.y});
    pts.push({x:tl.x+mw/2,y:tl.y});
    var text=tl.text||"";
    var charsPerLine=Math.max(1,Math.floor(mw/(fs*0.6)));
    var estLines=Math.max(1,Math.ceil(text.length/charsPerLine));
    pts.push({x:tl.x,y:tl.y+estLines*fs*1.4});
  });
  if(pts.length===0)return{x:0,y:0,w:1600,h:1000};
  var minX=Infinity,minY=Infinity,maxX=-Infinity,maxY=-Infinity;
  pts.forEach(function(p){if(p.x<minX)minX=p.x;if(p.y<minY)minY=p.y;if(p.x>maxX)maxX=p.x;if(p.y>maxY)maxY=p.y;});
  var pad=80;
  return{x:minX-pad,y:minY-pad,w:maxX-minX+pad*2,h:maxY-minY+pad*2};
}

function renderMinimap(){
  var dpr=window.devicePixelRatio||1;
  var cw=minimapEl.clientWidth;var ch=minimapEl.clientHeight;
  minimapCanvas.width=cw*dpr;minimapCanvas.height=ch*dpr;
  mmCtx.setTransform(dpr,0,0,dpr,0,0);
  mmCtx.clearRect(0,0,cw,ch);
  if(!contentBounds)return;
  var b=contentBounds;
  var scaleX=cw/b.w;var scaleY=ch/b.h;
  var ms=Math.min(scaleX,scaleY)*0.9;
  var ox=(cw-b.w*ms)/2;var oy=(ch-b.h*ms)/2;
  function mx(x){return ox+(x-b.x)*ms;}
  function my(y){return oy+(y-b.y)*ms;}

  // Draw connections
  mmCtx.strokeStyle=getComputedStyle(document.documentElement).getPropertyValue("--connection-stroke").trim()||"rgba(0,0,0,0.3)";
  mmCtx.lineWidth=Math.max(0.3,ms*2);mmCtx.setLineDash([Math.max(1,ms*6),Math.max(1,ms*6)]);
  DATA.connections.forEach(function(conn){
    var a=DATA.nodes.find(function(n){return n.id===conn.fromNodeId;});
    var b2=DATA.nodes.find(function(n){return n.id===conn.toNodeId;});
    if(!a||!b2)return;
    mmCtx.beginPath();mmCtx.moveTo(mx(a.x),my(a.y));mmCtx.lineTo(mx(b2.x),my(b2.y));mmCtx.stroke();
  });
  mmCtx.setLineDash([]);

  // Draw streams
  var sw=Math.max(1.5,Math.min(4,ms*16));
  mmCtx.lineWidth=sw;mmCtx.lineCap="round";mmCtx.lineJoin="round";
  DATA.streams.forEach(function(s){
    if(s.points.length<2)return;
    mmCtx.strokeStyle=s.color;mmCtx.beginPath();
    mmCtx.moveTo(mx(s.points[0].x),my(s.points[0].y));
    for(var i=1;i<s.points.length;i++) mmCtx.lineTo(mx(s.points[i].x),my(s.points[i].y));
    mmCtx.stroke();
  });

  // Draw nodes
  var nr=Math.max(1,Math.min(3,ms*10));
  DATA.nodes.forEach(function(node){
    var fill=STATUS_COLORS[node.status]||STATUS_COLORS.planned;
    mmCtx.fillStyle=fill;mmCtx.strokeStyle="#53565A";mmCtx.lineWidth=Math.max(0.3,ms*1.5);
    if(node.type==="activity"){
      mmCtx.beginPath();mmCtx.arc(mx(node.x),my(node.y),nr,0,Math.PI*2);mmCtx.fill();mmCtx.stroke();
    }else{
      var px=mx(node.x);var py=my(node.y);
      mmCtx.save();mmCtx.translate(px,py);mmCtx.rotate(Math.PI/4);
      mmCtx.fillRect(-nr,-nr,nr*2,nr*2);mmCtx.strokeRect(-nr,-nr,nr*2,nr*2);
      mmCtx.restore();
    }
  });
}

function updateMinimapViewport(){
  if(!contentBounds)return;
  var b=contentBounds;
  var cw=minimapEl.clientWidth;var ch=minimapEl.clientHeight;
  var scaleX=cw/b.w;var scaleY=ch/b.h;
  var ms=Math.min(scaleX,scaleY)*0.9;
  var ox=(cw-b.w*ms)/2;var oy=(ch-b.h*ms)/2;

  // Compute visible world rect from current transform
  // The SVG viewBox is 0..1600 x 0..1000, viewport transform is translate(tx,ty) scale(scale)
  // Visible world x range: (0 - tx)/scale .. (1600 - tx)/scale
  var vx0=(0-tx)/scale;var vy0=(0-ty)/scale;
  var vx1=(1600-tx)/scale;var vy1=(1000-ty)/scale;

  // Map to minimap pixels
  var left=ox+(vx0-b.x)*ms;var top2=oy+(vy0-b.y)*ms;
  var right=ox+(vx1-b.x)*ms;var bot=oy+(vy1-b.y)*ms;
  var w=right-left;var h=bot-top2;

  // Clamp to minimap bounds
  minimapVp.style.left=Math.max(0,left)+"px";
  minimapVp.style.top=Math.max(0,top2)+"px";
  minimapVp.style.width=Math.min(w,cw-Math.max(0,left))+"px";
  minimapVp.style.height=Math.min(h,ch-Math.max(0,top2))+"px";
  minimapVp.style.display=(w<cw*0.95||h<ch*0.95)?"block":"none";
}

// Click/drag on minimap to navigate
var mmDrag={active:false};
function minimapNavigate(e){
  var rect=minimapEl.getBoundingClientRect();
  var mx2=e.clientX-rect.left;var my2=e.clientY-rect.top;
  var b=contentBounds;if(!b)return;
  var cw=minimapEl.clientWidth;var ch=minimapEl.clientHeight;
  var scaleX=cw/b.w;var scaleY=ch/b.h;
  var ms=Math.min(scaleX,scaleY)*0.9;
  var ox=(cw-b.w*ms)/2;var oy=(ch-b.h*ms)/2;
  // Convert minimap pixel to world coords
  var worldX=b.x+(mx2-ox)/ms;
  var worldY=b.y+(my2-oy)/ms;
  // Center on that world point
  var ntx=800-worldX*animTarget.scale;
  var nty=500-worldY*animTarget.scale;
  setTransformAnimated(ntx,nty,animTarget.scale);
}
minimapEl.addEventListener("mousedown",function(e){
  e.preventDefault();e.stopPropagation();
  mmDrag.active=true;minimapNavigate(e);
});
window.addEventListener("mousemove",function(e){
  if(mmDrag.active){e.preventDefault();minimapNavigate(e);}
});
window.addEventListener("mouseup",function(){mmDrag.active=false;});

// Hook into transform updates
var _origApply=applyTransform;
applyTransform=function(){_origApply();updateMinimapViewport();};

/* ── Init ──────────────────────────────────────────── */
contentBounds=getContentBounds();
renderStreams();renderConnections();renderNodes();renderLabels();renderLegend();
renderMinimap();
fitToView();
document.getElementById("fitViewBtn").addEventListener("click",fitToView);
document.getElementById("todayBtn").addEventListener("click",goToToday);
window.addEventListener("resize",function(){renderCalendar();renderMinimap();updateMinimapViewport();});
})();
<\/script>
</body>
</html>`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

// --- Guides (context menu) ---

let contextMenuPoint = null;
let contextMenuNodeId = null;

function showCanvasContextMenu(event) {
  event.preventDefault();
  canvasContextMenu.style.display = "grid";
  canvasContextMenu.style.left = event.clientX + "px";
  canvasContextMenu.style.top = event.clientY + "px";
  contextMenuPoint = svgPointFromEvent(event);

  // Clean up previous dynamic entries
  canvasContextMenu.querySelectorAll("[data-action='delete-guide'], [data-action='delete-element'], .ctx-separator, [data-action^='set-status-'], .ctx-submenu-wrap, [data-action^='arrange-']").forEach((el) => el.remove());
  contextMenuNodeId = null;
  canvasContextMenu._deleteTarget = null;
  canvasContextMenu._arrangeTarget = null;

  // Detect what element was right-clicked
  const nodeG = event.target.closest("g[data-node-id]");
  const streamLine = event.target.closest("[data-stream-id]");
  const connEl = event.target.closest("[data-connection-id]");
  const textLabelG = event.target.closest("g[data-text-label-id]");
  const guideHit = event.target.closest(".guide-hit");

  // If right-clicked on a node, add status change options
  if (nodeG) {
    contextMenuNodeId = nodeG.dataset.nodeId;
    const node = state.nodes.find((n) => n.id === contextMenuNodeId);
    if (node) {
      const sep = document.createElement("div");
      sep.className = "ctx-separator";
      canvasContextMenu.appendChild(sep);

      const statuses = [
        { key: "planned", label: "Planned", color: STATUS_COLORS.planned },
        { key: "completed", label: "Completed", color: STATUS_COLORS.completed },
        { key: "ongoing", label: "On-going", color: STATUS_COLORS.ongoing },
        { key: "attention", label: "Needs Attention", color: STATUS_COLORS.attention },
      ];
      statuses.forEach((s) => {
        const btn = document.createElement("button");
        btn.dataset.action = "set-status-" + s.key;
        const dot = document.createElement("span");
        dot.style.cssText = "display:inline-block;width:10px;height:10px;border-radius:50%;border:1px solid rgba(0,0,0,0.25);margin-right:8px;vertical-align:middle;background:" + s.color + ";";
        btn.appendChild(dot);
        btn.appendChild(document.createTextNode(s.label));
        if (node.status === s.key) {
          btn.style.fontWeight = "700";
          btn.style.borderLeft = "3px solid var(--accent)";
        }
        canvasContextMenu.appendChild(btn);
      });
    }
  }

  // Determine delete target (priority: node > connection > text-label > stream > guide)
  let deleteLabel = null;
  if (nodeG) {
    canvasContextMenu._deleteTarget = { type: "node", id: nodeG.dataset.nodeId };
    deleteLabel = "Delete Node";
  } else if (connEl) {
    canvasContextMenu._deleteTarget = { type: "connection", id: connEl.dataset.connectionId };
    deleteLabel = "Delete Connection";
  } else if (textLabelG) {
    canvasContextMenu._deleteTarget = { type: "text-label", id: textLabelG.dataset.textLabelId };
    deleteLabel = "Delete Text Label";
  } else if (streamLine) {
    const clickedStreamId = streamLine.dataset.streamId;
    const groupIds = getSelectedStreamGroup();
    if (groupIds.size > 1 && groupIds.has(clickedStreamId)) {
      canvasContextMenu._deleteTarget = { type: "stream-group", ids: [...groupIds] };
      deleteLabel = "Delete Stream";
    } else {
      canvasContextMenu._deleteTarget = { type: "stream", id: clickedStreamId };
      deleteLabel = "Delete Stream Segment";
    }
  } else if (guideHit) {
    canvasContextMenu._deleteTarget = { type: "guide", id: guideHit.dataset.guideId };
    deleteLabel = "Delete Guide";
  }

  // Arrange submenu for streams and connections
  if (streamLine || connEl) {
    const arrangeType = streamLine ? "stream" : "connection";
    const arrangeId = streamLine ? streamLine.dataset.streamId : connEl.dataset.connectionId;
    canvasContextMenu._arrangeTarget = { type: arrangeType, id: arrangeId };

    const sep = document.createElement("div");
    sep.className = "ctx-separator";
    canvasContextMenu.appendChild(sep);

    const wrap = document.createElement("div");
    wrap.className = "ctx-submenu-wrap";

    const trigger = document.createElement("button");
    trigger.className = "ctx-submenu-trigger";
    trigger.innerHTML = `<span class="ctx-icon">⇅</span>Arrange<span class="ctx-arrow">›</span>`;
    wrap.appendChild(trigger);

    const sub = document.createElement("div");
    sub.className = "ctx-submenu";

    const items = [
      { action: "arrange-bring-to-front", icon: "⤒", label: "Bring to Front" },
      { action: "arrange-bring-forward", icon: "↑", label: "Bring Forward" },
      { action: "arrange-send-backward", icon: "↓", label: "Send Backward" },
      { action: "arrange-send-to-back", icon: "⤓", label: "Send to Back" },
    ];
    items.forEach((item) => {
      const btn = document.createElement("button");
      btn.dataset.action = item.action;
      btn.innerHTML = `<span class="ctx-icon">${item.icon}</span>${item.label}`;
      sub.appendChild(btn);
    });

    wrap.appendChild(sub);
    canvasContextMenu.appendChild(wrap);
  }

  if (deleteLabel) {
    const sep = document.createElement("div");
    sep.className = "ctx-separator";
    canvasContextMenu.appendChild(sep);

    const delBtn = document.createElement("button");
    delBtn.dataset.action = "delete-element";
    delBtn.innerHTML = `<span class="ctx-icon">&#xd7;</span>${deleteLabel}`;
    delBtn.className = "danger";
    canvasContextMenu.appendChild(delBtn);
  }

  // Legacy: If right-clicked on a guide, also keep the old delete-guide action
  if (guideHit) {
    const delBtn = document.createElement("button");
    delBtn.dataset.action = "delete-guide";
    delBtn.dataset.guideId = guideHit.dataset.guideId;
    delBtn.textContent = "Delete Guide";
    delBtn.className = "danger";
    delBtn.style.display = "none";
    canvasContextMenu.appendChild(delBtn);
  }
}

function hideCanvasContextMenu() {
  canvasContextMenu.style.display = "none";
  contextMenuPoint = null;
}

function addGuide(axis) {
  if (!contextMenuPoint) return;
  captureHistorySnapshot();
  const position = axis === "v"
    ? snap(contextMenuPoint.x)
    : snap(contextMenuPoint.y);
  state.guides.push({ id: uid("guide"), axis, position });
  render();
}

function deleteGuide(guideId) {
  const idx = state.guides.findIndex((g) => g.id === guideId);
  if (idx === -1) return;
  captureHistorySnapshot();
  state.guides.splice(idx, 1);
  render();
}

function initEvents() {
  toolButtons.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-tool]");
    if (!button) {
      return;
    }
    setTool(button.dataset.tool);
  });

  undoBtn.addEventListener("click", () => { undo(); syncStreamPresetDropdown(); });
  redoBtn.addEventListener("click", () => { redo(); syncStreamPresetDropdown(); });
  fitToViewBtn.addEventListener("click", fitToView);
  goToTodayBtn.addEventListener("click", goToToday);
  toggleGridBtn.addEventListener("click", () => {
    const grid = document.querySelector(".canvas-grid");
    const visible = grid.getAttribute("visibility") !== "hidden";
    grid.setAttribute("visibility", visible ? "hidden" : "visible");
    toggleGridBtn.classList.toggle("active", !visible);
  });

  function updateGridPatternColor() {
    const gridPath = document.querySelector("#gridPattern path");
    if (gridPath) {
      const isDark = document.documentElement.getAttribute("data-theme") === "dark";
      gridPath.setAttribute("stroke", isDark ? "rgba(255,255,255,0.06)" : "rgba(0,0,0,0.06)");
    }
  }

  themeToggle.addEventListener("click", () => {
    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    document.documentElement.setAttribute("data-theme", isDark ? "light" : "dark");
    updateGridPatternColor();
    localStorage.setItem("metrohack-theme", isDark ? "light" : "dark");
  });

  // Apply saved theme or default to light
  const savedTheme = localStorage.getItem("metrohack-theme") || "light";
  if (savedTheme === "dark") {
    document.documentElement.setAttribute("data-theme", "dark");
  }
  updateGridPatternColor();
  document.getElementById("toggleSnap").addEventListener("click", () => {
    snapEnabled = !snapEnabled;
    document.getElementById("toggleSnap").classList.toggle("active", snapEnabled);
  });

  document.getElementById("toggleCalendar").addEventListener("click", () => {
    const btn = document.getElementById("toggleCalendar");
    const isVisible = !calendarBar.classList.contains("calendar-hidden");
    calendarBar.classList.toggle("calendar-hidden", isVisible);
    calendarGuidesLayer.style.opacity = isVisible ? "0" : "1";
    btn.classList.toggle("active", !isVisible);
  });

  document.getElementById("labelFontSize").addEventListener("input", (event) => {
    labelFontSize = Number(event.target.value);
    document.getElementById("labelFontSizeValue").textContent = labelFontSize + "px";
    render();
  });
  deleteSelectedBtn.addEventListener("click", removeSelected);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Delete" && !event.target.closest("input, select, textarea")) {
      removeSelected();
    }
  });
  exportPngBtn.addEventListener("click", exportPng);
  exportSvgBtn.addEventListener("click", exportSvg);
  exportLegendPngBtn.addEventListener("click", exportLegendPng);
  exportLegendSvgBtn.addEventListener("click", exportLegendSvg);
  exportViewerBtn.addEventListener("click", exportViewer);
  saveMapBtn.addEventListener("click", saveMap);
  streamPresetSelect.addEventListener("change", applyStreamPreset);
  loadMapBtn.addEventListener("click", () => loadMapFileInput.click());
  loadMapFileInput.addEventListener("change", (event) => {
    const file = event.target.files[0];
    if (file) {
      loadMap(file);
      loadMapFileInput.value = "";
    }
  });

  wirePanAndZoom();

  // Calendar border drag
  calendarBar.addEventListener("mousedown", (e) => {
    const border = e.target.closest(".cal-border");
    if (border && e.button === 0) {
      e.preventDefault();
      e.stopPropagation();
      const key = border.dataset.monthKey;
      state.interaction.mode = "calendar-border-drag";
      state.interaction.calendarDragKey = key;
      state.interaction.calendarDragStartX = e.clientX;
      state.interaction.calendarDragStartWidth = calMonthWidth(key);
      state.interaction.historyCaptured = false;
      state.pan.startX = e.clientX;
      state.pan.startY = e.clientY;
    }
  });

  window.addEventListener("resize", () => renderCalendar());

  // Context menu for guides
  svg.addEventListener("contextmenu", showCanvasContextMenu);
  document.addEventListener("mousedown", (event) => {
    if (!canvasContextMenu.contains(event.target)) {
      hideCanvasContextMenu();
    }
  });
  canvasContextMenu.addEventListener("click", (event) => {
    const btn = event.target.closest("button[data-action]");
    if (!btn) return;
    const action = btn.dataset.action;
    if (action === "add-guide-v") addGuide("v");
    else if (action === "add-guide-h") addGuide("h");
    else if (action === "delete-guide") deleteGuide(btn.dataset.guideId);
    else if (action === "delete-element" && canvasContextMenu._deleteTarget) {
      const target = canvasContextMenu._deleteTarget;
      captureHistorySnapshot();
      if (target.type === "node") {
        state.nodes = state.nodes.filter((n) => n.id !== target.id);
        state.connections = state.connections.filter(
          (c) => c.fromNodeId !== target.id && c.toNodeId !== target.id,
        );
      } else if (target.type === "stream") {
        const nodesOnStream = state.nodes.filter((n) => n.streamId === target.id);
        const nodeIds = new Set(nodesOnStream.map((n) => n.id));
        state.streams = state.streams.filter((s) => s.id !== target.id);
        state.nodes = state.nodes.filter((n) => n.streamId !== target.id);
        state.connections = state.connections.filter(
          (c) => !nodeIds.has(c.fromNodeId) && !nodeIds.has(c.toNodeId),
        );
      } else if (target.type === "stream-group") {
        const idsToRemove = new Set(target.ids);
        const allNodeIds = new Set();
        idsToRemove.forEach((sid) => {
          state.nodes.filter((n) => n.streamId === sid).forEach((n) => allNodeIds.add(n.id));
        });
        state.streams = state.streams.filter((s) => !idsToRemove.has(s.id));
        state.nodes = state.nodes.filter((n) => !allNodeIds.has(n.id));
        state.connections = state.connections.filter(
          (c) => !allNodeIds.has(c.fromNodeId) && !allNodeIds.has(c.toNodeId),
        );
      } else if (target.type === "connection") {
        state.connections = state.connections.filter((c) => c.id !== target.id);
      } else if (target.type === "text-label") {
        state.textLabels = state.textLabels.filter((t) => t.id !== target.id);
      } else if (target.type === "guide") {
        deleteGuide(target.id);
      }
      clearSelection();
      syncStreamPresetDropdown();
      applyStreamPreset();
      render();
      renderSelectionEditor();
    }
    else if (action.startsWith("set-status-") && contextMenuNodeId) {
      const newStatus = action.replace("set-status-", "");
      const node = state.nodes.find((n) => n.id === contextMenuNodeId);
      if (node && node.status !== newStatus) {
        captureHistorySnapshot();
        node.status = newStatus;
        render();
        renderSelectionEditor();
      }
    }
    else if (action.startsWith("arrange-") && canvasContextMenu._arrangeTarget) {
      const target = canvasContextMenu._arrangeTarget;
      const arrangeAction = action.replace("arrange-", "");
      arrangeElement(target.type, target.id, arrangeAction);
    }
    hideCanvasContextMenu();
  });
}

function init() {
  setTool("select");
  initColorPalette();
  syncStreamPresetDropdown();
  applyStreamPreset();
  updateUndoRedoButtons();
  renderSelectionEditor();
  render();
  fitToView();
  initEvents();
}

init();
