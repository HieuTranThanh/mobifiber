/* =====================================================
   BIẾN TOÀN CỤC & CẤU HÌNH
   ===================================================== */

// Cấu hình tập trung để dễ chỉnh khi đổi nguồn dữ liệu hoặc rule nghiệp vụ.
const DEFAULT_CONFIG = {
  SHEET_ID: "1AkswJCHRClc7wAoagpRqtrG9OATm5_6qsRK0gTHZqcY",
  SHEET_TAPDIEM: "Danh sách tập điểm S2",
  SHEET_TAPDIEM_GID: "1472299907",
  SHEET_CUSTOMER: "Danh sách khách hàng",
  SHEET_CUSTOMER_GID: "341989509",
  FEEDBACK_FORM_URL: "https://docs.google.com/forms/d/e/1FAIpQLScf5EuKEWzllRuqEbjtLIIU_n0Z5HU7wZhpRW6NDkVh3Z36gw/viewform",
  TRACKING_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbypr-0OMKp6UopO_-cET65QxdwTpeoMV9q13nJMdt3PBR-GtTHkKBrfHsdqTJR24bko/exec",
  POINT_RADIUS_METERS: 220,
  MAX_SALE_DISTANCE_METERS: 1000,
  MAX_SUGGESTED_POINTS: 3
};

const CONFIG = {
  ...DEFAULT_CONFIG,
  ...(window.APP_CONFIG || {})
};

// Đối tượng map Leaflet
let map;

// Danh sách toàn bộ tập điểm hợp lệ
let allPoints = [];

// Map lưu danh sách khách hàng theo tập điểm
// key: tên tập điểm | value: mảng khách hàng
let customerByTapDiem = {};

// Danh sách khách hàng phục vụ search
let customerSearchList = [];

// Marker vị trí khách hàng
let customerMarker = null;

// Tọa độ khách hàng hiện tại (null = chưa có)
let customerLatLng = null;

// Các layer gợi ý (line, highlight)
let suggestionLayers = [];

// Cờ xác định lần load đầu tiên (fitBounds)
let isInitialLoad = true;

// Marker vị trí khách hàng khi double-click trên map
let clickMarker = null;

// Trang thai tai du lieu
let isPointDataReady = false;
let isCustomerDataReady = false;
let pointByName = new Map();
let normalizedPointBuckets = new Map();
let unmatchedCustomers = [];

/* =====================================================
   KHỞI TẠO MAP
   (tọa độ ban đầu mang tính mặc định, sẽ fitBounds lại
    sau khi load danh sách tập điểm)
   ===================================================== */

map = L.map('map', {
  doubleClickZoom: false,
  zoomControl: false
}).setView([12.688165492644124, 108.05569162987426], 11);  

// Tile OpenStreetMap
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  maxZoom: 19
}).addTo(map);

/* Double click trên bản đồ để xác định vị trí khách hàng
   (ghi đè marker cũ và kích hoạt logic gợi ý tập điểm) */
map.on("dblclick", function (e) {
  const lat = e.latlng.lat.toFixed(6);
  const lng = e.latlng.lng.toFixed(6);

  // Xóa marker cũ
  if (clickMarker) map.removeLayer(clickMarker);

  // Tạo marker mới
  clickMarker = L.marker([lat, lng]).addTo(map);

  // Điền tọa độ vào ô tìm kiếm
  document.getElementById("searchInput").value = `${lat}, ${lng}`;

  // Xử lý logic tìm tập điểm
  handleCustomer(+lat, +lng);
  collapseControlBox();
});

/* =====================================================
   HÀM TIỆN ÍCH (UTILS)
   ===================================================== */

/* Kiểm tra tọa độ hợp lệ */
function validateLatLng(lat, lng) {
  if (lat === null || lat === undefined) return false;
  if (lng === null || lng === undefined) return false;
  if (String(lat).trim() === "") return false;
  if (String(lng).trim() === "") return false;

  lat = Number(lat);
  lng = Number(lng);

  if (!Number.isFinite(lat) || !Number.isFinite(lng)) return false;
  if (lat === 0 && lng === 0) return false;

  return lat >= -90 && lat <= 90 && lng >= -180 && lng <= 180;
}

/* Xác định màu theo hiệu suất */
function getColor(eff) {
  if (eff < 0.2) return 'blue';
  if (eff < 0.5) return 'green';
  if (eff < 0.8) return 'orange';
  if (eff < 1) return 'red';
  return 'black';
}

/* Quy đổi màu sang format KML */
function colorToKML(color) {
  const map = {
    blue:  "7dff0000",
    green: "7d00ff00",
    orange: "7d00a5ff",
    red:   "7d0000ff",
    black: "7d000000"
  };
  return map[color] || "7dffffff";
}

/* Tính khoảng cách Haversine (mét) */
function haversine(lat1, lng1, lat2, lng2) {
  const R = 6371000;
  const toRad = d => d * Math.PI / 180;
  const dLat = toRad(lat2 - lat1);
  const dLng = toRad(lng2 - lng1);
  const a = Math.sin(dLat/2)**2 +
    Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
    Math.sin(dLng/2)**2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

/* Chuyển circle Leaflet thành polygon cho KML */
function circleToPolygon(lat, lng, radiusMeters, points = 60) {
  const coords = [];
  const earthRadius = 6378137;

  for (let i = 0; i <= points; i++) {
    const angle = (i / points) * 2 * Math.PI;
    const dx = radiusMeters * Math.cos(angle);
    const dy = radiusMeters * Math.sin(angle);

    const dLat = dy / earthRadius;
    const dLng = dx / (earthRadius * Math.cos(lat * Math.PI / 180));

    const latPoint = lat + dLat * 180 / Math.PI;
    const lngPoint = lng + dLng * 180 / Math.PI;

    coords.push(`${lngPoint},${latPoint},0`);
  }

  return coords.join(" ");
}

function normalizeText(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeHeader(value) {
  return normalizeText(value)
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function normalizeKey(value) {
  return normalizeText(value).replace(/\s+/g, "");
}

function canonicalName(value) {
  return String(value ?? "").trim();
}

function normalizeAddressQuery(value) {
  return canonicalName(value)
    .replace(/\s*,\s*/g, ", ")
    .replace(/\s+/g, " ");
}

function ensureVietnamSuffix(value) {
  const cleaned = normalizeAddressQuery(value);
  if (!cleaned) return "";

  const normalized = normalizeText(cleaned);
  if (normalized.includes("viet nam") || normalized.includes("vietnam")) {
    return cleaned;
  }

  return `${cleaned}, Việt Nam`;
}

function simplifyAddressQuery(value) {
  const cleaned = normalizeAddressQuery(value);
  if (!cleaned) return "";

  const parts = cleaned.split(",").map(part => part.trim()).filter(Boolean);
  if (!parts.length) return "";

  const simplifiedFirstPart = parts[0]
    .replace(/^(so|số)\s+/i, "")
    .replace(/^[0-9a-zA-Z]+(?:[\/-][0-9a-zA-Z]+)*(?:\s+[0-9a-zA-Z]+(?:[\/-][0-9a-zA-Z]+)*)*\s+/u, "")
    .trim();

  if (simplifiedFirstPart) {
    parts[0] = simplifiedFirstPart;
  }

  return parts.join(", ");
}

function getPointSearchBounds() {
  if (!allPoints.length) return null;

  const lats = allPoints.map(point => point.lat);
  const lngs = allPoints.map(point => point.lng);
  const minLat = Math.min(...lats);
  const maxLat = Math.max(...lats);
  const minLng = Math.min(...lngs);
  const maxLng = Math.max(...lngs);
  const latPad = Math.max((maxLat - minLat) * 0.15, 0.05);
  const lngPad = Math.max((maxLng - minLng) * 0.15, 0.05);

  return {
    left: minLng - lngPad,
    top: maxLat + latPad,
    right: maxLng + lngPad,
    bottom: minLat - latPad
  };
}

function buildAddressSearchPlan(rawQuery) {
  const exact = normalizeAddressQuery(rawQuery);
  const simplified = simplifyAddressQuery(exact);
  const plan = [];

  const addAttempt = (query, bounded) => {
    const finalQuery = ensureVietnamSuffix(query);
    if (!finalQuery) return;
    if (plan.some(item => item.query === finalQuery && item.bounded === bounded)) return;
    plan.push({ query: finalQuery, bounded });
  };

  addAttempt(exact, true);
  addAttempt(exact, false);

  if (simplified && normalizeText(simplified) !== normalizeText(exact)) {
    addAttempt(simplified, true);
    addAttempt(simplified, false);
  }

  return plan;
}

function buildNominatimUrl(query, { limit = 5, bounded = false } = {}) {
  const params = new URLSearchParams({
    format: "json",
    limit: String(limit),
    countrycodes: "vn",
    addressdetails: "1",
    q: query
  });

  if (bounded) {
    const bounds = getPointSearchBounds();
    if (bounds) {
      params.set("viewbox", `${bounds.left},${bounds.top},${bounds.right},${bounds.bottom}`);
      params.set("bounded", "1");
    }
  }

  return `https://nominatim.openstreetmap.org/search?${params.toString()}`;
}

async function geocodeAddress(rawQuery, limit = 5) {
  const plan = buildAddressSearchPlan(rawQuery);

  for (const attempt of plan) {
    const response = await fetch(buildNominatimUrl(attempt.query, {
      limit,
      bounded: attempt.bounded
    }));

    if (!response.ok) {
      throw new Error(`Nominatim HTTP ${response.status}`);
    }

    const data = await response.json();
    if (Array.isArray(data) && data.length) {
      return data;
    }
  }

  return [];
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, ch => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[ch]));
}

function escapeXml(value) {
  return String(value ?? "").replace(/[<>&'"]/g, ch => ({
    "<": "&lt;",
    ">": "&gt;",
    "&": "&amp;",
    "'": "&apos;",
    '"': "&quot;"
  }[ch]));
}

function makeSafeId(prefix, value) {
  const normalized = normalizeHeader(value) || "item";
  return `${prefix}-${normalized}`;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    const next = text[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        field += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === "," && !inQuotes) {
      row.push(field);
      field = "";
      continue;
    }

    if ((ch === "\n" || ch === "\r") && !inQuotes) {
      if (ch === "\r" && next === "\n") i++;
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      continue;
    }

    field += ch;
  }

  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

function rowsToObjects(csvRows) {
  if (!csvRows.length) return [];

  const headers = csvRows[0].map(h => h.trim());
  return csvRows.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index] ?? "";
    });
    return obj;
  });
}

function hasAnyHeader(headers, aliases) {
  const aliasList = Array.isArray(aliases) ? aliases : [aliases];
  return aliasList.some(alias => {
    const target = normalizeHeader(alias);
    return headers.some(header => normalizeHeader(header) === target);
  });
}

function validateSheetRows(rows, requiredHeaders) {
  if (!rows.length) return false;

  const headers = Object.keys(rows[0]);
  return requiredHeaders.every(group => hasAnyHeader(headers, group));
}

function getField(row, aliases) {
  const aliasList = Array.isArray(aliases) ? aliases : [aliases];
  const entries = Object.entries(row);

  for (const alias of aliasList) {
    const target = normalizeHeader(alias);
    const direct = entries.find(([key]) => normalizeHeader(key) === target);
    if (direct) return direct[1];
  }

  return "";
}

function toNumber(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return NaN;

  const cleaned = raw
    .replace(/\s+/g, "")
    .replace(/%/g, "")
    .replace(/\.(?=.*\.)/g, "")
    .replace(",", ".");

  return Number(cleaned);
}

function toPercentRatio(value) {
  const raw = String(value ?? "").trim();
  if (!raw) return NaN;

  const numeric = toNumber(raw);
  if (!Number.isFinite(numeric)) return NaN;
  if (raw.includes("%")) return numeric / 100;
  return numeric > 1 ? numeric / 100 : numeric;
}

async function fetchSheetRows({ gid, sheetName, preferExport = false, requiredHeaders = [] }) {
  const exportUrl = gid
    ? `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/export?format=csv&gid=${gid}`
    : null;
  const gvizCsvUrl = sheetName
    ? `https://docs.google.com/spreadsheets/d/${CONFIG.SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(sheetName)}`
    : null;

  const urls = preferExport
    ? [exportUrl, gvizCsvUrl].filter(Boolean)
    : [gvizCsvUrl, exportUrl].filter(Boolean);

  let lastError = null;

  for (const url of urls) {
    try {
      const response = await fetch(url);
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
      }

      const text = await response.text();
      const rows = rowsToObjects(parseCsv(text));
      if (requiredHeaders.length && !validateSheetRows(rows, requiredHeaders)) {
        throw new Error(`Schema mismatch for ${url}`);
      }

      return rows;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error("Không đọc được dữ liệu");
}

function registerPoint(point) {
  pointByName.set(point.nameKey, point);

  const normalized = normalizeKey(point.name);
  const bucket = normalizedPointBuckets.get(normalized) || [];
  bucket.push(point);
  normalizedPointBuckets.set(normalized, bucket);
}

function resolvePointByTapDiemName(tapDiemName) {
  const exactKey = canonicalName(tapDiemName);
  if (pointByName.has(exactKey)) {
    return pointByName.get(exactKey);
  }

  const normalized = normalizeKey(tapDiemName);
  const bucket = normalizedPointBuckets.get(normalized) || [];
  return bucket.length === 1 ? bucket[0] : null;
}

function buildBasePopup(point) {
  const directionHtml = customerLatLng
    ? renderDirection(point.lat, point.lng)
    : `<i style="color:#999">Chưa lấy được vị trí khách hàng</i>`;

  return `
    <b>${escapeHtml(point.name)}</b><br>

    <b>Tọa độ:</b>
    <a href="https://www.google.com/maps?q=${point.lat},${point.lng}" target="_blank" rel="noopener">
      ${point.lat.toFixed(6)}, ${point.lng.toFixed(6)}
    </a><br>

    <b>Dẫn đường:</b>
    <span class="direction-holder" data-lat="${point.lat}" data-lng="${point.lng}">
      ${directionHtml}
    </span><br>

    <b>Dung lượng:</b> ${escapeHtml(point.maxPort)} Port<br>
    <b>Đã dùng:</b> ${escapeHtml(point.usedPort)}<br>
    <b>Còn lại:</b> ${escapeHtml(point.freePort)}<br>
    <b>Hiệu suất:</b> ${(point.eff * 100).toFixed(1)}%
    <hr>

    <a href="#" class="show-customer-link" data-point-id="${point.id}">
      Xem danh sách khách hàng hiện tại
    </a>

    <div id="${point.customerContainerId}" style="margin-top:6px"></div>
  `;
}

// Render danh sách khách hàng thuộc một tập điểm (dùng trong popup)
function renderCustomerList(point) {
  const list = customerByTapDiem[point.id] || [];

  if (!isCustomerDataReady) {
    return `<i>Đang tải dữ liệu khách hàng...</i>`;
  }

  if (!list.length) {
    return `<i>Không có khách hàng trong tập điểm này</i>`;
  }

  let html = `
    <div style="max-height:200px; overflow:auto; font-size:12px">
      <table border="1" cellpadding="4" cellspacing="0" width="100%">
        <tr style="background:#f0f0f0">
          <th>Tên KH</th>
          <th>Địa chỉ</th>
          <th>Trạng thái</th>
        </tr>
  `;

  list.forEach(c => {
    const isNormal = normalizeText(c.status) === normalizeText('Hoạt động bình thường');
    const rowKey = escapeHtml(c.rowKey || "");
  
    html += `
      <tr data-customer-key="${rowKey}" data-default-bg="${isNormal ? '' : '#ffe6e6'}" data-default-weight="" style="${isNormal ? '' : 'background:#ffe6e6; color:#b30000;'}">
        <td>${escapeHtml(c.name || '')}</td>
        <td>${escapeHtml(c.address || '')}</td>
        <td>${escapeHtml(c.status || '')}</td>
      </tr>
    `;
  });

  html += `</table></div>`;
  return html;
}

// Hiển thị danh sách khách hàng trong popup tập điểm
function showCustomer(pointId) {
  const point = allPoints.find(item => item.id === pointId);
  if (!point) return;

  const el = document.getElementById(point.customerContainerId);
  if (!el) return;
  el.innerHTML = renderCustomerList(point);
}

function updateExportButtons() {
  const disabled = !isPointDataReady || allPoints.length === 0;
  if (exportUpgradeBtn) exportUpgradeBtn.disabled = disabled;
  if (exportKmlBtn) exportKmlBtn.disabled = disabled;
}

function updateUnmatchedExportButton() {
  if (!exportUnmatchedCustomersBtn) return;
  exportUnmatchedCustomersBtn.disabled = !isCustomerDataReady || unmatchedCustomers.length === 0;
}

function highlightCustomerInPopup(point, customer) {
  showCustomer(point.id);

  const container = document.getElementById(point.customerContainerId);
  if (!container) return;

  const rows = container.querySelectorAll("table tr");
  rows.forEach(row => {
    if (row.querySelector("th")) return;
    row.style.background = row.dataset.defaultBg || "";
    row.style.fontWeight = row.dataset.defaultWeight || "";
  });

  const targetRow = Array.from(rows).find(row => row.dataset.customerKey === customer.rowKey);
  if (!targetRow) return;

  targetRow.style.background = "#fff3cd";
  targetRow.style.fontWeight = "bold";
  targetRow.scrollIntoView({ block: "nearest" });
}

/* Load dữ liệu Google Sheet theo header thay vì số thứ tự cột. */
async function initData() {
  customerInput.disabled = true;
  customerInput.placeholder = "Đang tải dữ liệu khách hàng...";
  updateExportButtons();
  updateUnmatchedExportButton();

  await loadData();
  if (!isPointDataReady) {
    customerInput.placeholder = "Không tải được dữ liệu tập điểm";
    return;
  }

  await loadCustomerData();
}

async function loadData() {
  try {
    isPointDataReady = false;
    allPoints.forEach(point => {
      if (point.circle && map.hasLayer(point.circle)) {
        map.removeLayer(point.circle);
      }
    });
    suggestionLayers.forEach(layer => map.removeLayer(layer));
    suggestionLayers = [];
    if (customerMarker && map.hasLayer(customerMarker)) {
      map.removeLayer(customerMarker);
    }
    customerMarker = null;
    if (clickMarker && map.hasLayer(clickMarker)) {
      map.removeLayer(clickMarker);
    }
    clickMarker = null;
    customerLatLng = null;
    allPoints = [];
    pointByName = new Map();
    normalizedPointBuckets = new Map();

    const rows = await fetchSheetRows({
      gid: CONFIG.SHEET_TAPDIEM_GID,
      sheetName: CONFIG.SHEET_TAPDIEM,
      preferExport: true,
      requiredHeaders: [
        ["Tên tập điểm S2", "Ten tap diem S2"],
        ["Dung lượng S2", "Dung luong S2"],
        ["Port đã dùng", "Port da dung"],
        ["Port còn lại", "Port con lai"],
        ["Hiệu suất sử dụng", "Hieu suat su dung"],
        ["Lat"],
        ["Long", "Lng"]
      ]
    });

    rows.forEach(row => {
      const name = getField(row, ["Tên tập điểm S2", "Ten tap diem S2"]);
      if (!name || normalizeText(name) === normalizeText("Tổng")) return;

      const lat = toNumber(getField(row, "Lat"));
      const lng = toNumber(getField(row, ["Long", "Lng"]));
      if (!validateLatLng(lat, lng)) return;

      const eff = toPercentRatio(getField(row, ["Hiệu suất sử dụng", "Hieu suat su dung"]));
      if (!Number.isFinite(eff)) return;

      const maxPort = getField(row, ["Dung lượng S2", "Dung luong S2"]) || "0";
      const usedPort = getField(row, ["Port đã dùng", "Port da dung"]) || "0";
      const freePort = getField(row, ["Port còn lại", "Port còn lại ", "Port con lai"]) || "0";
      const freePortValue = toNumber(freePort);

      const point = {
        id: makeSafeId("point", name),
        key: canonicalName(name),
        nameKey: canonicalName(name),
        name,
        maxPort,
        usedPort,
        freePort,
        freePortValue,
        eff,
        lat,
        lng,
        customerContainerId: makeSafeId("customer", name)
      };

      point.circle = L.circle([lat, lng], {
        radius: CONFIG.POINT_RADIUS_METERS,
        color: getColor(eff),
        fillColor: getColor(eff),
        fillOpacity: 0.5,
        weight: 2
      });

      point.basePopup = buildBasePopup(point);
      point.circle.bindPopup(point.basePopup);

      point.circle.on("popupopen", e => {
        const popupEl = e.popup.getElement();
        if (!popupEl) return;

        const holder = popupEl.querySelector(".direction-holder");
        if (holder) {
          holder.innerHTML = customerLatLng
            ? renderDirection(holder.dataset.lat, holder.dataset.lng)
            : `<i style="color:#999">Chưa lấy được vị trí khách hàng</i>`;
        }

        const trigger = popupEl.querySelector(".show-customer-link");
        if (trigger) {
          trigger.onclick = event => {
            event.preventDefault();
            showCustomer(trigger.dataset.pointId);
          };
        }
      });

      registerPoint(point);
      allPoints.push(point);
    });

    renderPoints();
    updateEffCounts();
    isPointDataReady = true;
    updateExportButtons();
  } catch (error) {
    updateExportButtons();
    alert("Không đọc được dữ liệu Google Sheet (Danh sách tập điểm)");
    console.error("Point sheet error:", error);
  }
}

async function loadCustomerData() {
  try {
    customerByTapDiem = {};
    customerSearchList = [];
    unmatchedCustomers = [];
    const rows = await fetchSheetRows({
      gid: CONFIG.SHEET_CUSTOMER_GID,
      sheetName: CONFIG.SHEET_CUSTOMER,
      preferExport: true,
      requiredHeaders: [
        ["Tên khách hàng", "Tên KHG", "Ten khach hang", "Ten KHG"],
        ["Tap Diem/Tram Ket noi", "Tập điểm/Trạm Kết nối"],
        ["Tình Trạng HĐ", "Tinh Trang HD"]
      ]
    });

    rows.forEach(row => {
      const name = getField(row, ["Tên khách hàng", "Tên KHG", "Ten khach hang", "Ten KHG"]);
      const tapDiem = getField(row, ["Tap Diem/Tram Ket noi", "Tập điểm/Trạm Kết nối"]);
      const address = getField(row, ["Địa chỉ khách hàng", "Địa chỉ", "Dia chi khach hang", "Dia chi"]);
      const status = getField(row, ["Tình Trạng HĐ", "Tinh Trang HD"]);
      const hasCustomerInfo = Boolean(canonicalName(name) || canonicalName(address));

      if (!hasCustomerInfo) return;

      if (!tapDiem) {
        unmatchedCustomers.push({
          name,
          tapDiem: "Trống tập điểm",
          address,
          reason: "Thiếu tên tập điểm trên sheet khách hàng"
        });
        return;
      }

      const point = resolvePointByTapDiemName(tapDiem);
      if (!point) {
        unmatchedCustomers.push({
          name,
          tapDiem,
          address,
          reason: "Không tìm thấy tập điểm tương ứng trên sheet tập điểm"
        });
        return;
      }

      if (!customerByTapDiem[point.id]) {
        customerByTapDiem[point.id] = [];
      }

      const customer = {
        name,
        address,
        tapDiem,
        pointId: point.id,
        status,
        rowKey: canonicalName([name, address, status, tapDiem].join("|"))
      };

      customerByTapDiem[point.id].push(customer);
      customerSearchList.push(customer);
    });

    isCustomerDataReady = true;
    customerInput.disabled = false;
    customerInput.placeholder = "Tìm KH đã bán (gõ tên - địa chỉ)";
    updateUnmatchedExportButton();
  } catch (error) {
    customerInput.placeholder = "Không tải được dữ liệu KH";
    customerInput.disabled = true;
    customerByTapDiem = {};
    customerSearchList = [];
    unmatchedCustomers = [];
    isCustomerDataReady = false;
    updateUnmatchedExportButton();
    alert("Không đọc được dữ liệu Google Sheet (Danh sách khách hàng)");
    console.error("Customer sheet error:", error);
  }
}

// Lấy danh sách tên tập điểm đang được hiển thị trên map (phục vụ lọc KH)
function getVisibleTapDiemNames() {
  return allPoints
    .filter(p => map.hasLayer(p.circle))
    .map(p => p.id);
}
  
/* =====================================================
   HIỂN THỊ TẬP ĐIỂM
   ===================================================== */

function renderPoints() {
  const bounds = [];

  allPoints.forEach(p => {
    p.circle.addTo(map);
    bounds.push([p.lat, p.lng]);
  });

  // Zoom toàn bộ tập điểm ở lần load đầu
  if (isInitialLoad && bounds.length) {
    map.fitBounds(bounds, { padding: [40, 40] });
    isInitialLoad = false;
  }
}

/* =====================================================
   ĐẾM SỐ LƯỢNG TẬP ĐIỂM
   ===================================================== */
function updateEffCounts() {
  const counts = {
    lt20: 0,
    '20_50': 0,
    '50_80': 0,
    '80_100': 0,
    gte100: 0
  };

  allPoints.forEach(p => {
    if (p.eff < 0.2) counts.lt20++;
    else if (p.eff < 0.5) counts['20_50']++;
    else if (p.eff < 0.8) counts['50_80']++;
    else if (p.eff < 1) counts['80_100']++;
    else counts.gte100++;
  });

  // CẬP NHẬT SỐ LƯỢNG TỪNG NHÓM
  document.getElementById("cnt-lt20").textContent     = counts.lt20;
  document.getElementById("cnt-20_50").textContent   = counts['20_50'];
  document.getElementById("cnt-50_80").textContent   = counts['50_80'];
  document.getElementById("cnt-80_100").textContent  = counts['80_100'];
  document.getElementById("cnt-gte100").textContent  = counts.gte100;

  // TỔNG SỐ TẬP ĐIỂM
  document.getElementById("cnt-total").textContent = allPoints.length;
}

/* =====================================================
   FILTER THEO HIỆU SUẤT
   ===================================================== */

function applyFilter() {
  // Lấy tất cả checkbox đang được tick
  const checked = Array.from(
      document.querySelectorAll('.eff-filter:checked')
    ).map(cb => cb.value);
  
    allPoints.forEach(p => {
      let show = false;
  
      checked.forEach(v => {
        if (v === 'lt20'     && p.eff < 0.2) show = true;
        if (v === '20_50'    && p.eff >= 0.2 && p.eff < 0.5) show = true;
        if (v === '50_80'    && p.eff >= 0.5 && p.eff < 0.8) show = true;
        if (v === '80_100'   && p.eff >= 0.8 && p.eff < 1)   show = true;
        if (v === 'gte100'   && p.eff >= 1)  show = true;
      });
  
      show ? p.circle.addTo(map) : map.removeLayer(p.circle);
    });
}

/* =====================================================
   TÌM KIẾM KHÁCH HÀNG
   ===================================================== */

async function searchLocation() {
  if (!isPointDataReady) {
    alert("Dữ liệu tập điểm đang tải, vui lòng thử lại sau.");
    return;
  }

  const val = searchInput.value.trim();
  if (!val) return;

  // Nhập tọa độ trực tiếp
  const p = val.split(',');
  if (p.length === 2 && !isNaN(p[0]) && !isNaN(p[1])) {
    handleCustomer(+p[0], +p[1]);
    return;
  }

  try {
    const results = await geocodeAddress(val, 1);
    if (!results[0]) {
      alert("Không tìm thấy địa chỉ cần tìm.");
      return;
    }

    handleCustomer(+results[0].lat, +results[0].lon);
  } catch (err) {
    console.error("Search location error:", err);
    alert("Không tìm được địa chỉ. Vui lòng kiểm tra kết nối mạng.");
  }

  collapseControlBox();
}

/* =====================================================
   XỬ LÝ LOGIC KHÁCH HÀNG & GỢI Ý TẬP ĐIỂM
   ===================================================== */
/*
Xử lý khi xác định vị trí khách hàng:
- Reset marker & highlight cũ
- Kiểm tra phạm vi bán hàng ≤ 1km
- Chọn tối đa 3 tập điểm gần nhất còn port
- Highlight, vẽ đường và hiển thị gợi ý
*/
function handleCustomer(lat, lng) {
  if (!isPointDataReady) {
    alert("Dữ liệu tập điểm đang tải, vui lòng thử lại sau.");
    return;
  }

  customerLatLng = { lat, lng };
  // Xóa marker & layer cũ
  if (customerMarker) map.removeLayer(customerMarker);
  suggestionLayers.forEach(l => map.removeLayer(l));
  suggestionLayers = [];

  // Reset style tập điểm
  allPoints.forEach(p => {
    p.circle.setStyle({ weight: 2 });
    p.circle.setPopupContent(p.basePopup); // 👈 reset popup
    const el = p.circle.getElement();
    if (el) el.classList.remove('blink');
  });

  // Marker khách hàng
  customerMarker = L.marker([lat, lng]).addTo(map);
  map.setView([lat, lng], 15);
  
  // ===============================
  // Kiểm tra khoảng cách tới tập điểm gần nhất ≤ 1km
  // ===============================
  const candidates = allPoints
    .filter(p => Number.isFinite(p.freePortValue) && p.freePortValue > 0)
    .map(p => ({ ...p, dist: haversine(lat, lng, p.lat, p.lng) }))
    .sort((a, b) => a.dist - b.dist || a.eff - b.eff);

  const nearestPoint = candidates
    .sort((a, b) => a.dist - b.dist)[0];
  
  if (!nearestPoint || nearestPoint.dist > CONFIG.MAX_SALE_DISTANCE_METERS) {
    alert(
      `Không khả thi để bán hàng\n` +
      `Tập điểm còn port gần nhất cách ${Math.round(nearestPoint?.dist || 0)} m`
    );
    return;
  }

  // Chọn 3 tập điểm gần nhất còn port trống
  const topCandidates = candidates.slice(0, CONFIG.MAX_SUGGESTED_POINTS);

  topCandidates.forEach((p, i) => {
    const el = p.circle.getElement();
    if (el) el.classList.add('blink');

    p.circle.setStyle({ weight: 4 });

    // Vẽ đường nối khách hàng → tập điểm
    const line = L.polyline(
      [[lat, lng], [p.lat, p.lng]],
      { color: i === 0 ? 'red' : 'green', dashArray: '5,5' }
    ).addTo(map);

    // Hiển thị khoảng cách
    line.bindTooltip(`${Math.round(p.dist)} m`, {
      permanent: true,
      direction: 'center'
    });

    // Popup chi tiết tập điểm gợi ý quanh vị trí khách hàng
    const bestBadge =
      i === 0
        ? '<b style="color:green">Tập điểm gợi ý tốt nhất</b>'
        : '';
    
    const suggestPopup = `
      ${p.basePopup}
      <hr>
      <b>Khoảng cách:</b> ${Math.round(p.dist)} m<br>
      ${bestBadge}
    `;
    
    p.circle.setPopupContent(suggestPopup);

    suggestionLayers.push(line);
  });
}

/* =====================================================
   EXPORT DỮ LIỆU
   ===================================================== */

/* Export danh sách tập điểm cần nâng cấp ra Excel */
function exportUpgradeList() {
  if (!isPointDataReady || !allPoints.length) {
    alert("Dữ liệu tập điểm chưa tải xong. Vui lòng thử lại sau.");
    return;
  }

  const rows = [[
    'Tên tập điểm',
    'Lat',
    'Long',
    'Dung lượng (Port)',
    'Port đã dùng',
    'Port còn lại',
    'Hiệu suất (%)',
    'Mức ưu tiên'
  ]];

  allPoints.forEach(p => {
    if (p.eff >= 0.8) {
      rows.push([
        p.name,
        p.lat,
        p.lng,
        p.maxPort,
        p.usedPort,
        p.freePort,
        (p.eff * 100).toFixed(1),
        p.eff >= 1
          ? 'Ưu tiên 1 (Đầy tải)'
          : 'Ưu tiên 2 (Gần đầy tải)'
      ]);
    }
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Upgrade');
  XLSX.writeFile(wb, 'DS_tap_diem_can_nang_cap.xlsx');
}

/* Export toàn bộ tập điểm ra file KML
   (chuyển circle Leaflet thành polygon để tương thích Google Earth) */
function exportAllCirclesToKML() {
  if (!isPointDataReady || !allPoints.length) {
    alert("Dữ liệu tập điểm chưa tải xong. Vui lòng thử lại sau.");
    return;
  }

  let placemarks = "";

  allPoints.forEach(p => {
    const polygon = circleToPolygon(p.lat, p.lng, CONFIG.POINT_RADIUS_METERS);
    const color = colorToKML(getColor(p.eff));

    const description = `
      <![CDATA[
        <b>${escapeHtml(p.name)}</b><br>
        Dung lượng (Port): ${escapeHtml(p.maxPort)}<br>
        Đã dùng: ${escapeHtml(p.usedPort)}<br>
        Còn lại: ${escapeHtml(p.freePort)}<br>
        Hiệu suất: ${(p.eff * 100).toFixed(1)}%
      ]]>
    `;

    placemarks += `
      <Placemark>
        <name>${escapeXml(p.name)}</name>
        <description>${description}</description>
        <Style>
          <PolyStyle>
            <color>${color}</color>
            <outline>1</outline>
          </PolyStyle>
          <LineStyle>
            <color>ff854442</color>
            <width>2</width>
          </LineStyle>
        </Style>
        <Polygon>
          <outerBoundaryIs>
            <LinearRing>
              <coordinates>
                ${polygon}
              </coordinates>
            </LinearRing>
          </outerBoundaryIs>
        </Polygon>
      </Placemark>
    `;
  });

  const kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
<Document>
  <name>Vùng phủ GPON MobiFiber</name>
  ${placemarks}
</Document>
</kml>`;

  downloadKML(kml, "Vùng_phủ_GPON_MobiFiber.kml");
}

function exportUnmatchedCustomers() {
  if (!isCustomerDataReady) {
    alert("Dữ liệu khách hàng chưa tải xong. Vui lòng thử lại sau.");
    return;
  }

  if (!unmatchedCustomers.length) {
    alert("Không có khách hàng nào chưa ghép được với tập điểm.");
    return;
  }

  const rows = [[
    "Tên khách hàng",
    "Tập điểm trên sheet khách hàng",
    "Địa chỉ",
    "Lý do"
  ]];

  unmatchedCustomers.forEach(item => {
    rows.push([
      item.name || "",
      item.tapDiem || "",
      item.address || "",
      item.reason || ""
    ]);
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "KhachHangChuaGhep");
  XLSX.writeFile(wb, "DS_khach_hang_chua_ghep_tap_diem.xlsx");
}

/* Tải file KML */
function downloadKML(content, filename) {
  const blob = new Blob([content], {
    type: "application/vnd.google-earth.kml+xml"
  });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

/* =====================================================
   AUTOCOMPLETE & ĐỊNH VỊ
   ===================================================== */

const input = document.getElementById("searchInput");
const suggestBox = document.getElementById("suggestBox");
const customerInput = document.getElementById("customerSearchInput");
const customerSuggestBox = document.getElementById("customerSuggestBox");
const exportUpgradeBtn = document.getElementById("exportUpgradeBtn");
const exportKmlBtn = document.getElementById("exportKmlBtn");
const exportUnmatchedCustomersBtn = document.getElementById("exportUnmatchedCustomersBtn");
const feedbackBtn = document.getElementById("feedbackBtn");
// Timer debounce autocomplete để tránh gọi API liên tục
let suggestTimer = null;

if (feedbackBtn) {
  feedbackBtn.addEventListener("click", () => {
    if (!CONFIG.FEEDBACK_FORM_URL) {
      alert("Chưa cấu hình link phản ánh sự cố.");
      return;
    }
    window.open(CONFIG.FEEDBACK_FORM_URL, "_blank");
  });
}

initData();

/* Autocomplete địa chỉ */
input.addEventListener("input", () => {
  const q = input.value.trim();
  
  if (!q) {
    customerLatLng = null;         
    suggestBox.style.display = "none";
    return;
  }
  
  if (q.length < 3) {
    suggestBox.style.display = "none";
    return;
  }

  clearTimeout(suggestTimer);
  suggestTimer = setTimeout(async () => {
    try {
      const data = await geocodeAddress(q, 5);
      suggestBox.innerHTML = "";
      if (!data.length) {
        suggestBox.style.display = "none";
        return;
      }

      data.forEach(item => {
        const div = document.createElement("div");
        div.className = "suggest-item";
        div.textContent = item.display_name;

        div.onclick = () => {
          input.value = item.display_name;
          suggestBox.style.display = "none";
          handleCustomer(+item.lat, +item.lon);
        };

        suggestBox.appendChild(div);
      });

      suggestBox.style.display = "block";
    } catch (err) {
      console.error("Suggest error:", err);
      suggestBox.style.display = "none";
    }
  }, 300);
});

customerInput.addEventListener("input", () => {
  if (!isCustomerDataReady) {
    customerSuggestBox.style.display = "none";
    return;
  }

  const q = normalizeText(customerInput.value.trim());
  customerSuggestBox.innerHTML = "";

  if (q.length < 2) {
    customerSuggestBox.style.display = "none";
    return;
  }

  const visibleTapDiems = getVisibleTapDiemNames();

  const matches = customerSearchList
    .filter(c =>
      visibleTapDiems.includes(c.pointId) &&
      (
        (c.name && normalizeText(c.name).includes(q)) ||
        (c.address && normalizeText(c.address).includes(q))
      )
    )
    .slice(0, 10);

  if (!matches.length) {
    customerSuggestBox.style.display = "none";
    return;
  }

  matches.forEach(c => {
    const div = document.createElement("div");
    div.className = "suggest-item";
    const strong = document.createElement("strong");
    strong.textContent = c.name || "(Không có tên)";
    div.appendChild(strong);
    div.appendChild(document.createTextNode(` - ${c.address || ''}`));

    div.onclick = () => {
      customerInput.value = `${c.name} - ${c.address || ''}`;
      customerSuggestBox.style.display = "none";
      locateCustomerInTapDiem(c);
      collapseControlBox();
    };

    customerSuggestBox.appendChild(div);
  });

  customerSuggestBox.style.display = "block";
});

/* Ẩn gợi ý khi click ngoài */
document.addEventListener("click", e => {
  if (!e.target.closest(".search-row")) {
    suggestBox.style.display = "none";
    customerSuggestBox.style.display = "none";
  }
});

/* Định vị người dùng */
let myLocationMarker = null;
// Vòng tròn thể hiện sai số định vị (accuracy)
let myAccuracyCircle = null;

function locateMe() {
  if (!navigator.geolocation) {
    alert("Trình duyệt không hỗ trợ định vị");
    return;
  }

  navigator.geolocation.getCurrentPosition(
    pos => {
      const lat = pos.coords.latitude;
      const lng = pos.coords.longitude;
      const accuracy = pos.coords.accuracy;

      if (myLocationMarker) map.removeLayer(myLocationMarker);
      if (myAccuracyCircle) map.removeLayer(myAccuracyCircle);

      myLocationMarker = L.marker([lat, lng])
        .addTo(map)
        .bindPopup("📍 Vị trí của tôi")
        .openPopup();

      myAccuracyCircle = L.circle([lat, lng], {
        radius: accuracy,
        color: "blue",
        fillColor: "blue",
        fillOpacity: 0.1
      }).addTo(map);

      map.setView([lat, lng], 15);
      handleCustomer(lat, lng);
      collapseControlBox();
    },
    err => alert("Không lấy được vị trí: " + err.message),
    { enableHighAccuracy: true, timeout: 10000 }
  );
}

/* =====================================================
   THU GỌN / MỞ RỘNG PANEL
   ===================================================== */

function toggleControl() {
  const box = document.getElementById("controlBox");
  const arrow = document.getElementById("arrow");

  box.classList.toggle("collapsed");
  arrow.textContent = box.classList.contains("collapsed") ? "▼" : "▲";
}

function toggleFilter() {
  const content = document.getElementById("filterContent");
  const arrow = document.getElementById("filterArrow");

  const isOpen = content.style.display === "block";

  content.style.display = isOpen ? "none" : "block";
  arrow.textContent = isOpen ? "▼" : "▲";
}

function collapseControlBox() {
  const box = document.getElementById("controlBox");
  const arrow = document.getElementById("arrow");

  if (!box.classList.contains("collapsed")) {
    box.classList.add("collapsed");
    arrow.textContent = "▼";
  }
}

  // Gắn sự kiện thay đổi filter hiệu suất để cập nhật hiển thị tập điểm
  document.querySelectorAll('.eff-filter')
  .forEach(cb => cb.addEventListener('change', applyFilter));

// Xác định và zoom tới tập điểm mà khách hàng đang đấu nối
// Highlight tập điểm và khách hàng tương ứng trong danh sách
function locateCustomerInTapDiem(customer) {
  // customerLatLng = null;
  const p = allPoints.find(tp => tp.id === customer.pointId);
  if (!p) {
    alert("Không tìm thấy tập điểm của khách hàng");
    return;
  }

  map.setView([p.lat, p.lng], 16);

  allPoints.forEach(x => {
    x.circle.setStyle({ weight: 2 });
    const el = x.circle.getElement();
    if (el) el.classList.remove("blink");
  });

  p.circle.setStyle({ weight: 4 });
  const el = p.circle.getElement();
  if (el) el.classList.add("blink");

  if (p.circle.isPopupOpen && p.circle.isPopupOpen()) {
    highlightCustomerInPopup(p, customer);
    return;
  }

  p.circle.once("popupopen", () => {
    highlightCustomerInPopup(p, customer);
  });

  p.circle.openPopup();
}

// Tạo link Google Maps dẫn đường từ vị trí khách hàng tới tập điểm
function renderDirection(lat, lng) {
  if (!customerLatLng) {
    return `<i style="color:#999">Chưa lấy được vị trí khách hàng</i>`;
  }

  const url =
    `https://www.google.com/maps/dir/?api=1` +
    `&origin=${customerLatLng.lat},${customerLatLng.lng}` +
    `&destination=${lat},${lng}`;

  return `
    <a href="${url}" target="_blank" rel="noopener">
      🧭 đi đến tập điểm
    </a>
  `;

}
