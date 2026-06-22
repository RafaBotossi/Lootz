const DEFAULT_FILE = "inventario_venda_rpg_ordenado.xlsx";
const STORAGE_KEY = "lootz-inventory-v1";

const state = {
  items: [],
  members: 3,
  source: DEFAULT_FILE,
  search: "",
  category: "",
  sale: "",
  sort: "original",
  visible: []
};

const el = id => document.getElementById(id);
const money = value => `${new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value || 0)} gp`;
const number = value => Number(String(value ?? 0).replace(",", ".")) || 0;
const normalized = value => String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
const escapeHtml = value => String(value ?? "").replace(/[&<>"']/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" })[char]);
const isCashItem = item => ["moedas", "recompensa"].includes(normalized(item.categoria));

function saleValue(item) {
  return item.vender && !isCashItem(item)
    ? item.valorBase * item.quantidade * (1 + item.bonus)
    : 0;
}

function rowValue(row, names, fallback = "") {
  const entries = Object.entries(row);
  for (const wanted of names) {
    const found = entries.find(([key]) => normalized(key) === normalized(wanted));
    if (found && found[1] !== undefined && found[1] !== null) return found[1];
  }
  return fallback;
}

function rowsToItems(rows) {
  return rows
    .filter(row => Object.values(row).some(value => value !== null && value !== ""))
    .map((row, index) => {
      const real = number(rowValue(row, ["Valor real (gp)", "Valor real"]));
      const estimated = number(rowValue(row, ["Valor estimado (gp)", "Valor estimado"]));
      const base = number(rowValue(row, ["Valor base (gp)", "Valor base"], real || estimated));
      return {
        id: `${Date.now()}-${index}`,
        order: index,
        origem: rowValue(row, ["Origem"], "Sem origem"),
        item: rowValue(row, ["Item", "Nome"], "Item sem nome"),
        quantidade: Math.max(1, number(rowValue(row, ["Qtd.", "Qtd", "Quantidade"], 1))),
        valorReal: real,
        valorEstimado: estimated,
        estimativa: rowValue(row, ["Como foi estimado", "Estimativa"]),
        categoria: rowValue(row, ["Categoria"], "Outros"),
        vender: ["sim", "s", "true", "1", "yes"].includes(normalized(rowValue(row, ["Vender?", "Vender"]))),
        valorBase: base,
        bonus: number(rowValue(row, ["Bônus aplicado", "Bonus aplicado", "Bônus", "Bonus"])),
        observacao: rowValue(row, ["Observação", "Observacao", "Notas"])
      };
    });
}

function parseWorkbook(data, source, markImported = false) {
  if (!window.XLSX) throw new Error("O leitor de planilhas não carregou. Verifique sua conexão e tente novamente.");
  const workbook = XLSX.read(data, { type: "array" });
  const sheetName = workbook.SheetNames.find(name => normalized(name).includes("inventario")) || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  const items = rowsToItems(rows);
  if (!items.length) throw new Error("Não encontrei itens nessa planilha.");
  if (markImported) {
    items.forEach(item => {
      if (!isCashItem(item)) item.vender = true;
    });
  }

  const parameterSheet = workbook.SheetNames.find(name => normalized(name).includes("parametro"));
  let members = 3;
  if (parameterSheet) {
    const parameterRows = XLSX.utils.sheet_to_json(workbook.Sheets[parameterSheet], { header: 1, defval: "" });
    const memberRow = parameterRows.find(row => normalized(row[0]).includes("divisor"));
    if (memberRow) members = Math.max(1, number(memberRow[1]));
  }

  state.items = items;
  state.members = members;
  state.source = source;
  saveState();
  refresh();
  showToast(`${items.length} itens chegaram ao balcão.`);
}

async function loadDefault(force = false) {
  if (!force) {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed.items) && parsed.items.length) {
          Object.assign(state, parsed);
          refresh();
          return;
        }
      } catch (_) {}
    }
  }

  try {
    const response = await fetch(DEFAULT_FILE);
    if (!response.ok) throw new Error("Planilha padrão não encontrada.");
    parseWorkbook(await response.arrayBuffer(), DEFAULT_FILE);
  } catch (error) {
    el("sourceName").textContent = "Importe uma planilha para começar";
    showToast(error.message);
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify({
    items: state.items,
    members: state.members,
    source: state.source
  }));
}

function getVisibleItems() {
  const query = normalized(state.search);
  const filtered = state.items.filter(item => {
    const haystack = normalized(`${item.item} ${item.origem} ${item.observacao} ${item.estimativa}`);
    const searchMatch = !query || haystack.includes(query);
    const categoryMatch = !state.category || item.categoria === state.category;
    const saleMatch = !state.sale || (state.sale === "sell" ? item.vender : !item.vender);
    return searchMatch && categoryMatch && saleMatch;
  });

  return [...filtered].sort((a, b) => {
    if (state.sort === "name-asc") return a.item.localeCompare(b.item, "pt-BR");
    if (state.sort === "value-desc") return saleValue(b) - saleValue(a) || b.valorBase - a.valorBase;
    if (state.sort === "value-asc") return (saleValue(a) || a.valorBase) - (saleValue(b) || b.valorBase);
    if (state.sort === "origin") return a.origem.localeCompare(b.origem, "pt-BR");
    return a.order - b.order;
  });
}

function renderTable() {
  state.visible = getVisibleItems();
  el("inventoryBody").innerHTML = state.visible.map((item, index) => `
    <tr class="${item.vender ? "" : "kept"}" style="animation-delay:${Math.min(index * 18, 300)}ms">
      <td class="sell-column">
        <label class="sale-toggle" title="${item.vender ? "Remover da venda" : "Marcar para venda"}">
          <input type="checkbox" data-id="${escapeHtml(item.id)}" ${item.vender ? "checked" : ""}>
          <span></span>
        </label>
      </td>
      <td>
        <span class="item-name">${escapeHtml(item.item)}</span>
        ${item.observacao ? `<span class="item-note">${escapeHtml(item.observacao)}</span>` : ""}
      </td>
      <td class="origin">${escapeHtml(item.origem)}</td>
      <td><span class="badge">${escapeHtml(item.categoria)}</span></td>
      <td class="number-column">${new Intl.NumberFormat("pt-BR").format(item.quantidade)}</td>
      <td class="number-column">${money(item.valorBase * item.quantidade)}</td>
      <td class="number-column">${new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 1 }).format(item.bonus)}</td>
      <td class="number-column sale-value">${item.vender ? money(saleValue(item)) : "—"}</td>
    </tr>
  `).join("");

  el("visibleCount").textContent = state.visible.length;
  el("emptyState").hidden = state.visible.length > 0;
  document.querySelector(".table-scroll").hidden = state.visible.length === 0;
}

function renderSummary() {
  const cashItems = state.items.filter(isCashItem);
  const platinum = cashItems.reduce((sum, item) => {
    const match = item.item.match(/(?:^|\s)([\d.,]+)\s*pp\b/i);
    return sum + (match ? number(match[1]) * item.quantidade : 0);
  }, 0);
  const platinumInGold = platinum / 10;
  const cash = cashItems.reduce((sum, item) => sum + item.valorBase * item.quantidade, 0);
  const gold = cash - platinumInGold;
  const sales = state.items.reduce((sum, item) => sum + saleValue(item), 0);
  const kept = state.items
    .filter(item => !item.vender && !isCashItem(item))
    .reduce((sum, item) => sum + item.valorBase * item.quantidade, 0);
  const total = cash + sales;

  el("cashGpTotal").textContent = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 2 }).format(gold);
  el("cashPpTotal").textContent = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 2 }).format(platinum);
  el("cashConverted").textContent = `Total convertido: ${money(cash)}`;
  el("saleTotal").textContent = money(sales);
  el("grandTotal").textContent = money(total);
  el("splitTotal").textContent = money(total / Math.max(1, state.members));
  el("selectedCount").textContent = state.items.filter(item => item.vender).length;
  el("memberCount").value = state.members;
  el("itemCount").textContent = state.items.length;
  el("itemCountHero").textContent = state.items.length;
  el("sourceName").textContent = state.source;
}

function renderCategories() {
  const selected = state.category;
  const categories = [...new Set(state.items.map(item => item.categoria))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  el("categoryFilter").innerHTML = `<option value="">Todas as categorias</option>` +
    categories.map(category => `<option value="${escapeHtml(category)}">${escapeHtml(category)}</option>`).join("");
  el("categoryFilter").value = selected;
}

function refresh() {
  renderCategories();
  renderTable();
  renderSummary();
}

function setVisibleSale(value) {
  const ids = new Set(state.visible.map(item => item.id));
  state.items.forEach(item => {
    if (ids.has(item.id) && !isCashItem(item)) item.vender = value;
  });
  saveState();
  refresh();
  showToast(value ? "Itens visíveis marcados para venda." : "Itens visíveis voltaram para o baú.");
}

function exportCsv() {
  const headers = ["Origem", "Item", "Qtd.", "Categoria", "Vender?", "Valor base (gp)", "Bônus aplicado", "Venda líquida (gp)", "Observação"];
  const rows = state.visible.map(item => [
    item.origem, item.item, item.quantidade, item.categoria, item.vender ? "Sim" : "Não",
    item.valorBase, item.bonus, saleValue(item), item.observacao
  ]);
  const csv = [headers, ...rows]
    .map(row => row.map(value => `"${String(value ?? "").replaceAll('"', '""')}"`).join(";"))
    .join("\r\n");
  const blob = new Blob(["\ufeff", csv], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "lootz-selecao.csv";
  link.click();
  URL.revokeObjectURL(link.href);
  showToast("A seleção foi selada em CSV.");
}

let toastTimer;
function showToast(message) {
  clearTimeout(toastTimer);
  el("toast").textContent = message;
  el("toast").classList.add("show");
  toastTimer = setTimeout(() => el("toast").classList.remove("show"), 3200);
}

function setupEmbers() {
  if (matchMedia("(prefers-reduced-motion: reduce)").matches) return;
  el("embers").innerHTML = Array.from({ length: 18 }, () => {
    const left = Math.random() * 100;
    const duration = 7 + Math.random() * 10;
    const delay = Math.random() * -15;
    const drift = `${-50 + Math.random() * 100}px`;
    return `<i class="ember" style="left:${left}%;--duration:${duration}s;--delay:${delay}s;--drift:${drift}"></i>`;
  }).join("");
}

let audioContext;
let ambienceNodes = [];
let ambienceTimers = [];
let soundEnabled = true;
let soundStarted = false;

function makeNoise(seconds = 5) {
  const buffer = audioContext.createBuffer(2, audioContext.sampleRate * seconds, audioContext.sampleRate);
  for (let channelIndex = 0; channelIndex < 2; channelIndex++) {
    const channel = buffer.getChannelData(channelIndex);
    let brown = 0;
    for (let i = 0; i < channel.length; i++) {
      brown = (brown + 0.018 * (Math.random() * 2 - 1)) / 1.018;
      channel[i] = brown * 3.2;
    }
  }
  return buffer;
}

function randomMurmur(master) {
  if (!soundStarted) return;
  const now = audioContext.currentTime;
  const duration = 1.7 + Math.random() * 2.8;
  const voice = audioContext.createOscillator();
  const formant = audioContext.createBiquadFilter();
  const gain = audioContext.createGain();
  const pan = audioContext.createStereoPanner();
  voice.type = "sawtooth";
  voice.frequency.setValueAtTime(72 + Math.random() * 55, now);
  voice.frequency.linearRampToValueAtTime(65 + Math.random() * 70, now + duration);
  formant.type = "bandpass";
  formant.frequency.setValueAtTime(420 + Math.random() * 600, now);
  formant.Q.value = 1.2;
  pan.pan.value = Math.random() * 1.6 - .8;
  gain.gain.setValueAtTime(0, now);
  gain.gain.linearRampToValueAtTime(.008 + Math.random() * .009, now + .4);
  gain.gain.setValueAtTime(.007, now + duration - .5);
  gain.gain.linearRampToValueAtTime(0, now + duration);
  voice.connect(formant).connect(gain).connect(pan).connect(master);
  voice.start(now);
  voice.stop(now + duration);
  ambienceTimers.push(setTimeout(() => randomMurmur(master), 500 + Math.random() * 1800));
}

function randomClink(master) {
  if (!soundStarted) return;
  const now = audioContext.currentTime;
  [0, .045].forEach((offset, index) => {
    const osc = audioContext.createOscillator();
    const gain = audioContext.createGain();
    const pan = audioContext.createStereoPanner();
    osc.type = "sine";
    osc.frequency.setValueAtTime((index ? 2380 : 1740) + Math.random() * 350, now + offset);
    osc.frequency.exponentialRampToValueAtTime(900, now + offset + .18);
    gain.gain.setValueAtTime(.025, now + offset);
    gain.gain.exponentialRampToValueAtTime(.0001, now + offset + .22);
    pan.pan.value = Math.random() * 1.5 - .75;
    osc.connect(gain).connect(pan).connect(master);
    osc.start(now + offset);
    osc.stop(now + offset + .24);
  });
  ambienceTimers.push(setTimeout(() => randomClink(master), 5000 + Math.random() * 11000));
}

function randomLute(master) {
  if (!soundStarted) return;
  const notes = [146.83, 174.61, 196, 220, 246.94, 293.66];
  const now = audioContext.currentTime;
  const base = notes[Math.floor(Math.random() * notes.length)];
  [1, 1.5, 2].forEach((ratio, index) => {
    const osc = audioContext.createOscillator();
    const gain = audioContext.createGain();
    const pan = audioContext.createStereoPanner();
    osc.type = index ? "sine" : "triangle";
    osc.frequency.value = base * ratio;
    gain.gain.setValueAtTime(.018 / (index + 1), now + index * .11);
    gain.gain.exponentialRampToValueAtTime(.0001, now + 1.8 + index * .15);
    pan.pan.value = -.3;
    osc.connect(gain).connect(pan).connect(master);
    osc.start(now + index * .11);
    osc.stop(now + 2.1);
  });
  ambienceTimers.push(setTimeout(() => randomLute(master), 7000 + Math.random() * 9000));
}

function startAmbience(silent = false) {
  if (!soundEnabled || soundStarted) return;
  audioContext = new (window.AudioContext || window.webkitAudioContext)();
  const master = audioContext.createGain();
  const compressor = audioContext.createDynamicsCompressor();
  master.gain.value = .34;
  master.connect(compressor).connect(audioContext.destination);

  const room = audioContext.createBufferSource();
  const roomFilter = audioContext.createBiquadFilter();
  const roomGain = audioContext.createGain();
  room.buffer = makeNoise(7);
  room.loop = true;
  roomFilter.type = "lowpass";
  roomFilter.frequency.value = 580;
  roomGain.gain.value = .055;
  room.connect(roomFilter).connect(roomGain).connect(master);
  room.start();

  const fire = audioContext.createBufferSource();
  const fireFilter = audioContext.createBiquadFilter();
  const fireGain = audioContext.createGain();
  fire.buffer = makeNoise(4);
  fire.loop = true;
  fireFilter.type = "highpass";
  fireFilter.frequency.value = 1800;
  fireGain.gain.value = .025;
  fire.connect(fireFilter).connect(fireGain).connect(master);
  fire.start();

  const warmth = audioContext.createOscillator();
  const warmthGain = audioContext.createGain();
  warmth.type = "sine";
  warmth.frequency.value = 55;
  warmthGain.gain.value = .012;
  warmth.connect(warmthGain).connect(master);
  warmth.start();

  ambienceNodes = [room, roomFilter, roomGain, fire, fireFilter, fireGain, warmth, warmthGain, master, compressor];
  soundStarted = true;
  randomMurmur(master);
  randomMurmur(master);
  randomClink(master);
  randomLute(master);
  updateSoundButton();
  if (!silent) showToast("A taverna ganhou vida.");
}

function stopAmbience() {
  soundStarted = false;
  ambienceTimers.forEach(clearTimeout);
  ambienceTimers = [];
  ambienceNodes.forEach(node => { try { node.stop?.(); node.disconnect?.(); } catch (_) {} });
  ambienceNodes = [];
  audioContext?.close();
  audioContext = null;
}

function updateSoundButton() {
  const button = el("soundToggle");
  button.setAttribute("aria-pressed", String(soundEnabled));
  button.title = soundEnabled ? "Desativar ambiente da taverna" : "Ativar ambiente da taverna";
  button.querySelector(".sound-label").textContent = soundEnabled ? "Som ligado" : "Som desligado";
}

function toggleAmbience() {
  if (soundEnabled) {
    soundEnabled = false;
    stopAmbience();
    showToast("A taverna ficou em silêncio.");
  } else {
    soundEnabled = true;
    startAmbience();
  }
  updateSoundButton();
}

el("fileInput").addEventListener("change", event => {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try { parseWorkbook(reader.result, file.name, true); }
    catch (error) { showToast(error.message); }
  };
  reader.readAsArrayBuffer(file);
  event.target.value = "";
});

el("inventoryBody").addEventListener("change", event => {
  if (!event.target.matches("input[type='checkbox']")) return;
  const item = state.items.find(candidate => String(candidate.id) === event.target.dataset.id);
  if (!item) return;
  item.vender = event.target.checked;
  saveState();
  refresh();
});

el("searchInput").addEventListener("input", event => { state.search = event.target.value; renderTable(); });
el("categoryFilter").addEventListener("change", event => { state.category = event.target.value; renderTable(); });
el("saleFilter").addEventListener("change", event => { state.sale = event.target.value; renderTable(); });
el("sortSelect").addEventListener("change", event => { state.sort = event.target.value; renderTable(); });
el("exportCsv").addEventListener("click", exportCsv);
el("soundToggle").addEventListener("click", toggleAmbience);
el("memberMinus").addEventListener("click", () => { state.members = Math.max(1, state.members - 1); saveState(); renderSummary(); });
el("memberPlus").addEventListener("click", () => { state.members = Math.min(99, state.members + 1); saveState(); renderSummary(); });
el("memberCount").addEventListener("change", event => { state.members = Math.max(1, Math.min(99, number(event.target.value))); saveState(); renderSummary(); });
setupEmbers();
updateSoundButton();
document.addEventListener("pointerdown", event => {
  if (event.target.closest("#soundToggle")) return;
  startAmbience(true);
}, { once: true, capture: true });
document.addEventListener("keydown", () => startAmbience(true), { once: true });
loadDefault();
