const DEFAULT_FILE = "inventario_venda_rpg_ordenado.xlsx";
const STORAGE_KEY = "lootz-inventory-v1";

const state = {
  items: [],
  members: 3,
  source: "Importe uma planilha ou carregue o exemplo",
  search: "",
  categories: [],
  sale: "",
  sort: "original",
  visible: []
};

const el = id => document.getElementById(id);
const money = value => `${new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value || 0)} gp`;
const number = value => Number(String(value ?? 0).replace(",", ".")) || 0;
const editableNumber = value => Number(Number(value || 0).toFixed(4));
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

async function loadExample() {
  const button = el("loadExample");
  button.disabled = true;
  button.textContent = "Carregando...";
  try {
    const exampleUrl = new URL(`./${DEFAULT_FILE}`, document.baseURI);
    const response = await fetch(exampleUrl, {
      cache: "no-store",
      credentials: "same-origin"
    });
    if (!response.ok) throw new Error("Planilha de exemplo não encontrada.");
    parseWorkbook(await response.arrayBuffer(), DEFAULT_FILE);
  } catch (error) {
    const localFile = location.protocol === "file:";
    showToast(localFile
      ? "Abra o LootZ por um servidor web para carregar o exemplo."
      : `Não foi possível carregar o exemplo: ${error.message}`);
  } finally {
    button.disabled = false;
    button.textContent = "Carregar exemplo";
  }
}

function loadInitialState() {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    try {
      const parsed = JSON.parse(saved);
      if (Array.isArray(parsed.items) && parsed.items.length) {
        Object.assign(state, parsed);
      }
    } catch (_) {}
  }
  refresh();
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
    const categoryMatch = !state.categories.length || state.categories.includes(item.categoria);
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
      <td class="number-column editable-number">
        <label class="inline-value">
          <span class="sr-only">GP base de ${escapeHtml(item.item)}</span>
          <input type="number" min="0" step="0.01" inputmode="decimal"
            data-id="${escapeHtml(item.id)}" data-field="valorBase"
            value="${editableNumber(item.valorBase)}" aria-label="GP base de ${escapeHtml(item.item)}">
          <span>gp</span>
        </label>
      </td>
      <td class="number-column editable-number">
        <label class="inline-value inline-percent">
          <span class="sr-only">Bônus de ${escapeHtml(item.item)}</span>
          <input type="number" step="0.1" inputmode="decimal"
            data-id="${escapeHtml(item.id)}" data-field="bonus"
            value="${editableNumber(item.bonus * 100)}" aria-label="Bônus percentual de ${escapeHtml(item.item)}">
          <span>%</span>
        </label>
      </td>
      <td class="number-column sale-value">${item.vender ? money(saleValue(item)) : "—"}</td>
    </tr>
  `).join("");

  el("visibleCount").textContent = state.visible.length;
  const inventoryIsEmpty = state.items.length === 0;
  el("emptyTitle").textContent = inventoryIsEmpty ? "O balcão está vazio" : "Nada encontrado neste baú";
  el("emptyText").textContent = inventoryIsEmpty
    ? "Importe uma planilha ou carregue o exemplo para começar."
    : "Tente remover algum filtro ou buscar por outro termo.";
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
  const sellableItems = state.items.filter(item => !isCashItem(item));
  const allSelected = sellableItems.length > 0 && sellableItems.every(item => item.vender);
  document.querySelectorAll(".toggle-all-trigger").forEach(button => {
    button.textContent = allSelected ? "Desmarcar tudo" : "Marcar tudo";
    button.title = allSelected ? "Desmarcar todos os itens vendáveis" : "Marcar todos os itens vendáveis";
  });
}

function renderCategories() {
  const categories = [...new Set(state.items.map(item => item.categoria))].sort((a, b) => a.localeCompare(b, "pt-BR"));
  state.categories = state.categories.filter(category => categories.includes(category));
  el("categoryMenu").innerHTML = `
    <button class="category-clear" type="button">Todas as categorias</button>
    ${categories.map(category => `
      <label class="category-option">
        <input type="checkbox" value="${escapeHtml(category)}" ${state.categories.includes(category) ? "checked" : ""}>
        <span>${escapeHtml(category)}</span>
      </label>
    `).join("")}
  `;
  el("categoryLabel").textContent = state.categories.length
    ? `${state.categories.length} categoria${state.categories.length > 1 ? "s" : ""}`
    : "Todas as categorias";
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

function setAllSale(value) {
  state.items.forEach(item => {
    if (!isCashItem(item)) item.vender = value;
  });
  saveState();
  refresh();
  showToast(value ? "Todos os itens vendáveis foram marcados." : "Todos os itens foram desmarcados.");
}

function exportSpreadsheet() {
  if (!window.XLSX) {
    showToast("O exportador de planilhas não carregou.");
    return;
  }

  const inventoryRows = [...state.items]
    .sort((a, b) => a.order - b.order)
    .map(item => ({
      "Origem": item.origem,
      "Item": item.item,
      "Qtd.": item.quantidade,
      "Valor real (gp)": item.valorReal || "",
      "Valor estimado (gp)": item.valorEstimado || "",
      "Como foi estimado": item.estimativa,
      "Categoria": item.categoria,
      "Vender?": item.vender ? "Sim" : "Não",
      "Valor base (gp)": item.valorBase,
      "Bônus aplicado": item.bonus,
      "Venda líquida (gp)": saleValue(item),
      "Observação": item.observacao
    }));

  const cash = state.items
    .filter(isCashItem)
    .reduce((sum, item) => sum + item.valorBase * item.quantidade, 0);
  const sales = state.items.reduce((sum, item) => sum + saleValue(item), 0);
  const kept = state.items
    .filter(item => !item.vender && !isCashItem(item))
    .reduce((sum, item) => sum + item.valorBase * item.quantidade, 0);
  const total = cash + sales;

  const summaryRows = [
    ["Resumo do caixa e divisão — LootZ", ""],
    [],
    ["Componente", "Total (gp)"],
    ["Moedas e recompensas em caixa", cash],
    ["Receita estimada de vendas", sales],
    ["Total disponível para dividir", total],
    ["Divisão por pessoa", total / Math.max(1, state.members)],
    ["Itens mantidos (valor de referência)", kept]
  ];
  const categoryBonus = category => {
    const item = state.items.find(candidate => normalized(candidate.categoria) === normalized(category));
    return item?.bonus || 0;
  };
  const parameterRows = [
    ["Parâmetro", "Valor"],
    ["Bônus de venda — itens comuns", categoryBonus("Item comum")],
    ["Bônus de venda — gemas/arte", categoryBonus("Joia") || categoryBonus("Arte")],
    ["Bônus de venda — itens mágicos", categoryBonus("Mágico")],
    ["Divisores (membros)", state.members]
  ];
  const calculationRows = [
    ["Categoria", "Vender?", "Base", "Caixa", "Venda", "Mantido"],
    ...[...state.items]
      .sort((a, b) => a.order - b.order)
      .map(item => [
        item.categoria,
        item.vender ? "Sim" : "Não",
        item.valorBase * item.quantidade,
        isCashItem(item) ? item.valorBase * item.quantidade : 0,
        saleValue(item),
        !item.vender && !isCashItem(item) ? item.valorBase * item.quantidade : 0
      ])
  ];

  const workbook = XLSX.utils.book_new();
  const inventorySheet = XLSX.utils.json_to_sheet(inventoryRows);
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
  const parameterSheet = XLSX.utils.aoa_to_sheet(parameterRows);
  const calculationSheet = XLSX.utils.aoa_to_sheet(calculationRows);

  inventorySheet["!cols"] = [
    { wch: 30 }, { wch: 42 }, { wch: 8 }, { wch: 17 }, { wch: 21 }, { wch: 48 },
    { wch: 18 }, { wch: 11 }, { wch: 18 }, { wch: 17 }, { wch: 21 }, { wch: 48 }
  ];
  summarySheet["!cols"] = [{ wch: 40 }, { wch: 20 }];
  parameterSheet["!cols"] = [{ wch: 32 }, { wch: 16 }];
  calculationSheet["!cols"] = [{ wch: 20 }, { wch: 12 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];

  XLSX.utils.book_append_sheet(workbook, inventorySheet, "Inventário");
  XLSX.utils.book_append_sheet(workbook, summarySheet, "Resumo");
  XLSX.utils.book_append_sheet(workbook, parameterSheet, "Parâmetros");
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(parameterRows), "Parametros");
  XLSX.utils.book_append_sheet(workbook, calculationSheet, "Calculos");
  XLSX.writeFile(workbook, "inventario_lootz_atualizado.xlsx");
  showToast("A planilha atualizada foi exportada.");
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

let soundEnabled = true;
let soundStarted = false;

async function startAmbience(silent = false) {
  if (!soundEnabled || soundStarted) return;
  const audio = el("tavernAudio");
  audio.volume = number(el("volumeSlider").value) / 100;
  try {
    await audio.play();
    soundStarted = true;
    updateSoundButton();
    if (!silent) showToast("A música da taverna começou.");
  } catch (_) {
    soundStarted = false;
  }
}

function stopAmbience() {
  const audio = el("tavernAudio");
  audio.pause();
  soundStarted = false;
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
el("loadExample").addEventListener("click", loadExample);

el("inventoryBody").addEventListener("change", event => {
  const item = state.items.find(candidate => String(candidate.id) === event.target.dataset.id);
  if (!item) return;

  if (event.target.matches("input[type='checkbox']")) {
    item.vender = event.target.checked;
  } else if (event.target.matches("input[data-field='valorBase']")) {
    item.valorBase = Math.max(0, number(event.target.value));
  } else if (event.target.matches("input[data-field='bonus']")) {
    item.bonus = number(event.target.value) / 100;
  } else {
    return;
  }

  saveState();
  refresh();
});

el("searchInput").addEventListener("input", event => { state.search = event.target.value; renderTable(); });
el("categoryToggle").addEventListener("click", () => {
  const menu = el("categoryMenu");
  const willOpen = menu.hidden;
  menu.hidden = !willOpen;
  el("categoryToggle").setAttribute("aria-expanded", String(willOpen));
});
el("categoryMenu").addEventListener("change", event => {
  if (!event.target.matches("input[type='checkbox']")) return;
  state.categories = [...el("categoryMenu").querySelectorAll("input:checked")].map(input => input.value);
  renderCategories();
  renderTable();
});
el("categoryMenu").addEventListener("click", event => {
  if (!event.target.closest(".category-clear")) return;
  state.categories = [];
  renderCategories();
  renderTable();
});
document.addEventListener("click", event => {
  if (event.target.closest("#categoryFilter")) return;
  el("categoryMenu").hidden = true;
  el("categoryToggle").setAttribute("aria-expanded", "false");
});
el("saleFilter").addEventListener("change", event => { state.sale = event.target.value; renderTable(); });
el("sortSelect").addEventListener("change", event => { state.sort = event.target.value; renderTable(); });
document.querySelectorAll(".toggle-all-trigger").forEach(button => button.addEventListener("click", () => {
  const sellableItems = state.items.filter(item => !isCashItem(item));
  const allSelected = sellableItems.length > 0 && sellableItems.every(item => item.vender);
  setAllSale(!allSelected);
}));

let syncingTableScroll = false;
const tableScroll = el("tableScroll");
const mobileTableHeader = el("mobileTableHeader");
tableScroll.addEventListener("scroll", () => {
  if (syncingTableScroll) return;
  syncingTableScroll = true;
  mobileTableHeader.scrollLeft = tableScroll.scrollLeft;
  requestAnimationFrame(() => { syncingTableScroll = false; });
});
mobileTableHeader.addEventListener("scroll", () => {
  if (syncingTableScroll) return;
  syncingTableScroll = true;
  tableScroll.scrollLeft = mobileTableHeader.scrollLeft;
  requestAnimationFrame(() => { syncingTableScroll = false; });
});
el("exportSpreadsheet").addEventListener("click", exportSpreadsheet);
el("soundToggle").addEventListener("click", toggleAmbience);
el("volumeSlider").addEventListener("input", event => {
  const volume = number(event.target.value);
  el("tavernAudio").volume = volume / 100;
  event.target.setAttribute("aria-valuetext", `${volume}%`);
  event.target.style.background = `linear-gradient(90deg, var(--gold) 0 ${volume}%, rgba(255,255,255,.18) ${volume}% 100%)`;
  if (volume > 0 && soundEnabled && !soundStarted) startAmbience(true);
});
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
loadInitialState();
