const STORAGE_KEY = "bookkeeping-webapp-v3";
const LEGACY_KEYS = ["bookkeeping-webapp-v2", "bookkeeping-webapp-v1"];

const MONTHS = ["1 月", "2 月", "3 月", "4 月", "5 月", "6 月", "7 月", "8 月", "9 月", "10 月", "11 月", "12 月"];
const ACCOUNTS = ["儲蓄險", "現金", "中國信託", "國泰世華", "Richart", "Line bank", "郵局"];
const INCOME_FIELDS = ["薪資收入", "團體分紅", "分紅", "申報", "績效獎金", "利息收入", "其他收入", "補助津貼"];
const FIXED_EXPENSE_FIELDS = ["房租", "YT訂閱費", "Office365", "勞保費", "健保費", "小二保險費", "小二安親班"];
const DEFAULT_VARIABLE_EXPENSE_FIELDS = ["交通費", "娛樂", "食物", "日用品", "醫療", "個人項目", "小二賞銀", "其他支出"];
const COMPUTED_FIELDS = new Set(["收入小計", "支出總計", "收支", "固定支出小計", "支出小計", "總資產"]);
const PIE_COLORS = ["#bf5f38", "#d48758", "#e8aa76", "#557c53", "#7ba06f", "#a76d43", "#8f4e2e", "#d9b07f", "#a0b999"];

const elements = {
  yearSelect: document.querySelector("#yearSelect"),
  addYearButton: document.querySelector("#addYearButton"),
  seedDemoButton: document.querySelector("#seedDemoButton"),
  annualTableContainer: document.querySelector("#annualTableContainer"),
  expenseForm: document.querySelector("#expenseForm"),
  editingExpenseId: document.querySelector("#editingExpenseId"),
  expenseDate: document.querySelector("#expenseDate"),
  expenseCategory: document.querySelector("#expenseCategory"),
  expenseAmount: document.querySelector("#expenseAmount"),
  expenseNote: document.querySelector("#expenseNote"),
  expenseSubmitButton: document.querySelector("#expenseSubmitButton"),
  cancelEditButton: document.querySelector("#cancelEditButton"),
  dailyFormTitle: document.querySelector("#dailyFormTitle"),
  dailyMonthSelect: document.querySelector("#dailyMonthSelect"),
  dailyMonthlySubtotal: document.querySelector("#dailyMonthlySubtotal"),
  dailyExpenseList: document.querySelector("#dailyExpenseList"),
  categoryForm: document.querySelector("#categoryForm"),
  editingCategoryId: document.querySelector("#editingCategoryId"),
  categoryName: document.querySelector("#categoryName"),
  categorySubmitButton: document.querySelector("#categorySubmitButton"),
  cancelCategoryEditButton: document.querySelector("#cancelCategoryEditButton"),
  categoryList: document.querySelector("#categoryList"),
  monthRange: document.querySelector("#monthRange"),
  monthLabels: document.querySelector("#monthLabels"),
  monthlyDetailTitle: document.querySelector("#monthlyDetailTitle"),
  monthlySummaryCards: document.querySelector("#monthlySummaryCards"),
  monthlyManualSections: document.querySelector("#monthlyManualSections"),
  monthlyExpenseList: document.querySelector("#monthlyExpenseList"),
  incomePie: document.querySelector("#incomePie"),
  fixedPie: document.querySelector("#fixedPie"),
  variablePie: document.querySelector("#variablePie"),
  incomeLegend: document.querySelector("#incomeLegend"),
  fixedLegend: document.querySelector("#fixedLegend"),
  variableLegend: document.querySelector("#variableLegend"),
  expenseItemTemplate: document.querySelector("#expenseItemTemplate"),
  navLinks: document.querySelectorAll(".nav-link"),
  views: document.querySelectorAll(".view"),
};

let state = loadState();

bootstrap();

function bootstrap() {
  ensureYear(state.selectedYear);
  renderMonthLabels();
  bindEvents();
  renderAll();
}

function bindEvents() {
  elements.yearSelect.addEventListener("change", (event) => {
    state.selectedYear = event.target.value;
    ensureYear(state.selectedYear);
    saveState();
    renderAll();
  });

  elements.addYearButton.addEventListener("click", () => {
    const nextYear = String(Number(state.selectedYear) + 1);
    if (!state.years[nextYear]) {
      state.years[nextYear] = createEmptyYear();
    }
    state.selectedYear = nextYear;
    saveState();
    renderAll();
  });

  elements.seedDemoButton.addEventListener("click", () => {
    state.years[state.selectedYear] = createDemoYear(state.selectedYear);
    saveState();
    renderAll();
  });

  elements.navLinks.forEach((button) => {
    button.addEventListener("click", () => {
      setActiveView(button.dataset.view);
    });
  });

  window.addEventListener("hashchange", () => {
    const hashView = normalizeHashView(window.location.hash);
    if (hashView && hashView !== state.activeView) {
      state.activeView = hashView;
      renderViewState();
      saveState();
    }
  });

  elements.expenseForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const expense = {
      id: elements.editingExpenseId.value || crypto.randomUUID(),
      date: elements.expenseDate.value,
      category: elements.expenseCategory.value,
      amount: Number(elements.expenseAmount.value || 0),
      note: elements.expenseNote.value.trim(),
    };

    if (!expense.date || !expense.category || !expense.amount) {
      return;
    }

    const yearData = state.years[state.selectedYear];
    const existingIndex = yearData.dailyExpenses.findIndex((item) => item.id === expense.id);
    if (existingIndex >= 0) {
      yearData.dailyExpenses[existingIndex] = expense;
    } else {
      yearData.dailyExpenses.push(expense);
    }

    state.selectedMonth = new Date(expense.date).getMonth() + 1;
    resetExpenseForm();
    saveState();
    renderAll();
    setActiveView("daily");
  });

  elements.cancelEditButton.addEventListener("click", resetExpenseForm);

  elements.categoryForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const name = elements.categoryName.value.trim();
    if (!name) {
      return;
    }

    const yearData = state.years[state.selectedYear];
    const editingId = elements.editingCategoryId.value;
    if (editingId) {
      const category = yearData.categories.find((item) => item.id === editingId);
      if (category) {
        const oldName = category.name;
        category.name = name;
        yearData.dailyExpenses.forEach((expense) => {
          if (expense.category === oldName) {
            expense.category = name;
          }
        });
      }
    } else {
      yearData.categories.push({ id: crypto.randomUUID(), name });
    }

    resetCategoryForm();
    saveState();
    renderAll();
  });

  elements.cancelCategoryEditButton.addEventListener("click", resetCategoryForm);

  elements.dailyMonthSelect.addEventListener("change", (event) => {
    state.selectedMonth = Number(event.target.value);
    saveState();
    renderDailyPage();
    renderMonthlyDetail();
    renderMonthLabels();
  });

  elements.monthRange.addEventListener("input", (event) => {
    state.selectedMonth = Number(event.target.value);
    saveState();
    renderDailyPage();
    renderMonthlyDetail();
    renderMonthLabels();
  });

  elements.monthLabels.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-month]");
    if (!button) {
      return;
    }
    state.selectedMonth = Number(button.dataset.month);
    saveState();
    renderDailyPage();
    renderMonthlyDetail();
    renderMonthLabels();
  });
}

function renderAll() {
  renderYearSelect();
  renderExpenseCategoryOptions();
  renderAnnualTable();
  renderDailyPage();
  renderCategoryList();
  renderMonthlyDetail();
  elements.expenseDate.value = elements.expenseDate.value || todayForInput();
  setActiveView(state.activeView || normalizeHashView(window.location.hash) || "annual", true);
}

function renderViewState() {
  elements.navLinks.forEach((button) => {
    button.classList.toggle("active", button.dataset.view === state.activeView);
  });
  elements.views.forEach((panel) => {
    panel.classList.toggle("active", panel.dataset.viewPanel === state.activeView);
  });
}

function setActiveView(view, skipHashUpdate = false) {
  state.activeView = view;
  renderViewState();
  if (!skipHashUpdate) {
    const desiredHash = `#${view}`;
    if (window.location.hash !== desiredHash) {
      window.location.hash = desiredHash;
    }
  }
  saveState();
}

function normalizeHashView(hash) {
  const raw = hash.replace(/^#/, "");
  return ["annual", "daily", "monthly"].includes(raw) ? raw : "";
}

function renderYearSelect() {
  const years = Object.keys(state.years).sort();
  elements.yearSelect.innerHTML = years
    .map((year) => `<option value="${year}" ${year === state.selectedYear ? "selected" : ""}>${year}</option>`)
    .join("");
}

function getYearData() {
  ensureYear(state.selectedYear);
  return state.years[state.selectedYear];
}

function getCategories() {
  return getYearData().categories;
}

function renderExpenseCategoryOptions() {
  elements.expenseCategory.innerHTML = getCategories()
    .map((category) => `<option value="${escapeHtml(category.name)}">${escapeHtml(category.name)}</option>`)
    .join("");
}

function renderCategoryList() {
  elements.categoryList.innerHTML = getCategories()
    .map(
      (category) => `
        <article class="category-item">
          <strong>${escapeHtml(category.name)}</strong>
          <div class="category-item-actions">
            <button class="mini-button" type="button" data-action="edit-category" data-category-id="${category.id}">編輯</button>
            <button class="mini-button danger" type="button" data-action="delete-category" data-category-id="${category.id}">刪除</button>
          </div>
        </article>
      `,
    )
    .join("");

  elements.categoryList.querySelectorAll("[data-action='edit-category']").forEach((button) => {
    button.addEventListener("click", () => startCategoryEdit(button.dataset.categoryId));
  });

  elements.categoryList.querySelectorAll("[data-action='delete-category']").forEach((button) => {
    button.addEventListener("click", () => deleteCategory(button.dataset.categoryId));
  });
}

function startCategoryEdit(categoryId) {
  const category = getCategories().find((item) => item.id === categoryId);
  if (!category) {
    return;
  }
  elements.editingCategoryId.value = category.id;
  elements.categoryName.value = category.name;
  elements.categorySubmitButton.textContent = "儲存分類";
  elements.cancelCategoryEditButton.classList.remove("hidden");
}

function deleteCategory(categoryId) {
  const yearData = getYearData();
  const category = yearData.categories.find((item) => item.id === categoryId);
  if (!category) {
    return;
  }
  if (yearData.dailyExpenses.some((expense) => expense.category === category.name)) {
    window.alert("這個分類已經有支出在用，先把那些支出改分類後再刪除。");
    return;
  }
  yearData.categories = yearData.categories.filter((item) => item.id !== categoryId);
  saveState();
  renderAll();
}

function resetCategoryForm() {
  elements.editingCategoryId.value = "";
  elements.categoryName.value = "";
  elements.categorySubmitButton.textContent = "新增分類";
  elements.cancelCategoryEditButton.classList.add("hidden");
}

function renderAnnualTable() {
  const yearData = getYearData();
  const computed = computeYear(state.selectedYear, yearData);
  const categoryNames = getCategories().map((category) => category.name);
  const tableSections = [
    { label: "Summary", rows: ["收入小計", "支出總計", "收支"] },
    { label: "Assets", rows: [...ACCOUNTS, "總資產"], asset: true },
    { label: "Income", rows: INCOME_FIELDS },
    { label: "Fixed Expense", rows: [...FIXED_EXPENSE_FIELDS, "固定支出小計"] },
    { label: "Daily Expense", rows: [...categoryNames, "支出小計"] },
  ];

  const header = `
    <thead>
      <tr>
        <th>項目</th>
        ${MONTHS.map((month) => `<th>${month}</th>`).join("")}
      </tr>
    </thead>
  `;

  const sections = tableSections.map((section) => {
    const rows = section.rows.map((field) => {
      const isComputed = COMPUTED_FIELDS.has(field);
      const isDynamicExpense = categoryNames.includes(field);
      const rowClass = [isComputed ? "computed-row" : "", section.asset ? "asset-row" : ""].filter(Boolean).join(" ");

      const cells = MONTHS.map((_, index) => {
        const month = index + 1;
        if (isComputed || isDynamicExpense) {
          return `<td><span class="readonly-money">${formatMoney(computed[month][field] || 0)}</span></td>`;
        }

        const value = yearData.monthlyData[month][field] ?? 0;
        return `
          <td>
            <input
              class="money-input"
              type="number"
              min="0"
              step="1"
              value="${value}"
              data-month="${month}"
              data-field="${escapeHtml(field)}"
            />
          </td>
        `;
      }).join("");

      return `<tr class="${rowClass}"><th scope="row">${escapeHtml(field)}</th>${cells}</tr>`;
    }).join("");

    return `
      <tbody>
        <tr class="section-row"><th colspan="13">${section.label}</th></tr>
        ${rows}
      </tbody>
    `;
  }).join("");

  elements.annualTableContainer.innerHTML = `<table class="annual-table">${header}${sections}</table>`;
  elements.annualTableContainer.querySelectorAll(".money-input").forEach((input) => {
    input.addEventListener("change", handleAnnualInputChange);
  });
}

function handleAnnualInputChange(event) {
  const { month, field } = event.target.dataset;
  getYearData().monthlyData[month][field] = Number(event.target.value || 0);
  saveState();
  renderAnnualTable();
  renderMonthlyDetail();
}

function renderDailyPage() {
  const expenses = getMonthExpenses(state.selectedYear, state.selectedMonth);
  elements.dailyMonthSelect.innerHTML = MONTHS.map(
    (month, index) => `<option value="${index + 1}" ${index + 1 === state.selectedMonth ? "selected" : ""}>${month}</option>`,
  ).join("");
  elements.dailyMonthlySubtotal.textContent = `本月小計：${formatMoney(expenses.reduce((sum, item) => sum + item.amount, 0))}`;
  renderGroupedExpenseList(elements.dailyExpenseList, expenses);
}

function renderMonthlyDetail() {
  const yearData = getYearData();
  const computed = computeYear(state.selectedYear, yearData);
  const categories = getCategories().map((category) => category.name);
  const month = state.selectedMonth;
  const monthData = computed[month];
  const manual = yearData.monthlyData[month];
  const expenses = getMonthExpenses(state.selectedYear, month);

  elements.monthlyDetailTitle.textContent = `${MONTHS[month - 1]}明細`;
  elements.monthRange.value = String(month);
  renderMonthLabels();

  const summaryCards = [
    ["收入小計", monthData["收入小計"]],
    ["支出總計", monthData["支出總計"]],
    ["收支", monthData["收支"]],
    ["總資產", monthData["總資產"]],
  ];

  elements.monthlySummaryCards.innerHTML = summaryCards
    .map(
      ([label, value]) => `
        <article class="summary-card">
          <span>${label}</span>
          <strong>${formatMoney(value)}</strong>
        </article>
      `,
    )
    .join("");

  renderPieChart(elements.incomePie, elements.incomeLegend, INCOME_FIELDS.map((field) => [field, manual[field] || 0]));
  renderPieChart(elements.fixedPie, elements.fixedLegend, FIXED_EXPENSE_FIELDS.map((field) => [field, manual[field] || 0]));
  renderPieChart(elements.variablePie, elements.variableLegend, categories.map((field) => [field, monthData[field] || 0]));

  const manualSections = [
    { title: "收入區", fields: INCOME_FIELDS, source: manual },
    { title: "固定支出區", fields: FIXED_EXPENSE_FIELDS, source: manual },
    { title: "支出分類區", fields: categories, source: monthData },
  ];

  elements.monthlyManualSections.innerHTML = manualSections
    .map(
      (section) => `
        <section class="manual-section">
          <h4>${section.title}</h4>
          <div class="manual-grid">
            ${section.fields
              .map(
                (field) => `
                  <article class="manual-card">
                    <span>${escapeHtml(field)}</span>
                    <strong>${formatMoney(section.source[field] || 0)}</strong>
                  </article>
                `,
              )
              .join("")}
          </div>
        </section>
      `,
    )
    .join("");

  renderGroupedExpenseList(elements.monthlyExpenseList, expenses);
}

function renderPieChart(chartElement, legendElement, entries) {
  const filtered = entries.filter(([, value]) => Number(value) > 0);
  if (!filtered.length) {
    chartElement.style.background = "conic-gradient(#ece5d8 0deg 360deg)";
    legendElement.innerHTML = `<div class="empty-state">這個區塊本月沒有資料。</div>`;
    return;
  }

  const total = filtered.reduce((sum, [, value]) => sum + Number(value), 0);
  let start = 0;
  const segments = filtered.map(([label, value], index) => {
    const portion = (Number(value) / total) * 360;
    const color = PIE_COLORS[index % PIE_COLORS.length];
    const segment = `${color} ${start}deg ${start + portion}deg`;
    start += portion;
    return { label, value, color, segment };
  });

  chartElement.style.background = `conic-gradient(${segments.map((segment) => segment.segment).join(", ")})`;
  legendElement.innerHTML = segments
    .map(
      (segment) => `
        <article class="legend-item">
          <div class="legend-left">
            <span class="swatch" style="background:${segment.color}"></span>
            <span>${escapeHtml(segment.label)}</span>
          </div>
          <strong>${formatMoney(segment.value)}</strong>
        </article>
      `,
    )
    .join("");
}

function renderGroupedExpenseList(container, expenses) {
  if (!expenses.length) {
    container.innerHTML = `<div class="empty-state">這個月份還沒有支出紀錄。</div>`;
    return;
  }

  const groups = groupExpensesByDate(expenses);
  container.innerHTML = groups
    .map((group) => {
      const items = group.items
        .map((expense) => {
          const node = elements.expenseItemTemplate.content.cloneNode(true);
          const article = node.querySelector(".expense-item");
          article.dataset.expenseId = expense.id;
          node.querySelector(".expense-item-title").textContent = expense.category;
          node.querySelector(".expense-item-note").textContent = expense.note || "沒有備註";
          node.querySelector(".expense-item-date").textContent = formatDate(expense.date);
          node.querySelector(".expense-item-amount").textContent = formatMoney(expense.amount);
          return article.outerHTML;
        })
        .join("");

      return `
        <section class="expense-group">
          <div class="expense-group-head">
            <strong>${formatDateHeadline(group.date)}</strong>
            <span class="expense-group-total">每日小計 ${formatMoney(group.total)}</span>
          </div>
          <div class="expense-list">${items}</div>
        </section>
      `;
    })
    .join("");

  container.querySelectorAll(".expense-edit").forEach((button) => {
    button.addEventListener("click", (event) => {
      const article = event.target.closest(".expense-item");
      if (article) {
        startExpenseEdit(article.dataset.expenseId);
      }
    });
  });

  container.querySelectorAll(".expense-delete").forEach((button) => {
    button.addEventListener("click", (event) => {
      const article = event.target.closest(".expense-item");
      if (article) {
        deleteExpense(article.dataset.expenseId);
      }
    });
  });
}

function groupExpensesByDate(expenses) {
  const groups = new Map();
  expenses
    .slice()
    .sort((a, b) => a.date.localeCompare(b.date))
    .forEach((expense) => {
      if (!groups.has(expense.date)) {
        groups.set(expense.date, { date: expense.date, total: 0, items: [] });
      }
      const group = groups.get(expense.date);
      group.items.push(expense);
      group.total += Number(expense.amount || 0);
    });
  return Array.from(groups.values());
}

function startExpenseEdit(expenseId) {
  const expense = getYearData().dailyExpenses.find((item) => item.id === expenseId);
  if (!expense) {
    return;
  }
  elements.editingExpenseId.value = expense.id;
  elements.expenseDate.value = expense.date;
  elements.expenseCategory.value = expense.category;
  elements.expenseAmount.value = String(expense.amount);
  elements.expenseNote.value = expense.note || "";
  elements.dailyFormTitle.textContent = "編輯每日支出";
  elements.expenseSubmitButton.textContent = "儲存修改";
  elements.cancelEditButton.classList.remove("hidden");
  setActiveView("daily");
}

function deleteExpense(expenseId) {
  const yearData = getYearData();
  yearData.dailyExpenses = yearData.dailyExpenses.filter((item) => item.id !== expenseId);
  saveState();
  renderAll();
}

function resetExpenseForm() {
  elements.editingExpenseId.value = "";
  elements.expenseForm.reset();
  elements.expenseDate.value = todayForInput();
  elements.dailyFormTitle.textContent = "新增每日支出";
  elements.expenseSubmitButton.textContent = "新增支出";
  elements.cancelEditButton.classList.add("hidden");
}

function getMonthExpenses(selectedYear, month) {
  return getYearData()
    .dailyExpenses
    .filter((expense) => {
      const expenseDate = new Date(expense.date);
      return expenseDate.getFullYear() === Number(selectedYear) && expenseDate.getMonth() + 1 === month;
    })
    .sort((a, b) => a.date.localeCompare(b.date));
}

function renderMonthLabels() {
  elements.monthLabels.innerHTML = MONTHS.map(
    (month, index) => `
      <button type="button" data-month="${index + 1}" class="${index + 1 === state.selectedMonth ? "active" : ""}">
        ${index + 1}
      </button>
    `,
  ).join("");
}

function computeYear(selectedYear, yearData) {
  const result = {};
  const categoryNames = yearData.categories.map((category) => category.name);

  for (let month = 1; month <= 12; month += 1) {
    const source = yearData.monthlyData[month];
    const monthExpenses = aggregateDailyExpenses(selectedYear, yearData.dailyExpenses, month, categoryNames);
    const incomeSubtotal = sumFields(source, INCOME_FIELDS);
    const fixedExpenseSubtotal = sumFields(source, FIXED_EXPENSE_FIELDS);
    const variableSubtotal = sumFields(monthExpenses, categoryNames);
    const totalExpense = fixedExpenseSubtotal + variableSubtotal;
    const assetTotal = sumFields(source, ACCOUNTS);

    result[month] = {
      ...source,
      ...monthExpenses,
      "收入小計": incomeSubtotal,
      "固定支出小計": fixedExpenseSubtotal,
      "支出小計": variableSubtotal,
      "支出總計": totalExpense,
      "收支": incomeSubtotal - totalExpense,
      "總資產": assetTotal,
    };
  }

  return result;
}

function aggregateDailyExpenses(selectedYear, expenses, targetMonth, categories) {
  const totals = Object.fromEntries(categories.map((field) => [field, 0]));
  expenses.forEach((expense) => {
    const expenseDate = new Date(expense.date);
    const expenseMonth = expenseDate.getMonth() + 1;
    if (
      expenseDate.getFullYear() === Number(selectedYear) &&
      expenseMonth === targetMonth &&
      totals[expense.category] !== undefined
    ) {
      totals[expense.category] += Number(expense.amount || 0);
    }
  });
  return totals;
}

function sumFields(source, fields) {
  return fields.reduce((sum, field) => sum + Number(source[field] || 0), 0);
}

function loadState() {
  const currentYear = String(new Date().getFullYear());
  const fallback = {
    selectedYear: currentYear,
    selectedMonth: new Date().getMonth() + 1,
    activeView: normalizeHashView(window.location.hash) || "annual",
    years: {
      [currentYear]: createEmptyYear(),
    },
  };

  const raw = [localStorage.getItem(STORAGE_KEY), ...LEGACY_KEYS.map((key) => localStorage.getItem(key))].find(Boolean);
  if (!raw) {
    return fallback;
  }

  try {
    const parsed = JSON.parse(raw);
    const years = parsed.years || fallback.years;
    Object.values(years).forEach(upgradeYearData);
    return {
      selectedYear: parsed.selectedYear || currentYear,
      selectedMonth: parsed.selectedMonth || new Date().getMonth() + 1,
      activeView: normalizeHashView(window.location.hash) || parsed.activeView || "annual",
      years,
    };
  } catch {
    return fallback;
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function ensureYear(year) {
  if (!state.years[year]) {
    state.years[year] = createEmptyYear();
  } else {
    upgradeYearData(state.years[year]);
  }
}

function upgradeYearData(yearData) {
  if (!yearData.categories) {
    yearData.categories = DEFAULT_VARIABLE_EXPENSE_FIELDS.map((name) => ({ id: crypto.randomUUID(), name }));
  }
  if (!yearData.monthlyData) {
    yearData.monthlyData = createEmptyYear().monthlyData;
  }
  if (!yearData.dailyExpenses) {
    yearData.dailyExpenses = [];
  }
  for (let month = 1; month <= 12; month += 1) {
    if (!yearData.monthlyData[month]) {
      yearData.monthlyData[month] = {};
    }
    [...ACCOUNTS, ...INCOME_FIELDS, ...FIXED_EXPENSE_FIELDS].forEach((field) => {
      if (typeof yearData.monthlyData[month][field] !== "number") {
        yearData.monthlyData[month][field] = Number(yearData.monthlyData[month][field] || 0);
      }
    });
  }
}

function createEmptyYear() {
  const monthlyData = {};
  for (let month = 1; month <= 12; month += 1) {
    monthlyData[month] = {};
    [...ACCOUNTS, ...INCOME_FIELDS, ...FIXED_EXPENSE_FIELDS].forEach((field) => {
      monthlyData[month][field] = 0;
    });
  }

  return {
    monthlyData,
    dailyExpenses: [],
    categories: DEFAULT_VARIABLE_EXPENSE_FIELDS.map((name) => ({ id: crypto.randomUUID(), name })),
  };
}

function createDemoYear(yearLabel) {
  const year = createEmptyYear();
  for (let month = 1; month <= 12; month += 1) {
    year.monthlyData[month]["薪資收入"] = 52000;
    year.monthlyData[month]["現金"] = 6000;
    year.monthlyData[month]["中國信託"] = 38000 + month * 1200;
    year.monthlyData[month]["國泰世華"] = 20000 + month * 900;
    year.monthlyData[month]["Richart"] = 10000 + month * 500;
    year.monthlyData[month]["Line bank"] = 8000;
    year.monthlyData[month]["郵局"] = 5000;
    year.monthlyData[month]["儲蓄險"] = 120000;
    year.monthlyData[month]["房租"] = 14000;
    year.monthlyData[month]["YT訂閱費"] = 199;
    year.monthlyData[month]["Office365"] = 219;
    year.monthlyData[month]["勞保費"] = 1100;
    year.monthlyData[month]["健保費"] = 900;
    year.monthlyData[month]["小二保險費"] = 700;
    year.monthlyData[month]["小二安親班"] = 4500;
  }

  year.monthlyData[1]["績效獎金"] = 6000;
  year.monthlyData[7]["團體分紅"] = 8000;
  year.monthlyData[9]["補助津貼"] = 4000;

  year.dailyExpenses = [
    { id: crypto.randomUUID(), date: `${yearLabel}-01-03`, category: "食物", amount: 220, note: "早餐和午餐" },
    { id: crypto.randomUUID(), date: `${yearLabel}-01-09`, category: "交通費", amount: 1200, note: "加油" },
    { id: crypto.randomUUID(), date: `${yearLabel}-01-14`, category: "日用品", amount: 530, note: "家用清潔" },
    { id: crypto.randomUUID(), date: `${yearLabel}-02-02`, category: "娛樂", amount: 450, note: "電影" },
    { id: crypto.randomUUID(), date: `${yearLabel}-02-19`, category: "個人項目", amount: 800, note: "剪髮" },
    { id: crypto.randomUUID(), date: `${yearLabel}-03-11`, category: "醫療", amount: 300, note: "掛號費" },
    { id: crypto.randomUUID(), date: `${yearLabel}-03-21`, category: "小二賞銀", amount: 200, note: "考試進步獎勵" },
  ];

  return year;
}

function formatMoney(value) {
  return new Intl.NumberFormat("zh-TW", {
    style: "currency",
    currency: "TWD",
    maximumFractionDigits: 0,
  }).format(Number(value || 0));
}

function formatDate(value) {
  return new Intl.DateTimeFormat("zh-TW", {
    month: "short",
    day: "numeric",
    weekday: "short",
  }).format(new Date(value));
}

function formatDateHeadline(value) {
  return new Intl.DateTimeFormat("zh-TW", {
    month: "long",
    day: "numeric",
    weekday: "long",
  }).format(new Date(value));
}

function todayForInput() {
  const today = new Date();
  const local = new Date(today.getTime() - today.getTimezoneOffset() * 60000);
  return local.toISOString().slice(0, 10);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}
