(() => {
  const datasets = resolveDatasets();
  const datasetKeys = orderProjectKeys(Object.keys(datasets || {}));
  if (!datasetKeys.length) {
    document.body.innerHTML = '<p style="padding:24px">Не знайдено дані платежів.</p>';
    return;
  }

  const PROJECT_KEY = "multi-projects-v1";
  const STORAGE_OVERRIDES = `payment_calendar_overrides_v2_${PROJECT_KEY}`;
  const STORAGE_SETTINGS = `payment_calendar_settings_v2_${PROJECT_KEY}`;
  const STORAGE_SNAPSHOT = `payment_calendar_snapshot_v2_${PROJECT_KEY}`;
  const STORAGE_SNAPSHOT_HISTORY = `payment_calendar_snapshot_history_v2_${PROJECT_KEY}`;
  const MAX_SNAPSHOT_HISTORY = 40;
  const NIVKI_ANNUAL_RATE = 0.07;
  const VIEW_MODE_LABELS = {
    month: "Місяці",
    quarter: "Квартали",
    year: "Роки",
  };

  const STATUS_ORDER = ["unpaid", "paid", "early"];
  const STATUS_LABELS = {
    paid: "Оплачено",
    unpaid: "Не оплачено",
    partial: "Не оплачено",
    early: "Достроково",
  };

  const $summaryPaid = document.getElementById("summaryPaid");
  const $summaryRemainingPlan = document.getElementById("summaryRemainingPlan");
  const $summaryRemainingEarly = document.getElementById("summaryRemainingEarly");
  const $summaryRemainingEarlyLabel = document.getElementById("summaryRemainingEarlyLabel");
  const $summaryTotal = document.getElementById("summaryTotal");
  const $summaryCount = document.getElementById("summaryCount");

  const $projectFilter = document.getElementById("projectFilter");
  const $viewModeSelect = document.getElementById("viewModeSelect");
  const $viewModeButtons = [...document.querySelectorAll(".mode-btn")];
  const $yearSelect = document.getElementById("yearSelect");
  const $monthSelect = document.getElementById("monthSelect");
  const $periodFilterLabel = document.getElementById("periodFilterLabel");
  const $calendarTitle = document.getElementById("calendarTitle");
  const $monthInfo = document.getElementById("monthInfo");
  const $calendarGrid = document.getElementById("calendarGrid");

  const $detailEmpty = document.getElementById("detailEmpty");
  const $detailForm = document.getElementById("detailForm");
  const $detailTitle = document.getElementById("detailTitle");
  const $detailMeta = document.getElementById("detailMeta");

  const $statusInput = document.getElementById("statusInput");
  const $paidUsdInput = document.getElementById("paidUsdInput");
  const $paidUahInput = document.getElementById("paidUahInput");
  const $paymentDateInput = document.getElementById("paymentDateInput");
  const $noteInput = document.getElementById("noteInput");
  const $clearItemBtn = document.getElementById("clearItemBtn");

  const $searchInput = document.getElementById("searchInput");
  const $fxRateInput = document.getElementById("fxRateInput");
  const $earlyTariffInput = document.getElementById("earlyTariffInput");
  const $exportBackupBtn = document.getElementById("exportBackupBtn");
  const $importBackupBtn = document.getElementById("importBackupBtn");
  const $importBackupFile = document.getElementById("importBackupFile");
  const $resetAllBtn = document.getElementById("resetAllBtn");
  const $currencyButtons = [...document.querySelectorAll(".currency-btn")];

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const startOfCurrentMonth = new Date(today.getFullYear(), today.getMonth(), 1);

  const payments = [];
  datasetKeys.forEach((projectKey) => {
    const dataset = datasets[projectKey];
    const rows = Array.isArray(dataset && dataset.payments) ? dataset.payments : [];
    rows.forEach((item, idx) => {
      const rawId = item && item.id !== undefined && item.id !== null ? String(item.id) : String(idx);
      const uid = `${projectKey}:${rawId}`;
      payments.push({
        ...item,
        uid,
        raw_id: rawId,
        project_key: projectKey,
        project_name: dataset.project_name || projectKey,
      });
    });
  });
  const idToPayment = new Map(payments.map((item) => [item.uid, item]));

  const years = [
    ...new Set(
      payments
        .map((item) => {
          if (item.period_year) return item.period_year;
          if (item.due_date) return Number(item.due_date.slice(0, 4));
          return null;
        })
        .filter((value) => Number.isFinite(value))
    ),
  ].sort((a, b) => a - b);
  const fallbackYear = years[0] || new Date().getFullYear();

  const defaultFx = computeDefaultFx();
  const initialProject = inferInitialProjectKey(datasetKeys);
  const initialSelection = payments.find((item) => item.project_key === initialProject) || payments[0];

  const state = {
    overrides: loadJson(STORAGE_OVERRIDES, {}),
    currency: "usd",
    selectedProjects: [initialProject],
    viewMode: "month",
    selectedYear: "all",
    selectedMonth: "all",
    selectedId: initialSelection ? initialSelection.uid : "",
    search: "",
    fxRate: defaultFx,
    earlyUnitUsd: 0,
    earlyUnitUah: 0,
    lastSavedAt: "",
  };

  hydrateSettings();
  bindEvents();
  renderAll();

  function bindEvents() {
    $currencyButtons.forEach((btn) => {
      btn.addEventListener("click", () => {
        state.currency = btn.dataset.currency === "uah" ? "uah" : "usd";
        persistSettings();
        renderAll();
      });
    });

    $searchInput.addEventListener("input", () => {
      state.search = ($searchInput.value || "").trim().toLowerCase();
      renderAll();
    });

    $fxRateInput.addEventListener("change", () => {
      const next = toNumber($fxRateInput.value);
      if (next !== null && next > 0) {
        state.fxRate = next;
        persistSettings();
        renderAll();
      }
    });

    if ($projectFilter) {
      $projectFilter.addEventListener("click", (event) => {
        const btn = event.target.closest(".project-chip");
        if (!btn) return;
        const project = btn.dataset.project;
        if (!project) return;
        if (project === "all") {
          state.selectedProjects = [...datasetKeys];
        } else {
          const current = new Set(state.selectedProjects);
          if (current.has(project)) {
            if (current.size === 1) return;
            current.delete(project);
          } else {
            current.add(project);
          }
          state.selectedProjects = [...current].filter((key) => datasetKeys.includes(key));
        }
        ensureProjectSelection();
        ensureSelectedMonthInRange();
        persistSettings();
        renderAll();
      });
    }

    if ($viewModeSelect) {
      const onModeSelect = () => applyViewMode($viewModeSelect.value);
      $viewModeSelect.addEventListener("change", onModeSelect);
      $viewModeSelect.addEventListener("input", onModeSelect);
    }
    $viewModeButtons.forEach((btn) => {
      btn.addEventListener("click", () => applyViewMode(btn.dataset.viewMode));
    });

    const onYearChange = () => {
      state.selectedYear = $yearSelect.value || "all";
      ensureSelectedMonthInRange();
      persistSettings();
      renderAll();
    };
    $yearSelect.addEventListener("change", onYearChange);
    $yearSelect.addEventListener("input", onYearChange);

    const onPeriodChange = () => {
      state.selectedMonth = $monthSelect.value || "all";
      persistSettings();
      renderAll();
    };
    $monthSelect.addEventListener("change", onPeriodChange);
    $monthSelect.addEventListener("input", onPeriodChange);

    $calendarGrid.addEventListener("click", (event) => {
      if (state.viewMode !== "month") return;
      const toggle = event.target.closest(".status-toggle");
      if (toggle) {
        const id = toggle.dataset.id;
        cycleStatus(id);
        event.stopPropagation();
        return;
      }

      const card = event.target.closest(".payment-card");
      if (!card) return;
      state.selectedId = card.dataset.id;
      persistSettings();
      renderCalendar();
      renderDetail();
    });

    $detailForm.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!state.selectedId) return;

      const status = $statusInput.value;
      const paidUsd = toNumber($paidUsdInput.value);
      const paidUah = toNumber($paidUahInput.value);
      const paymentDate = ($paymentDateInput.value || "").trim();
      const note = ($noteInput.value || "").trim();

      const patch = {
        status,
        paid_usd: paidUsd,
        paid_uah: paidUah,
        payment_date: paymentDate || null,
        note: note || null,
      };
      updateOverride(state.selectedId, patch);
      renderAll();
    });

    $clearItemBtn.addEventListener("click", () => {
      if (!state.selectedId) return;
      delete state.overrides[state.selectedId];
      persistOverrides();
      renderAll();
    });

    $resetAllBtn.addEventListener("click", () => {
      if (!window.confirm("Скинути всі локальні зміни статусів/сум?")) return;
      state.overrides = {};
      persistOverrides();
      renderAll();
    });

    $exportBackupBtn.addEventListener("click", () => {
      const snapshot = buildSnapshot();
      const json = JSON.stringify(snapshot, null, 2);
      const blob = new Blob([json], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      const stamp = snapshot.savedAt
        .replace(/[-:]/g, "")
        .replace("T", "-")
        .replace(/\..*$/, "");
      anchor.href = url;
      anchor.download = `payment-calendar-${PROJECT_KEY}-backup-${stamp}.json`;
      anchor.click();
      URL.revokeObjectURL(url);
    });

    $importBackupBtn.addEventListener("click", () => {
      $importBackupFile.click();
    });

    $importBackupFile.addEventListener("change", async () => {
      const file = $importBackupFile.files && $importBackupFile.files[0];
      if (!file) return;
      try {
        const text = await file.text();
        const payload = JSON.parse(text);
        if (!payload || typeof payload !== "object") throw new Error("invalid payload");
        if (payload.projectKey && payload.projectKey !== PROJECT_KEY) {
          throw new Error("project mismatch");
        }
        if (!window.confirm("Імпортувати backup і перезаписати поточні локальні зміни?")) return;
        applySnapshot(payload);
        persistOverrides();
        persistSettings();
        renderAll();
      } catch (error) {
        window.alert("Не вдалося імпортувати backup. Перевірте, що це JSON з цього календаря.");
      } finally {
        $importBackupFile.value = "";
      }
    });
  }

  function renderAll() {
    ensureProjectSelection();
    renderProjectFilters();
    renderCurrencyControls();
    renderSummary();
    renderPeriodFilters();
    syncSelectedWithVisible();
    renderCalendar();
    renderDetail();
  }

  function renderProjectFilters() {
    if (!$projectFilter) return;
    const selected = new Set(state.selectedProjects);
    const allActive = datasetKeys.every((key) => selected.has(key));
    const chips = [
      `<button type="button" class="project-chip ${allActive ? "active" : ""}" data-project="all">Всі</button>`,
      ...datasetKeys.map((key) => {
        const name = datasets[key] && datasets[key].project_name ? datasets[key].project_name : key;
        const count = payments.filter((item) => item.project_key === key).length;
        return `<button type="button" class="project-chip ${selected.has(key) ? "active" : ""}" data-project="${key}">${escapeHtml(
          name
        )} (${count})</button>`;
      }),
    ];
    $projectFilter.innerHTML = chips.join("");
  }

  function renderCurrencyControls() {
    $currencyButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.currency === state.currency);
    });
    $fxRateInput.value = state.fxRate.toFixed(4);
    if ($earlyTariffInput) {
      const baseValue = state.currency === "usd" ? state.earlyUnitUsd : state.earlyUnitUah;
      $earlyTariffInput.value = round2(baseValue);
      $earlyTariffInput.readOnly = true;
      $earlyTariffInput.title = "Базова ціна для розрахунку дострокового закриття";
    }
  }

  function renderSummary() {
    const rows = getSelectedPayments()
      .filter((payment) => !payment.exclude_from_summary)
      .map((payment) => ({ payment, view: buildView(payment) }));

    const paidRows = rows.filter(({ view }) => isSettledStatus(view.status));
    const futureUnpaidRows = rows.filter(({ payment, view }) => {
      if (!isCurrentOrFuturePeriod(payment)) return false;
      if (view.status === "paid") return false;
      if (view.status === "early") return false;
      return true;
    });
    const payableFutureRows = futureUnpaidRows.filter(
      ({ view }) => view.scheduleUsd > 0.009 || view.scheduleUah > 0.009
    );

    const paid = paidRows.reduce(
      (sum, row) => sum + (state.currency === "usd" ? row.view.paidUsd : row.view.paidUah),
      0
    );
    const remainingPlan = futureUnpaidRows.reduce(
      (sum, row) => sum + (state.currency === "usd" ? row.view.scheduleUsd : row.view.scheduleUah),
      0
    );

    const earlyCalc = computeEarlySettlement(payableFutureRows);
    state.earlyUnitUsd = earlyCalc.unitUsd;
    state.earlyUnitUah = earlyCalc.unitUah;

    const futureCount = payableFutureRows.length;
    const remainingEarlyUsd = earlyCalc.totalUsd;
    const remainingEarlyUah = earlyCalc.totalUah;
    const remainingEarly = state.currency === "usd" ? remainingEarlyUsd : remainingEarlyUah;
    const total = paid + remainingPlan;

    const paidCount = paidRows.length;
    const allCount = rows.length;
    $summaryPaid.textContent = formatCurrency(paid, state.currency);
    $summaryRemainingPlan.textContent = formatCurrency(remainingPlan, state.currency);
    $summaryRemainingEarly.textContent = formatCurrency(remainingEarly, state.currency);
    $summaryRemainingEarlyLabel.textContent = earlyCalc.label || `Залишок достроково (${futureCount} платежів)`;
    $summaryTotal.textContent = formatCurrency(total, state.currency);
    $summaryCount.textContent = `${paidCount}/${allCount} оплачено, майбутніх неоплачених: ${futureUnpaidRows.length}`;
  }

  function computeEarlySettlement(payableFutureRows) {
    const groups = new Map();
    payableFutureRows.forEach((row) => {
      const key = row.payment.project_key || "default";
      const list = groups.get(key) || [];
      list.push(row);
      groups.set(key, list);
    });

    let totalUsd = 0;
    let totalUah = 0;
    let unitUsd = 0;
    let unitUah = 0;
    const labels = [];
    const multipleProjects = groups.size > 1;

    groups.forEach((rows, projectKey) => {
      const calc = projectKey === "nivki" ? computeNivkiEarlySettlement(rows) : computeGenericEarlySettlement(rows);
      totalUsd += numberOr(calc.totalUsd, 0);
      totalUah += numberOr(calc.totalUah, 0);
      if (groups.size === 1) {
        unitUsd = numberOr(calc.unitUsd, 0);
        unitUah = numberOr(calc.unitUah, 0);
      }
      const projectName = datasets[projectKey] && datasets[projectKey].project_name ? datasets[projectKey].project_name : projectKey;
      labels.push(multipleProjects ? `${projectName}: ${calc.label}` : calc.label);
    });

    if (!groups.size) {
      labels.push("без майбутніх платежів");
    }

    return {
      unitUsd,
      unitUah,
      totalUsd,
      totalUah,
      label: `Залишок достроково (${labels.join("; ")})`,
    };
  }

  function computeGenericEarlySettlement(payableFutureRows) {
    const earlyBaseRow = getEarliestPayableRow(payableFutureRows);
    const unitUsd = earlyBaseRow ? earlyBaseRow.view.scheduleUsd : 0;
    const unitUah = earlyBaseRow ? earlyBaseRow.view.scheduleUah : 0;
    const futureCount = payableFutureRows.length;
    const baseLabel = earlyBaseRow
      ? readablePaymentTitle(earlyBaseRow.payment)
      : "немає базового місяця";
    return {
      unitUsd,
      unitUah,
      totalUsd: futureCount * unitUsd,
      totalUah: futureCount * unitUah,
      label: `${futureCount} платежів, база: ${baseLabel}`,
    };
  }

  function computeNivkiEarlySettlement(payableFutureRows) {
    const monthlyRate = NIVKI_ANNUAL_RATE / 12;
    const sortedRows = [...payableFutureRows].sort((a, b) => {
      const da = parseDate(a.payment.due_date);
      const db = parseDate(b.payment.due_date);
      if (da && db) return da - db;
      return String(a.payment.uid).localeCompare(String(b.payment.uid), "uk");
    });
    const count = sortedRows.length;

    if (!count) {
      return {
        unitUsd: 0,
        unitUah: 0,
        totalUsd: 0,
        totalUah: 0,
        label: "Залишок достроково (без майбутніх платежів)",
      };
    }

    const scheduleUah = sortedRows.map((row) => numberOr(row.view.scheduleUah, 0));
    let principalPerMonthUah = 0;
    for (let i = 0; i < scheduleUah.length - 1; i += 1) {
      const diff = scheduleUah[i] - scheduleUah[i + 1];
      if (diff > 0.0001) {
        principalPerMonthUah = diff / monthlyRate;
        break;
      }
    }
    if (principalPerMonthUah <= 0) {
      const firstPaymentUah = scheduleUah[0];
      principalPerMonthUah = firstPaymentUah / (1 + count * monthlyRate);
    }

    const totalUah = Math.max(principalPerMonthUah, 0) * count;
    const firstRate = effectiveRate(sortedRows[0].payment);
    const totalUsd = firstRate > 0 ? totalUah / firstRate : 0;

    const unitUah = principalPerMonthUah;
    const unitUsd = count > 0 ? totalUsd / count : 0;

    return {
      unitUsd,
      unitUah,
      totalUsd,
      totalUah,
      label: `Залишок достроково (тіло кредиту без 7%: ${count} міс.)`,
    };
  }

  function renderPeriodFilters() {
    const scopedPayments = getSelectedPayments();

    if ($calendarTitle) {
      $calendarTitle.textContent = VIEW_MODE_LABELS[state.viewMode] || VIEW_MODE_LABELS.month;
    }
    if ($periodFilterLabel) {
      if (state.viewMode === "month") $periodFilterLabel.textContent = "Місяць";
      else if (state.viewMode === "quarter") $periodFilterLabel.textContent = "Квартал";
      else $periodFilterLabel.textContent = "Рік";
    }
    if ($viewModeSelect) {
      $viewModeSelect.value = state.viewMode;
    }
    $viewModeButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.viewMode === state.viewMode);
    });

    const yearCounts = new Map();
    scopedPayments.forEach((item) => {
      const year = getPaymentYear(item);
      yearCounts.set(year, (yearCounts.get(year) || 0) + 1);
    });
    const scopedYears = [...yearCounts.keys()].sort((a, b) => a - b);

    if (state.viewMode === "year") {
      state.selectedYear = "all";
    }
    if (state.selectedYear !== "all" && !scopedYears.includes(Number(state.selectedYear))) {
      state.selectedYear = "all";
    }

    const yearOptions = [
      `<option value="all"${state.selectedYear === "all" ? " selected" : ""}>Всі роки (${scopedPayments.length})</option>`,
      ...scopedYears.map((year) => {
        const selected = state.selectedYear === String(year) ? " selected" : "";
        return `<option value="${year}"${selected}>${year} (${yearCounts.get(year) || 0})</option>`;
      }),
    ];
    $yearSelect.innerHTML = yearOptions.join("");
    $yearSelect.disabled = state.viewMode === "year";

    ensureSelectedMonthInRange();
    const periodOptionsData = getPeriodOptions(scopedPayments, state.selectedYear, state.viewMode);
    const totalInMode = periodOptionsData.reduce((sum, entry) => sum + entry.count, 0);
    const allLabel =
      state.viewMode === "month"
        ? "Всі місяці"
        : state.viewMode === "quarter"
          ? "Всі квартали"
          : "Всі роки";
    const monthOptions = [
      `<option value="all"${state.selectedMonth === "all" ? " selected" : ""}>${allLabel} (${totalInMode})</option>`,
      ...periodOptionsData.map((entry) => {
        const selected = state.selectedMonth === String(entry.key) ? " selected" : "";
        return `<option value="${entry.key}"${selected}>${escapeHtml(entry.label)} (${entry.count})</option>`;
      }),
    ];
    $monthSelect.innerHTML = monthOptions.join("");

    const visible = getVisiblePayments();
    const groupCount = state.viewMode === "month" ? visible.length : buildAggregateGroups(visible, state.viewMode).length;
    const yearLabel = state.selectedYear === "all" ? "усі роки" : state.selectedYear;
    const periodLabel = describeSelectedPeriod(state.selectedMonth, state.selectedYear, state.viewMode);
    const searchLabel = state.search ? `, пошук: “${state.search}”` : "";
    const countLabel = state.viewMode === "month" ? "місяців" : state.viewMode === "quarter" ? "кварталів" : "років";
    $monthInfo.textContent = `Показано ${groupCount} ${countLabel}: ${yearLabel}, ${periodLabel}${searchLabel}`;
  }

  function applyViewMode(nextMode) {
    const normalized = nextMode === "quarter" || nextMode === "year" ? nextMode : "month";
    state.viewMode = normalized;
    if (state.viewMode === "year") {
      state.selectedYear = "all";
    }
    state.selectedMonth = "all";
    persistSettings();
    renderAll();
  }

  function getPeriodOptions(items, yearValue, mode) {
    const bucket = new Map();
    items.forEach((item) => {
      const year = getPaymentYear(item);
      if (mode !== "year" && yearValue !== "all" && year !== Number(yearValue)) return;

      if (mode === "month") {
        const month = getPaymentMonth(item);
        if (!month) return;
        const key = String(month);
        const entry = bucket.get(key) || { key, label: monthLong(month), count: 0, sortA: month };
        entry.count += 1;
        bucket.set(key, entry);
        return;
      }

      if (mode === "quarter") {
        const month = getPaymentMonth(item);
        if (!month) return;
        const quarter = quarterOfMonth(month);
        const scopedToYear = yearValue !== "all";
        const key = scopedToYear ? `Q${quarter}` : `${year}-Q${quarter}`;
        const label = scopedToYear ? `Q${quarter}` : `Q${quarter} ${year}`;
        const entry = bucket.get(key) || {
          key,
          label,
          count: 0,
          sortA: year,
          sortB: quarter,
        };
        entry.count += 1;
        bucket.set(key, entry);
        return;
      }

      const key = String(year);
      const entry = bucket.get(key) || { key, label: String(year), count: 0, sortA: year };
      entry.count += 1;
      bucket.set(key, entry);
    });

    return [...bucket.values()].sort((a, b) => {
      const a1 = numberOr(a.sortA, 0);
      const b1 = numberOr(b.sortA, 0);
      if (a1 !== b1) return a1 - b1;
      return numberOr(a.sortB, 0) - numberOr(b.sortB, 0);
    });
  }

  function ensureSelectedMonthInRange() {
    if (state.selectedMonth === "all") return;
    const exists = getPeriodOptions(getSelectedPayments(), state.selectedYear, state.viewMode).some(
      (entry) => String(entry.key) === String(state.selectedMonth)
    );
    if (!exists) {
      state.selectedMonth = "all";
    }
  }

  function syncSelectedWithVisible() {
    if (state.viewMode !== "month") {
      state.selectedId = "";
      return;
    }
    const visible = getVisiblePayments();
    if (!visible.length) {
      state.selectedId = "";
      return;
    }
    if (!state.selectedId || !visible.some((item) => item.uid === state.selectedId)) {
      state.selectedId = visible[0].uid;
    }
  }

  function renderCalendar() {
    const visible = getVisiblePayments();

    if (!visible.length) {
      $calendarGrid.innerHTML = '<div class="empty-note">Нічого не знайдено за поточним фільтром.</div>';
      return;
    }

    if (state.viewMode !== "month") {
      renderAggregateCalendar(visible, state.viewMode);
      return;
    }

    const cards = visible.map((item) => {
      const view = buildView(item);
      const selected = state.selectedId === item.uid ? "selected" : "";
      const overdue = view.overdue ? "overdue" : "";
      const amount = state.currency === "usd" ? view.amountUsd : view.amountUah;
      const projectText =
        state.selectedProjects.length > 1 ? `Проєкт: ${escapeHtml(item.project_name || item.project_key)}` : "";
      const plannedText = `По графіку: ${formatCurrency(view.scheduleUsd, "usd")} / ${formatCurrency(
        view.scheduleUah,
        "uah"
      )}`;
      const paidText = `Оплачено: ${formatCurrency(view.paidUsd, "usd")} / ${formatCurrency(
        view.paidUah,
        "uah"
      )}`;
      const paymentDateText = view.paymentDate ? `Оплата: ${formatDate(view.paymentDate)}` : "Дата оплати не вказана";
      const note = view.note || (item.flags && item.flags.early ? "Позначено як достроково в джерелі." : "");
      const monthIndexText =
        typeof item.month_index === "number" && Number.isFinite(item.month_index)
          ? `Місяць плану: №${item.month_index}`
          : "";

      return `
        <article class="payment-card ${selected} ${overdue}" id="payment-${safeDomId(item.uid)}" data-id="${item.uid}">
          <div class="card-top">
            <h3 class="card-title">${escapeHtml(readablePaymentTitle(item))}</h3>
            <button class="status-toggle" type="button" data-id="${item.uid}" title="Змінити статус">
              <span class="status-pill ${view.status}">${escapeHtml(STATUS_LABELS[view.status])}</span>
            </button>
          </div>
          <div class="card-amount">${formatCurrency(amount, state.currency)}</div>
          ${projectText ? `<p class="card-meta">${projectText}</p>` : ""}
          <p class="card-meta">${plannedText}</p>
          <p class="card-meta">${paidText}</p>
          ${monthIndexText ? `<p class="card-meta">${monthIndexText}</p>` : ""}
          <p class="card-meta">${paymentDateText}</p>
          ${note ? `<p class="card-note">${escapeHtml(note)}</p>` : ""}
        </article>
      `;
    });

    $calendarGrid.innerHTML = cards.join("");
  }

  function renderAggregateCalendar(items, mode) {
    const groups = buildAggregateGroups(items, mode);
    const cards = groups.map((group) => {
      const selectedAmount = state.currency === "usd" ? group.remainingUsd : group.remainingUah;
      let statusClass = "unpaid";
      if (group.unpaidCount === 0 && group.paidCount > 0) {
        statusClass = "paid";
      } else if (group.unpaidCount === 0 && group.paidCount === 0 && group.earlyCount > 0) {
        statusClass = "early";
      }

      return `
        <article class="payment-card" data-group="${escapeHtml(group.key)}">
          <div class="card-top">
            <h3 class="card-title">${escapeHtml(group.label)}</h3>
            <span class="status-pill ${statusClass}">${
              statusClass === "paid" ? "Оплачено" : statusClass === "early" ? "Достроково" : "Не оплачено"
            }</span>
          </div>
          <div class="card-amount">${formatCurrency(selectedAmount, state.currency)}</div>
          <p class="card-meta">По графіку: ${formatCurrency(group.scheduleUsd, "usd")} / ${formatCurrency(
            group.scheduleUah,
            "uah"
          )}</p>
          <p class="card-meta">Оплачено: ${formatCurrency(group.paidUsd, "usd")} / ${formatCurrency(
            group.paidUah,
            "uah"
          )}</p>
          <p class="card-meta">Залишок: ${formatCurrency(group.remainingUsd, "usd")} / ${formatCurrency(
            group.remainingUah,
            "uah"
          )}</p>
          <p class="card-meta">Періодів у групі: ${group.count}</p>
        </article>
      `;
    });
    $calendarGrid.innerHTML = cards.join("");
  }

  function buildAggregateGroups(items, mode) {
    const bucket = new Map();
    items.forEach((item) => {
      const view = buildView(item);
      const year = getPaymentYear(item);
      const month = getPaymentMonth(item) || 1;
      const quarter = quarterOfMonth(month);

      let key = "";
      let label = "";
      let sortA = year;
      let sortB = 0;
      if (mode === "quarter") {
        key = `${year}-Q${quarter}`;
        label = `Q${quarter} ${year}`;
        sortB = quarter;
      } else {
        key = String(year);
        label = String(year);
      }

      const entry = bucket.get(key) || {
        key,
        label,
        sortA,
        sortB,
        count: 0,
        scheduleUsd: 0,
        scheduleUah: 0,
        paidUsd: 0,
        paidUah: 0,
        remainingUsd: 0,
        remainingUah: 0,
        paidCount: 0,
        unpaidCount: 0,
        earlyCount: 0,
      };

      const includeInRemaining = view.status !== "paid" && view.status !== "early";
      entry.count += 1;
      if (view.status === "paid") entry.paidCount += 1;
      else if (view.status === "early") entry.earlyCount += 1;
      else entry.unpaidCount += 1;
      entry.scheduleUsd += view.scheduleUsd;
      entry.scheduleUah += view.scheduleUah;
      entry.paidUsd += view.paidUsd;
      entry.paidUah += view.paidUah;
      if (includeInRemaining) {
        entry.remainingUsd += view.scheduleUsd;
        entry.remainingUah += view.scheduleUah;
      }
      bucket.set(key, entry);
    });

    return [...bucket.values()]
      .sort((a, b) => (a.sortA !== b.sortA ? a.sortA - b.sortA : a.sortB - b.sortB))
      .map((entry) => ({
        ...entry,
        scheduleUsd: round2(entry.scheduleUsd),
        scheduleUah: round2(entry.scheduleUah),
        paidUsd: round2(entry.paidUsd),
        paidUah: round2(entry.paidUah),
        remainingUsd: round2(entry.remainingUsd),
        remainingUah: round2(entry.remainingUah),
      }));
  }

  function renderDetail() {
    if (state.viewMode !== "month") {
      $detailForm.hidden = true;
      $detailEmpty.hidden = false;
      $detailEmpty.textContent = "У режимі кварталів/років картки агреговані. Для редагування відкрийте режим «Місяці».";
      return;
    }

    if (!state.selectedId) {
      $detailForm.hidden = true;
      $detailEmpty.hidden = false;
      $detailEmpty.textContent = "Виберіть місяць або перший внесок, щоб змінити статус/суму.";
      return;
    }

    const item = idToPayment.get(state.selectedId);
    if (!item) {
      $detailForm.hidden = true;
      $detailEmpty.hidden = false;
      $detailEmpty.textContent = "Виберіть місяць або перший внесок, щоб змінити статус/суму.";
      return;
    }

    const view = buildView(item);
    const override = state.overrides[state.selectedId] || {};
    const selectedRemainingUsd = Math.max(numberOr(view.scheduleUsd, 0) - numberOr(view.paidUsd, 0), 0);
    const selectedRemainingUah = Math.max(numberOr(view.scheduleUah, 0) - numberOr(view.paidUah, 0), 0);

    const projectPrefix = state.selectedProjects.length > 1 ? `[${item.project_name}] ` : "";
    $detailTitle.textContent = `${projectPrefix}${readablePaymentTitle(item)}`;
    $detailMeta.textContent = `По графіку: ${formatCurrency(view.scheduleUsd, "usd")} / ${formatCurrency(
      view.scheduleUah,
      "uah"
    )}, оплачено: ${formatCurrency(view.paidUsd, "usd")} / ${formatCurrency(
      view.paidUah,
      "uah"
    )}, залишок: ${formatCurrency(selectedRemainingUsd, "usd")} / ${formatCurrency(
      selectedRemainingUah,
      "uah"
    )}${item.rate ? `, курс: ${item.rate}` : ""}`;

    $statusInput.value = view.status;
    $statusInput.disabled = isPastPeriod(item);
    $statusInput.title = isPastPeriod(item)
      ? "Минулі місяці автоматично позначаються як оплачені."
      : "";

    $paidUsdInput.value =
      override.paid_usd !== undefined && override.paid_usd !== null ? String(override.paid_usd) : "";
    $paidUahInput.value =
      override.paid_uah !== undefined && override.paid_uah !== null ? String(override.paid_uah) : "";
    $paymentDateInput.value =
      override.payment_date !== undefined && override.payment_date !== null
        ? override.payment_date
        : item.payment_date || "";
    $noteInput.value = override.note || "";

    $paidUsdInput.placeholder = view.paidUsd > 0 ? String(round2(view.paidUsd)) : "Авто";
    $paidUahInput.placeholder = view.paidUah > 0 ? String(round2(view.paidUah)) : "Авто";

    $detailEmpty.hidden = true;
    $detailForm.hidden = false;
  }

  function getSelectedPayments() {
    ensureProjectSelection();
    const selected = new Set(state.selectedProjects);
    return payments.filter((item) => selected.has(item.project_key));
  }

  function getVisiblePayments() {
    let list = getSelectedPayments();

    if (state.viewMode !== "year" && state.selectedYear !== "all") {
      const year = Number(state.selectedYear);
      list = list.filter((item) => getPaymentYear(item) === year);
    }
    if (state.selectedMonth !== "all") {
      if (state.viewMode === "month") {
        const month = Number(state.selectedMonth);
        list = list.filter((item) => getPaymentMonth(item) === month);
      } else if (state.viewMode === "quarter") {
        if (state.selectedMonth.includes("-Q")) {
          const [yearText, quarterText] = state.selectedMonth.split("-Q");
          const targetYear = Number(yearText);
          const targetQuarter = Number(quarterText);
          list = list.filter((item) => {
            const year = getPaymentYear(item);
            const quarter = quarterOfMonth(getPaymentMonth(item) || 1);
            return year === targetYear && quarter === targetQuarter;
          });
        } else {
          const targetQuarter = Number(String(state.selectedMonth).replace("Q", ""));
          list = list.filter((item) => quarterOfMonth(getPaymentMonth(item) || 1) === targetQuarter);
        }
      } else if (state.viewMode === "year") {
        const targetYear = Number(state.selectedMonth);
        list = list.filter((item) => getPaymentYear(item) === targetYear);
      }
    }

    if (state.search) {
      list = list.filter((item) => {
        const view = buildView(item);
        const text = [
          item.project_name,
          readablePaymentTitle(item),
          item.due_label,
          view.paymentDate,
          STATUS_LABELS[view.status],
          view.note,
        ]
          .join(" ")
          .toLowerCase();
        return text.includes(state.search);
      });
    }

    return list.sort((a, b) => {
      const da = parseDate(a.due_date);
      const db = parseDate(b.due_date);
      if (da && db) return da - db;
      return String(a.uid).localeCompare(String(b.uid), "uk");
    });
  }

  function normalizeStatus(rawStatus, item) {
    let status = String(rawStatus || "unpaid").trim().toLowerCase();
    if (status === "partial") status = "unpaid";
    if (status !== "paid" && status !== "unpaid" && status !== "early") {
      status = "unpaid";
    }
    if (isPastPeriod(item)) {
      return "paid";
    }
    return status;
  }

  function isPastPeriod(item) {
    const dueDate = parseDate(item.due_date);
    if (dueDate) return dueDate < startOfCurrentMonth;
    const year = getPaymentYear(item);
    const month = getPaymentMonth(item) || 1;
    const currentYear = startOfCurrentMonth.getFullYear();
    const currentMonth = startOfCurrentMonth.getMonth() + 1;
    if (year < currentYear) return true;
    if (year === currentYear && month < currentMonth) return true;
    return false;
  }

  function isCurrentOrFuturePeriod(item) {
    return !isPastPeriod(item);
  }

  function getEarliestPayableRow(rows) {
    if (!rows || !rows.length) return null;
    const sorted = [...rows].sort((a, b) => {
      const da = parseDate(a.payment.due_date);
      const db = parseDate(b.payment.due_date);
      if (da && db) return da - db;
      return String(a.payment.uid).localeCompare(String(b.payment.uid), "uk");
    });
    return sorted[0] || null;
  }

  function buildView(item) {
    const key = String(item.uid);
    const override = state.overrides[key] || {};

    const status = normalizeStatus(override.status || item.status || "unpaid", item);
    const paymentDate = normalizeDateString(
      override.payment_date !== undefined && override.payment_date !== null
        ? override.payment_date
        : item.payment_date
    );
    const note = override.note || "";

    const explicitUsd = toNumber(override.paid_usd);
    const explicitUah = toNumber(override.paid_uah);

    let paidUsd = explicitUsd;
    let paidUah = explicitUah;

    if (paidUsd === null && paidUah !== null) {
      paidUsd = paidUah / effectiveRate(item);
    }
    if (paidUah === null && paidUsd !== null) {
      paidUah = usdToUah(item, paidUsd);
    }

    if (paidUsd === null) {
      if (status === "paid") {
        paidUsd = numberOr(item.fact_usd, numberOr(item.schedule_usd, 0));
      } else if (status === "early") {
        paidUsd = 0;
      } else {
        paidUsd = 0;
      }
    }

    if (paidUah === null) {
      if (status === "paid") {
        if (item.fact_uah !== null && item.fact_usd !== null && roughlyEqual(paidUsd, item.fact_usd)) {
          paidUah = item.fact_uah;
        } else {
          paidUah = usdToUah(item, paidUsd);
        }
      } else if (status === "early") {
        paidUah = 0;
      } else {
        paidUah = 0;
      }
    }

    if (status === "unpaid") {
      if (explicitUsd === null) paidUsd = 0;
      if (explicitUah === null) paidUah = 0;
    }

    const amountUsd = numberOr(item.fact_usd, numberOr(item.schedule_usd, 0));
    const amountUah =
      item.fact_uah !== null ? item.fact_uah : usdToUah(item, amountUsd);

    const scheduleUsd = numberOr(item.schedule_usd, item.fact_usd);
    const scheduleUah =
      item.fact_uah !== null && item.fact_usd !== null && scheduleUsd === item.fact_usd
        ? item.fact_uah
        : usdToUah(item, scheduleUsd);

    const dueDate = parseDate(item.due_date);
    const overdue = Boolean(dueDate && dueDate < today && status === "unpaid");

    return {
      status,
      paymentDate,
      note,
      paidUsd: round2(paidUsd),
      paidUah: round2(paidUah),
      amountUsd: round2(amountUsd),
      amountUah: round2(amountUah),
      scheduleUsd: round2(numberOr(scheduleUsd, 0)),
      scheduleUah: round2(numberOr(scheduleUah, 0)),
      dueDate,
      overdue,
    };
  }

  function cycleStatus(id) {
    if (!id) return;
    const item = idToPayment.get(String(id));
    if (!item) return;
    const current = buildView(item).status;
    const idx = STATUS_ORDER.indexOf(current);
    const next = STATUS_ORDER[(idx + 1) % STATUS_ORDER.length];
    updateOverride(id, { status: next });
    state.selectedId = String(id);
    persistSettings();
    renderAll();
  }

  function updateOverride(id, patch) {
    const key = String(id);
    const prev = state.overrides[key] || {};
    const next = { ...prev };

    if (patch.status) next.status = patch.status;
    if (patch.paid_usd !== undefined) {
      if (patch.paid_usd === null || Number.isNaN(patch.paid_usd)) delete next.paid_usd;
      else next.paid_usd = round2(patch.paid_usd);
    }
    if (patch.paid_uah !== undefined) {
      if (patch.paid_uah === null || Number.isNaN(patch.paid_uah)) delete next.paid_uah;
      else next.paid_uah = round2(patch.paid_uah);
    }
    if (patch.payment_date !== undefined) {
      if (!patch.payment_date) delete next.payment_date;
      else next.payment_date = patch.payment_date;
    }
    if (patch.note !== undefined) {
      if (!patch.note) delete next.note;
      else next.note = patch.note;
    }

    const base = idToPayment.get(key);
    if (base && next.status === normalizeStatus(base.status, base)) delete next.status;

    if (!Object.keys(next).length) {
      delete state.overrides[key];
    } else {
      state.overrides[key] = next;
    }

    persistOverrides();
  }

  function computeDefaultFx() {
    const rates = payments.map((p) => p.rate).filter((v) => typeof v === "number" && v > 0);
    if (rates.length) return round4(rates.reduce((sum, v) => sum + v, 0) / rates.length);

    const implied = payments
      .map((p) => {
        if (numberOr(p.fact_uah, 0) > 0 && numberOr(p.fact_usd, 0) > 0) {
          return p.fact_uah / p.fact_usd;
        }
        return null;
      })
      .filter((v) => typeof v === "number" && Number.isFinite(v) && v > 0);
    if (implied.length) return round4(implied.reduce((sum, v) => sum + v, 0) / implied.length);
    return 42;
  }

  function usdToUah(item, usd) {
    const amount = numberOr(usd, 0);
    return amount * effectiveRate(item);
  }

  function effectiveRate(item) {
    return numberOr(item.rate, state.fxRate > 0 ? state.fxRate : defaultFx);
  }

  function getPaymentYear(item) {
    if (item.period_year) return item.period_year;
    if (item.due_date) return Number(item.due_date.slice(0, 4));
    return fallbackYear;
  }

  function getPaymentMonth(item) {
    if (item.period_month) return item.period_month;
    if (item.due_date) return Number(item.due_date.slice(5, 7));
    return null;
  }

  function readablePaymentTitle(item) {
    if (item.flags && item.flags.initial) return "Перший внесок";
    const date = parseDate(item.due_date);
    if (!date) return item.due_label || `Платіж #${item.id}`;
    return `${monthLong(date.getMonth() + 1)} ${date.getFullYear()}`;
  }

  function quarterOfMonth(month) {
    return Math.floor((Number(month) - 1) / 3) + 1;
  }

  function describeSelectedPeriod(selectedValue, selectedYear, mode) {
    if (selectedValue === "all") {
      if (mode === "month") return "усі місяці";
      if (mode === "quarter") return "усі квартали";
      return "усі роки";
    }
    if (mode === "month") {
      return monthLong(Number(selectedValue));
    }
    if (mode === "quarter") {
      if (String(selectedValue).includes("-Q")) {
        const [yearText, quarterText] = String(selectedValue).split("-Q");
        return `Q${quarterText} ${yearText}`;
      }
      const yearLabel = selectedYear === "all" ? "" : ` ${selectedYear}`;
      return `Q${String(selectedValue).replace("Q", "")}${yearLabel}`;
    }
    return String(selectedValue);
  }

  function monthLong(month) {
    const label = new Intl.DateTimeFormat("uk-UA", { month: "long" }).format(
      new Date(2026, month - 1, 1)
    );
    return label.charAt(0).toUpperCase() + label.slice(1);
  }

  function parseDate(value) {
    if (!value || typeof value !== "string") return null;
    const out = new Date(`${value}T00:00:00`);
    if (Number.isNaN(out.getTime())) return null;
    return out;
  }

  function normalizeDateString(value) {
    if (!value) return "";
    return String(value).trim();
  }

  function formatDate(value) {
    const d = parseDate(value);
    if (!d) return value || "-";
    const dd = String(d.getDate()).padStart(2, "0");
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const yyyy = d.getFullYear();
    return `${dd}.${mm}.${yyyy}`;
  }

  function formatCurrency(value, currencyKey) {
    const code = currencyKey === "uah" ? "UAH" : "USD";
    return new Intl.NumberFormat("uk-UA", {
      style: "currency",
      currency: code,
      maximumFractionDigits: 2,
    }).format(numberOr(value, 0));
  }

  function numberOr(value, fallback) {
    return typeof value === "number" && Number.isFinite(value) ? value : fallback;
  }

  function toNumber(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const text = String(value).trim();
    if (!text) return null;
    const parsed = Number(text.replace(/\s+/g, "").replace(",", "."));
    return Number.isFinite(parsed) ? parsed : null;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function round2(value) {
    return Math.round(numberOr(value, 0) * 100) / 100;
  }

  function round4(value) {
    return Math.round(numberOr(value, 0) * 10000) / 10000;
  }

  function isSettledStatus(status) {
    return status === "paid";
  }

  function roughlyEqual(a, b) {
    return Math.abs(numberOr(a, 0) - numberOr(b, 0)) < 0.01;
  }

  function persistOverrides() {
    localStorage.setItem(STORAGE_OVERRIDES, JSON.stringify(state.overrides));
    persistSnapshot();
  }

  function persistSettings() {
    const payload = {
      currency: state.currency,
      selectedProjects: state.selectedProjects,
      viewMode: state.viewMode,
      selectedYear: state.selectedYear,
      selectedMonth: state.selectedMonth,
      selectedId: state.selectedId,
      fxRate: state.fxRate,
    };
    localStorage.setItem(STORAGE_SETTINGS, JSON.stringify(payload));
    persistSnapshot();
  }

  function hydrateSettings() {
    const settings = loadJson(STORAGE_SETTINGS, {});
    applySettings(settings);
    applySnapshot(loadJson(STORAGE_SNAPSHOT, null));
    migrateLegacyOverrides();

    $searchInput.value = "";
    $fxRateInput.value = state.fxRate.toFixed(4);
    if ($earlyTariffInput) {
      $earlyTariffInput.value = "0";
    }
  }

  function buildSnapshot() {
    const snapshot = {
      version: 4,
      projectKey: PROJECT_KEY,
      sourceProjects: state.selectedProjects,
      savedAt: new Date().toISOString(),
      settings: {
        currency: state.currency,
        selectedProjects: state.selectedProjects,
        viewMode: state.viewMode,
        selectedYear: state.selectedYear,
        selectedMonth: state.selectedMonth,
        selectedId: state.selectedId,
        fxRate: state.fxRate,
      },
      overrides: state.overrides,
    };
    return snapshot;
  }

  function applySettings(settings) {
    if (!settings || typeof settings !== "object") return;
    if (settings.currency === "usd" || settings.currency === "uah") {
      state.currency = settings.currency;
    }
    if (Array.isArray(settings.selectedProjects)) {
      state.selectedProjects = settings.selectedProjects.filter((key) => datasetKeys.includes(String(key)));
      ensureProjectSelection();
    }
    if (settings.viewMode === "month" || settings.viewMode === "quarter" || settings.viewMode === "year") {
      state.viewMode = settings.viewMode;
    }
    if (settings.selectedYear === "all" || years.includes(Number(settings.selectedYear))) {
      state.selectedYear = String(settings.selectedYear);
    }
    if (
      settings.selectedMonth === "all" ||
      typeof settings.selectedMonth === "string" ||
      typeof settings.selectedMonth === "number"
    ) {
      state.selectedMonth = String(settings.selectedMonth);
    }
    if (settings.selectedId && idToPayment.has(String(settings.selectedId))) {
      state.selectedId = String(settings.selectedId);
    }
    if (toNumber(settings.fxRate) && toNumber(settings.fxRate) > 0) {
      state.fxRate = toNumber(settings.fxRate);
    }
  }

  function applySnapshot(snapshot) {
    if (!snapshot || typeof snapshot !== "object") return;
    if (snapshot.settings && typeof snapshot.settings === "object") {
      applySettings(snapshot.settings);
    }
    if (snapshot.overrides && typeof snapshot.overrides === "object") {
      state.overrides = { ...snapshot.overrides };
    }
    if (typeof snapshot.savedAt === "string") {
      state.lastSavedAt = snapshot.savedAt;
    }
  }

  function persistSnapshot() {
    const snapshot = buildSnapshot();
    localStorage.setItem(STORAGE_SNAPSHOT, JSON.stringify(snapshot));
    state.lastSavedAt = snapshot.savedAt;

    let history = loadJson(STORAGE_SNAPSHOT_HISTORY, []);
    if (!Array.isArray(history)) history = [];
    history.unshift(snapshot);
    if (history.length > MAX_SNAPSHOT_HISTORY) {
      history = history.slice(0, MAX_SNAPSHOT_HISTORY);
    }
    localStorage.setItem(STORAGE_SNAPSHOT_HISTORY, JSON.stringify(history));
  }

  function loadJson(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return fallback;
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === "object") return parsed;
      return fallback;
    } catch (error) {
      return fallback;
    }
  }

  function migrateLegacyOverrides() {
    if (Object.keys(state.overrides || {}).length) return;
    const merged = {};
    datasetKeys.forEach((projectKey) => {
      const legacyKey = `payment_calendar_overrides_v1_${projectKey}`;
      const legacy = loadJson(legacyKey, {});
      if (!legacy || typeof legacy !== "object") return;
      Object.keys(legacy).forEach((id) => {
        merged[`${projectKey}:${id}`] = legacy[id];
      });
    });
    if (Object.keys(merged).length) {
      state.overrides = merged;
      localStorage.setItem(STORAGE_OVERRIDES, JSON.stringify(state.overrides));
    }
  }

  function ensureProjectSelection() {
    state.selectedProjects = (state.selectedProjects || []).filter((key) => datasetKeys.includes(key));
    if (!state.selectedProjects.length) {
      state.selectedProjects = [datasetKeys[0]];
    }
  }

  function inferInitialProjectKey(keys) {
    const path = String(window.location.pathname || "").toLowerCase();
    if (path.includes("respublika2") && keys.includes("respublika2")) return "respublika2";
    if (path.includes("nivki") && keys.includes("nivki")) return "nivki";
    if (keys.includes("respublika1")) return "respublika1";
    return keys[0];
  }

  function orderProjectKeys(keys) {
    const preferred = ["respublika1", "respublika2", "nivki"];
    const set = new Set(keys);
    const ordered = preferred.filter((k) => set.has(k));
    keys.forEach((k) => {
      if (!ordered.includes(k)) ordered.push(k);
    });
    return ordered;
  }

  function resolveDatasets() {
    const out = {};
    const primary = window.PAYMENTS_DATA;
    const existing = window.PAYMENTS_DATASETS;
    if (existing && typeof existing === "object") {
      Object.keys(existing).forEach((key) => {
        const item = existing[key];
        if (item && Array.isArray(item.payments)) out[key] = item;
      });
    }
    if (primary && Array.isArray(primary.payments)) {
      const key = primary.project_key || inferKeyFromName(primary.project_name) || "default";
      if (!out[key]) out[key] = primary;
    }
    return out;
  }

  function inferKeyFromName(name) {
    const text = String(name || "").toLowerCase();
    if (text.includes("республіка 1") || text.includes("respublika 1")) return "respublika1";
    if (text.includes("республіка 2") || text.includes("respublika 2")) return "respublika2";
    if (text.includes("нивки") || text.includes("nivki")) return "nivki";
    return "default";
  }

  function safeDomId(value) {
    return String(value || "item")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "");
  }
})();
