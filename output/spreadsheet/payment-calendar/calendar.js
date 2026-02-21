(() => {
  const source = window.PAYMENTS_DATA;
  if (!source || !Array.isArray(source.payments)) {
    document.body.innerHTML = '<p style="padding:24px">Не знайдено дані платежів.</p>';
    return;
  }

  const PROJECT_KEY = normalizeProjectKey(
    source.project_key || source.project_name || source.apartment || source.source_pdf || "default"
  );
  const STORAGE_OVERRIDES = `payment_calendar_overrides_v1_${PROJECT_KEY}`;
  const STORAGE_SETTINGS = `payment_calendar_settings_v1_${PROJECT_KEY}`;
  const STORAGE_SNAPSHOT = `payment_calendar_snapshot_v1_${PROJECT_KEY}`;
  const STORAGE_SNAPSHOT_HISTORY = `payment_calendar_snapshot_history_v1_${PROJECT_KEY}`;
  const MAX_SNAPSHOT_HISTORY = 40;

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

  const $yearSelect = document.getElementById("yearSelect");
  const $monthSelect = document.getElementById("monthSelect");
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

  const payments = source.payments.map((item) => ({ ...item }));
  const idToPayment = new Map(payments.map((item) => [String(item.id), item]));

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

  const state = {
    overrides: loadJson(STORAGE_OVERRIDES, {}),
    currency: "usd",
    selectedYear: "all",
    selectedMonth: "all",
    selectedId: String(payments.length ? payments[0].id : ""),
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

    $yearSelect.addEventListener("change", () => {
      state.selectedYear = $yearSelect.value || "all";
      ensureSelectedMonthInRange();
      persistSettings();
      renderAll();
    });

    $monthSelect.addEventListener("change", () => {
      state.selectedMonth = $monthSelect.value || "all";
      persistSettings();
      renderAll();
    });

    $calendarGrid.addEventListener("click", (event) => {
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
    renderCurrencyControls();
    renderSummary();
    renderPeriodFilters();
    syncSelectedWithVisible();
    renderCalendar();
    renderDetail();
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
    const rows = payments
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

    const earlyBaseRow = getEarliestPayableRow(payableFutureRows);
    const earlyUnitUsd = earlyBaseRow ? earlyBaseRow.view.scheduleUsd : 0;
    const earlyUnitUah = earlyBaseRow ? earlyBaseRow.view.scheduleUah : 0;
    state.earlyUnitUsd = earlyUnitUsd;
    state.earlyUnitUah = earlyUnitUah;

    const futureCount = payableFutureRows.length;
    const remainingEarlyUsd = futureCount * earlyUnitUsd;
    const remainingEarlyUah = futureCount * earlyUnitUah;
    const remainingEarly = state.currency === "usd" ? remainingEarlyUsd : remainingEarlyUah;
    const total = paid + remainingPlan;

    const paidCount = paidRows.length;
    const allCount = rows.length;
    const baseLabel = earlyBaseRow
      ? readablePaymentTitle(earlyBaseRow.payment)
      : "немає базового місяця";

    $summaryPaid.textContent = formatCurrency(paid, state.currency);
    $summaryRemainingPlan.textContent = formatCurrency(remainingPlan, state.currency);
    $summaryRemainingEarly.textContent = formatCurrency(remainingEarly, state.currency);
    $summaryRemainingEarlyLabel.textContent = `Залишок достроково (${futureCount} платежів, база: ${baseLabel})`;
    $summaryTotal.textContent = formatCurrency(total, state.currency);
    $summaryCount.textContent = `${paidCount}/${allCount} оплачено, майбутніх неоплачених: ${futureUnpaidRows.length}`;
  }

  function renderPeriodFilters() {
    const yearCounts = new Map();
    payments.forEach((item) => {
      const year = getPaymentYear(item);
      yearCounts.set(year, (yearCounts.get(year) || 0) + 1);
    });

    if (state.selectedYear !== "all" && !years.includes(Number(state.selectedYear))) {
      state.selectedYear = "all";
    }

    const yearOptions = [
      `<option value="all"${state.selectedYear === "all" ? " selected" : ""}>Всі роки (${payments.length})</option>`,
      ...years.map((year) => {
        const selected = state.selectedYear === String(year) ? " selected" : "";
        return `<option value="${year}"${selected}>${year} (${yearCounts.get(year) || 0})</option>`;
      }),
    ];
    $yearSelect.innerHTML = yearOptions.join("");

    ensureSelectedMonthInRange();
    const monthOptionsData = getMonthCounts(state.selectedYear);
    const totalInYear = monthOptionsData.reduce((sum, entry) => sum + entry.count, 0);
    const monthOptions = [
      `<option value="all"${state.selectedMonth === "all" ? " selected" : ""}>Всі місяці (${totalInYear})</option>`,
      ...monthOptionsData.map((entry) => {
        const selected = state.selectedMonth === String(entry.month) ? " selected" : "";
        return `<option value="${entry.month}"${selected}>${escapeHtml(
          monthLong(entry.month)
        )} (${entry.count})</option>`;
      }),
    ];
    $monthSelect.innerHTML = monthOptions.join("");

    const visible = getVisiblePayments();
    const yearLabel = state.selectedYear === "all" ? "усі роки" : state.selectedYear;
    const monthLabel = state.selectedMonth === "all" ? "усі місяці" : monthLong(Number(state.selectedMonth));
    const searchLabel = state.search ? `, пошук: “${state.search}”` : "";
    $monthInfo.textContent = `Показано ${visible.length} періодів: ${yearLabel}, ${monthLabel}${searchLabel}`;
  }

  function getMonthCounts(yearValue) {
    const bucket = new Map();
    payments.forEach((item) => {
      const year = getPaymentYear(item);
      if (yearValue !== "all" && year !== Number(yearValue)) return;
      const month = getPaymentMonth(item);
      if (!month) return;
      bucket.set(month, (bucket.get(month) || 0) + 1);
    });

    return [...bucket.entries()]
      .map(([month, count]) => ({ month, count }))
      .sort((a, b) => a.month - b.month);
  }

  function ensureSelectedMonthInRange() {
    if (state.selectedMonth === "all") return;
    const exists = getMonthCounts(state.selectedYear).some(
      (entry) => String(entry.month) === String(state.selectedMonth)
    );
    if (!exists) {
      state.selectedMonth = "all";
    }
  }

  function syncSelectedWithVisible() {
    const visible = getVisiblePayments();
    if (!visible.length) {
      state.selectedId = "";
      return;
    }
    if (!state.selectedId || !visible.some((item) => String(item.id) === state.selectedId)) {
      state.selectedId = String(visible[0].id);
    }
  }

  function renderCalendar() {
    const visible = getVisiblePayments();

    if (!visible.length) {
      $calendarGrid.innerHTML = '<div class="empty-note">Нічого не знайдено за поточним фільтром.</div>';
      return;
    }

    const cards = visible.map((item) => {
      const view = buildView(item);
      const selected = state.selectedId === String(item.id) ? "selected" : "";
      const overdue = view.overdue ? "overdue" : "";
      const amount = state.currency === "usd" ? view.amountUsd : view.amountUah;
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
        <article class="payment-card ${selected} ${overdue}" id="payment-${item.id}" data-id="${item.id}">
          <div class="card-top">
            <h3 class="card-title">${escapeHtml(readablePaymentTitle(item))}</h3>
            <button class="status-toggle" type="button" data-id="${item.id}" title="Змінити статус">
              <span class="status-pill ${view.status}">${escapeHtml(STATUS_LABELS[view.status])}</span>
            </button>
          </div>
          <div class="card-amount">${formatCurrency(amount, state.currency)}</div>
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

  function renderDetail() {
    if (!state.selectedId) {
      $detailForm.hidden = true;
      $detailEmpty.hidden = false;
      return;
    }

    const item = idToPayment.get(state.selectedId);
    if (!item) {
      $detailForm.hidden = true;
      $detailEmpty.hidden = false;
      return;
    }

    const view = buildView(item);
    const override = state.overrides[state.selectedId] || {};
    const selectedRemainingUsd = Math.max(numberOr(view.scheduleUsd, 0) - numberOr(view.paidUsd, 0), 0);
    const selectedRemainingUah = Math.max(numberOr(view.scheduleUah, 0) - numberOr(view.paidUah, 0), 0);

    $detailTitle.textContent = readablePaymentTitle(item);
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

  function getVisiblePayments() {
    let list = [...payments];

    if (state.selectedYear !== "all") {
      const year = Number(state.selectedYear);
      list = list.filter((item) => getPaymentYear(item) === year);
    }
    if (state.selectedMonth !== "all") {
      const month = Number(state.selectedMonth);
      list = list.filter((item) => getPaymentMonth(item) === month);
    }

    if (state.search) {
      list = list.filter((item) => {
        const view = buildView(item);
        const text = [
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
      return a.id - b.id;
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
      return a.payment.id - b.payment.id;
    });
    return sorted[0] || null;
  }

  function buildView(item) {
    const key = String(item.id);
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
    const fromSummary =
      numberOr(getSummaryValue("total", "fact_uah"), 0) /
      numberOr(getSummaryValue("total", "fact_usd"), 1);
    if (Number.isFinite(fromSummary) && fromSummary > 0) return round4(fromSummary);

    const rates = payments.map((p) => p.rate).filter((v) => typeof v === "number" && v > 0);
    if (!rates.length) return 42;
    return round4(rates.reduce((sum, v) => sum + v, 0) / rates.length);
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

  function getSummaryValue(group, key) {
    const summary = source.summary;
    if (!summary || typeof summary !== "object") return null;
    const section = summary[group];
    if (!section || typeof section !== "object") return null;
    const value = section[key];
    return typeof value === "number" && Number.isFinite(value) ? value : null;
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

    $searchInput.value = "";
    $fxRateInput.value = state.fxRate.toFixed(4);
    if ($earlyTariffInput) {
      $earlyTariffInput.value = "0";
    }
  }

  function buildSnapshot() {
    const snapshot = {
      version: 3,
      projectKey: PROJECT_KEY,
      sourcePdf: source.source_pdf || "",
      savedAt: new Date().toISOString(),
      settings: {
        currency: state.currency,
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
    if (settings.selectedYear === "all" || years.includes(Number(settings.selectedYear))) {
      state.selectedYear = String(settings.selectedYear);
    }
    if (settings.selectedMonth === "all" || (toNumber(settings.selectedMonth) || 0) >= 1) {
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

  function normalizeProjectKey(input) {
    return (
      String(input || "default")
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "-")
        .replace(/^-+|-+$/g, "")
        .slice(0, 80) || "default"
    );
  }
})();
