(() => {
  const source = window.PAYMENTS_DATA;
  if (!source || !Array.isArray(source.payments)) {
    document.body.innerHTML = '<p style="padding:24px">Не знайдено дані платежів.</p>';
    return;
  }

  const STORAGE_OVERRIDES = "payment_calendar_overrides_v1";
  const STORAGE_SETTINGS = "payment_calendar_settings_v1";
  const STORAGE_SNAPSHOT = "payment_calendar_snapshot_v1";
  const STORAGE_SNAPSHOT_HISTORY = "payment_calendar_snapshot_history_v1";
  const MAX_SNAPSHOT_HISTORY = 40;

  const STATUS_ORDER = ["unpaid", "partial", "paid", "early"];
  const STATUS_LABELS = {
    paid: "Оплачено",
    unpaid: "Не оплачено",
    partial: "Частково",
    early: "Достроково",
  };

  const $summaryPaid = document.getElementById("summaryPaid");
  const $summaryRemainingPlan = document.getElementById("summaryRemainingPlan");
  const $summaryRemainingEarly = document.getElementById("summaryRemainingEarly");
  const $summaryRemainingEarlyLabel = document.getElementById("summaryRemainingEarlyLabel");
  const $summaryTotal = document.getElementById("summaryTotal");
  const $summaryCount = document.getElementById("summaryCount");

  const $yearList = document.getElementById("yearList");
  const $monthLinks = document.getElementById("monthLinks");
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

  const summaryTotals = {
    usd: numberOr(
      getSummaryValue("total", "fact_usd"),
      payments.reduce((sum, item) => sum + numberOr(item.fact_usd, 0), 0)
    ),
    uah: numberOr(
      getSummaryValue("total", "fact_uah"),
      payments.reduce((sum, item) => sum + numberOr(item.fact_uah, 0), 0)
    ),
  };

  const defaultFx = computeDefaultFx();

  const state = {
    overrides: loadJson(STORAGE_OVERRIDES, {}),
    currency: "usd",
    selectedYear: "all",
    selectedId: String(payments.length ? payments[0].id : ""),
    search: "",
    fxRate: defaultFx,
    earlyTariffPercent: 100,
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
      renderCalendar();
    });

    $fxRateInput.addEventListener("change", () => {
      const next = toNumber($fxRateInput.value);
      if (next !== null && next > 0) {
        state.fxRate = next;
        persistSettings();
        renderAll();
      }
    });

    $earlyTariffInput.addEventListener("change", () => {
      const next = toNumber($earlyTariffInput.value);
      if (next !== null) {
        state.earlyTariffPercent = clampTariff(next);
        persistSettings();
        renderAll();
      }
    });

    $yearList.addEventListener("click", (event) => {
      const btn = event.target.closest(".year-btn");
      if (!btn) return;
      state.selectedYear = btn.dataset.year;
      persistSettings();
      renderCalendar();
      renderYearButtons();
      renderMonthLinks();
    });

    $monthLinks.addEventListener("click", (event) => {
      const btn = event.target.closest(".month-link-btn");
      if (!btn) return;
      const target = document.getElementById(btn.dataset.target);
      if (!target) return;
      target.scrollIntoView({ behavior: "smooth", block: "nearest" });
      target.classList.add("selected");
      state.selectedId = target.dataset.id;
      persistSettings();
      renderDetail();
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
      anchor.download = `payment-calendar-backup-${stamp}.json`;
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
    renderYearButtons();
    renderMonthLinks();
    renderCalendar();
    renderDetail();
  }

  function renderCurrencyControls() {
    $currencyButtons.forEach((btn) => {
      btn.classList.toggle("active", btn.dataset.currency === state.currency);
    });
    $fxRateInput.value = state.fxRate.toFixed(4);
    $earlyTariffInput.value = formatTariffPercent(state.earlyTariffPercent);
  }

  function renderSummary() {
    const views = payments.map((payment) => buildView(payment));

    const paid = views.reduce(
      (sum, view) => sum + (state.currency === "usd" ? view.paidUsd : view.paidUah),
      0
    );
    const total = state.currency === "usd" ? summaryTotals.usd : summaryTotals.uah;
    const remainingPlan = Math.max(total - paid, 0);
    const remainingEarly = remainingPlan * tariffCoefficient();
    const saving = Math.max(remainingPlan - remainingEarly, 0);

    const paidCount = views.filter((view) => isSettledStatus(view.status)).length;
    const overdueCount = views.filter((view) => view.overdue).length;

    $summaryPaid.textContent = formatCurrency(paid, state.currency);
    $summaryRemainingPlan.textContent = formatCurrency(remainingPlan, state.currency);
    $summaryRemainingEarly.textContent = formatCurrency(remainingEarly, state.currency);
    $summaryRemainingEarlyLabel.textContent = `Залишок достроково (${formatTariffPercent(
      state.earlyTariffPercent
    )}%)`;
    $summaryTotal.textContent = formatCurrency(total, state.currency);
    $summaryCount.textContent = `${paidCount}/${views.length} оплачено${
      overdueCount ? `, прострочено: ${overdueCount}` : ""
    }${saving > 0 ? `, економія: ${formatCurrency(saving, state.currency)}` : ""}`;
  }

  function renderYearButtons() {
    const counts = new Map();
    payments.forEach((item) => {
      const year = getPaymentYear(item);
      counts.set(year, (counts.get(year) || 0) + 1);
    });

    const buttons = [
      `<button class="year-btn ${state.selectedYear === "all" ? "active" : ""}" data-year="all">Всі (${payments.length})</button>`,
      ...years.map((year) => {
        const active = state.selectedYear === String(year) ? "active" : "";
        return `<button class="year-btn ${active}" data-year="${year}">${year} (${counts.get(year) || 0})</button>`;
      }),
    ];

    $yearList.innerHTML = buttons.join("");
  }

  function renderMonthLinks() {
    const visible = getVisiblePayments();
    const links = [];
    const seen = new Set();

    visible.forEach((item) => {
      const year = getPaymentYear(item);
      const month = getPaymentMonth(item);
      if (!month) return;
      const key = `${year}-${String(month).padStart(2, "0")}`;
      if (seen.has(key)) return;
      seen.add(key);

      const label = state.selectedYear === "all" ? `${shortMonth(month)} ${year}` : shortMonth(month);
      links.push(
        `<button class="month-link-btn" type="button" data-target="payment-${item.id}">${escapeHtml(label)}</button>`
      );
    });

    $monthLinks.innerHTML = links.length ? links.join("") : "";
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
      const scheduled = state.currency === "usd" ? view.scheduleUsd : view.scheduleUah;
      const plannedText = Number.isFinite(scheduled)
        ? `По графіку: ${formatCurrency(scheduled, state.currency)}`
        : "По графіку: -";
      const paidByCard = state.currency === "usd" ? view.paidUsd : view.paidUah;
      const remainingForCard = Math.max(numberOr(scheduled, 0) - numberOr(paidByCard, 0), 0);
      const earlyForCard = remainingForCard * tariffCoefficient();
      const earlyText =
        remainingForCard > 0
          ? `Достроково (${formatTariffPercent(state.earlyTariffPercent)}%): ${formatCurrency(
              earlyForCard,
              state.currency
            )}`
          : "";
      const paymentDateText = view.paymentDate ? `Оплата: ${formatDate(view.paymentDate)}` : "Дата оплати не вказана";
      const note = view.note || (item.flags && item.flags.early ? "Позначено як достроково в джерелі." : "");

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
          ${earlyText ? `<p class="card-meta">${earlyText}</p>` : ""}
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
    const selectedSchedule = state.currency === "usd" ? view.scheduleUsd : view.scheduleUah;
    const selectedPaid = state.currency === "usd" ? view.paidUsd : view.paidUah;
    const selectedRemaining = Math.max(numberOr(selectedSchedule, 0) - numberOr(selectedPaid, 0), 0);
    const selectedEarly = selectedRemaining * tariffCoefficient();

    $detailTitle.textContent = readablePaymentTitle(item);
    $detailMeta.textContent = `По графіку: ${formatCurrency(
      selectedSchedule,
      state.currency
    )}, залишок: ${formatCurrency(selectedRemaining, state.currency)}, достроково (${formatTariffPercent(
      state.earlyTariffPercent
    )}%): ${formatCurrency(selectedEarly, state.currency)}${item.rate ? `, курс: ${item.rate}` : ""}`;

    $statusInput.value = view.status;

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

  function buildView(item) {
    const key = String(item.id);
    const override = state.overrides[key] || {};

    const status = override.status || item.status || "unpaid";
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
        paidUsd = numberOr(item.schedule_usd, numberOr(item.fact_usd, 0)) * tariffCoefficient();
      } else if (status === "partial") {
        paidUsd = numberOr(item.base_paid_usd, 0);
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
        paidUah = usdToUah(item, paidUsd);
      } else if (status === "partial") {
        if (
          item.base_paid_uah !== null &&
          item.base_paid_usd !== null &&
          roughlyEqual(paidUsd, item.base_paid_usd)
        ) {
          paidUah = item.base_paid_uah;
        } else {
          paidUah = usdToUah(item, paidUsd);
        }
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
    const overdue = Boolean(dueDate && dueDate < today && (status === "unpaid" || status === "partial"));

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
    if (base && next.status === base.status) delete next.status;

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

  function shortMonth(month) {
    return new Intl.DateTimeFormat("uk-UA", { month: "short" })
      .format(new Date(2026, month - 1, 1))
      .replace(".", "");
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

  function clampTariff(value) {
    const numeric = numberOr(toNumber(value), 100);
    if (!Number.isFinite(numeric)) return 100;
    if (numeric < 1) return 1;
    if (numeric > 200) return 200;
    return round2(numeric);
  }

  function tariffCoefficient() {
    return clampTariff(state.earlyTariffPercent) / 100;
  }

  function formatTariffPercent(value) {
    const normalized = clampTariff(value);
    return Number.isInteger(normalized) ? String(normalized) : normalized.toFixed(2).replace(/\.?0+$/, "");
  }

  function isSettledStatus(status) {
    return status === "paid" || status === "early";
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
      selectedId: state.selectedId,
      fxRate: state.fxRate,
      earlyTariffPercent: state.earlyTariffPercent,
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
    $earlyTariffInput.value = formatTariffPercent(state.earlyTariffPercent);
  }

  function buildSnapshot() {
    const snapshot = {
      version: 2,
      sourcePdf: source.source_pdf || "",
      savedAt: new Date().toISOString(),
      settings: {
        currency: state.currency,
        selectedYear: state.selectedYear,
        selectedId: state.selectedId,
        fxRate: state.fxRate,
        earlyTariffPercent: state.earlyTariffPercent,
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
    if (settings.selectedId && idToPayment.has(String(settings.selectedId))) {
      state.selectedId = String(settings.selectedId);
    }
    if (toNumber(settings.fxRate) && toNumber(settings.fxRate) > 0) {
      state.fxRate = toNumber(settings.fxRate);
    }
    if (toNumber(settings.earlyTariffPercent) && toNumber(settings.earlyTariffPercent) > 0) {
      state.earlyTariffPercent = clampTariff(toNumber(settings.earlyTariffPercent));
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
})();
