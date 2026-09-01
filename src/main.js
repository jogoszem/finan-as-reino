import { readFirstSheetRows } from "./xlsx-reader.js";
import { excludedParticipants, manualRecords } from "./manual-records.js";
import { registeredMinistryEntries } from "./registered-ministry-entries.js";
import "./style.css";

const sourceWorkbook = new URL("../Libras.xlsx", import.meta.url);
const ministryFinanceJson = new URL("../finançassinaisdoreino.json", import.meta.url);
const brandLogo = new URL("./logos/logo.svg", import.meta.url);
const crownLogo = new URL("./logos/coroa.svg", import.meta.url);

const yearEndPlan = {
  months: ["Setembro", "Outubro", "Novembro", "Dezembro"],
  fixedExpenses: [
    { label: "Café e comida", monthly: 500, reason: "Alimentação prevista para as atividades até dezembro." },
    { label: "Uber essencial", monthly: 200, reason: "Deslocamentos essenciais previstos até dezembro." },
    { label: "Ajuda Éder", monthly: 150, reason: "Ajuda mensal prevista para Éder." },
  ],
  courseExpenses: 4000,
  certainCourseCredit: 2200,
  uncertainDelinquency: 1600,
  alexMonthlyContribution: 500,
};

const state = {
  records: [],
  filtered: [],
  ministryRecords: [],
  filteredMinistry: [],
  page: 1,
  perPage: 8,
  search: "",
  donor: "Todos",
  method: "Todos",
  status: "Todos",
  installment: "Todas",
  dateFrom: "",
  dateTo: "",
  participantSearch: "",
  participantName: "Todos",
  participantStatus: "Todos",
  participantPage: 1,
  participantPerPage: 8,
  ministrySearch: "",
  ministryKind: "Todos",
  ministryType: "Todos",
  ministryCategory: "Todas",
  ministryDateFrom: "",
  ministryDateTo: "",
  presentationPage: 0,
  activeView: "overview",
  sourceName: "Libras.xlsx",
};

const viewTitles = {
  presentation: "Apresentação financeira",
  overview: "Auditoria financeira",
  participants: "Curso e alunos",
  cashflow: "Fluxo e projeções",
  ministry: "Ministério e despesas",
};

const icons = {
  overview: `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="3" width="7" height="7" rx="2"/><rect x="14" y="3" width="7" height="7" rx="2"/><rect x="3" y="14" width="7" height="7" rx="2"/><rect x="14" y="14" width="7" height="7" rx="2"/></svg>`,
  chart: `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 19V9m6 10V5m6 14v-7m4 7H2" stroke-linecap="round"/></svg>`,
  people: `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M16 20v-1.5A3.5 3.5 0 0 0 12.5 15h-5A3.5 3.5 0 0 0 4 18.5V20M9.75 11a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7ZM17 11a3 3 0 0 0 0-6m1.5 10.5A3.5 3.5 0 0 1 22 19v1" stroke-linecap="round"/></svg>`,
  rows: `<svg class="nav-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M8 6h13M8 12h13M8 18h13M3.5 6h.01M3.5 12h.01M3.5 18h.01" stroke-linecap="round"/></svg>`,
  upload: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5M5 14v5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-5" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  download: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 4v12m0 0 4.5-4.5M12 16l-4.5-4.5M5 20h14" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  search: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="11" cy="11" r="7"/><path d="m16.5 16.5 4 4" stroke-linecap="round"/></svg>`,
  money: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="5" width="18" height="14" rx="3"/><path d="M7 9h.01M17 15h.01M12 15a3 3 0 1 0 0-6 3 3 0 0 0 0 6Z" stroke-linecap="round"/></svg>`,
  check: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="m8 12 2.6 2.6L16.5 9" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  bank: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m3 9 9-5 9 5M5 10v7m4-7v7m6-7v7m4-7v7M3 20h18" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  clock: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  spark: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m12 3 1.5 4.5L18 9l-4.5 1.5L12 15l-1.5-4.5L6 9l4.5-1.5L12 3Zm6 11 .8 2.2L21 17l-2.2.8L18 20l-.8-2.2L15 17l2.2-.8L18 14Z" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  arrow: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m9 18 6-6-6-6" stroke-linecap="round" stroke-linejoin="round"/></svg>`,
  menu: `<svg class="icon mobile-menu-open-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9"><path d="M4 7h16M4 12h16M4 17h16" stroke-linecap="round"/></svg>`,
  close: `<svg class="icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m7 7 10 10M17 7 7 17" stroke-linecap="round"/></svg>`,
};

const currency = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
});

const number = new Intl.NumberFormat("pt-BR");

function toNumber(value) {
  if (typeof value === "number") return value;
  if (!value) return 0;
  let normalized = String(value)
    .replace(/R\$\s?/g, "")
    .replace(/\s/g, "")
    .trim();
  if (normalized.includes(",")) {
    normalized = normalized.replace(/\./g, "").replace(",", ".");
  }
  return Number.parseFloat(normalized) || 0;
}

function readMinistryFinance(jsonData) {
  const sourceRecords = Array.isArray(jsonData) ? jsonData : jsonData?.records;
  if (!Array.isArray(sourceRecords)) throw new Error("O arquivo finançassinaisdoreino.json não contém uma lista válida de lançamentos.");

  const registeredIds = new Set(sourceRecords.map((record) => String(record.id || "")));
  const allSourceRecords = [
    ...sourceRecords,
    ...registeredMinistryEntries.filter((record) => !registeredIds.has(record.id)),
  ];

  const records = allSourceRecords
    .map((row, index) => ({
      id: String(row.id || `MINISTRY-${index + 1}`).trim(),
      financeKind: String(row.finance_kind || "Não informado").trim(),
      type: String(row.type || "Não informado").trim(),
      date: String(row.date || "").trim(),
      amount: toNumber(row.amount),
      category: String(row.category || "Sem categoria").trim(),
      eventTitle: String(row.event_title || "").trim(),
      name: String(row.name || "").trim(),
      specification: String(row.specification || "").trim(),
      description: String(row.description || "").trim(),
      paymentMethod: String(row.payment_method || "Não informado").trim(),
      createdBy: String(row.created_by_name || "Não informado").trim(),
      createdAt: String(row.created_at || "").trim(),
      paymentDetails: String(row.payment_details || "").trim(),
    }))
    .filter((record) => record.date && record.amount > 0)
    .sort((a, b) => b.date.localeCompare(a.date) || a.category.localeCompare(b.category, "pt-BR"));

  if (!records.length) throw new Error("Não encontrei lançamentos válidos em finançassinaisdoreino.json.");
  state.ministryRecords = records;
}

function summarizeMinistry(records) {
  const totalByType = (type) => records
    .filter((record) => record.type.toLocaleLowerCase("pt-BR") === type)
    .reduce((total, record) => total + record.amount, 0);
  const expensesByKind = (kind) => records
    .filter((record) => record.financeKind.toLocaleLowerCase("pt-BR") === kind && record.type.toLocaleLowerCase("pt-BR") === "saida")
    .reduce((total, record) => total + record.amount, 0);
  const entries = totalByType("entrada");
  const exits = totalByType("saida");
  return {
    count: records.length,
    entries,
    exits,
    balance: entries - exits,
    classes: expensesByKind("aulas"),
    general: expensesByKind("geral"),
  };
}

function formatIsoDate(value) {
  const [year, month, day] = String(value || "").split("-");
  return year && month && day ? `${day}/${month}/${year}` : "—";
}

function ministryPeriodLabel(records) {
  const dates = records
    .map((record) => record.date ? new Date(`${record.date}T12:00:00`) : null)
    .filter((date) => date && !Number.isNaN(date.getTime()))
    .sort((a, b) => a - b);
  if (!dates.length) return "Período não informado";
  const formatMonth = (date) => {
    const label = new Intl.DateTimeFormat("pt-BR", { month: "short" }).format(date).replace(".", "");
    return `${label.charAt(0).toUpperCase()}${label.slice(1)}`;
  };
  const first = dates[0];
  const last = dates.at(-1);
  if (first.getFullYear() === last.getFullYear()) {
    return `${formatMonth(first)} — ${formatMonth(last)}/${last.getFullYear()}`;
  }
  return `${formatMonth(first)}/${first.getFullYear()} — ${formatMonth(last)}/${last.getFullYear()}`;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value;
  const [datePart] = String(value).split(" ");
  const [day, month, year] = datePart.split("/").map(Number);
  if (!day || !month || !year) return null;
  return new Date(year, month - 1, day, 12);
}

function getStatus(record) {
  if (record.credited > 0) return "Creditado";
  if (record.paid > 0) return "A creditar";
  const dueDate = parseDate(record.dueDate);
  if (dueDate && dueDate < new Date()) return "Em atraso";
  return "Pendente";
}

function statusClass(status) {
  return {
    Creditado: "status-credited",
    "A creditar": "status-waiting",
    "Em atraso": "status-overdue",
    Pendente: "status-pending",
  }[status];
}

async function readWorkbook(arrayBuffer, name) {
  const rows = await readFirstSheetRows(arrayBuffer);
  const records = rows.slice(1).flatMap((row, index) => {
    const donor = String(row[2] || "").trim();
    const event = String(row[14] || "").trim();
    if (!donor || !event) return [];
    const isExcluded = excludedParticipants.some(
      (name) => name.toLocaleLowerCase("pt-BR") === donor.toLocaleLowerCase("pt-BR"),
    );
    if (isExcluded) return [];

    const record = {
      id: `L${index + 26}`,
      sourceRow: index + 26,
      chargeType: String(row[0] || "Cobrança de Inscrição / Matrícula").trim(),
      method: String(row[1] || "Não informado").trim(),
      donor,
      campaign: String(row[3] || "").trim(),
      dueDate: String(row[4] || "").trim(),
      receivedAt: String(row[5] || "").trim(),
      creditedAt: String(row[6] || "").trim(),
      expectedCreditAt: String(row[7] || "").trim(),
      receivable: toNumber(row[8]),
      paid: toNumber(row[9]),
      credited: toNumber(row[10]),
      fee: toNumber(row[11]),
      installment: String(row[12] || "—").trim(),
      costCenter: String(row[13] || "").trim(),
      event,
      acquisition: String(row[15] || "").trim(),
      sponsorship: String(row[16] || "").trim(),
    };

    record.status = getStatus(record);
    return [record];
  });

  manualRecords.forEach((manualRecord) => {
    const alreadyExists = records.some(
      (record) =>
        record.donor.toLocaleLowerCase("pt-BR") === manualRecord.donor.toLocaleLowerCase("pt-BR") &&
        record.dueDate === manualRecord.dueDate &&
        record.installment === manualRecord.installment &&
        record.receivable === manualRecord.receivable,
    );
    if (alreadyExists) return;
    const record = { ...manualRecord, sourceRow: null };
    record.status = getStatus(record);
    records.push(record);
  });

  if (!records.length) {
    throw new Error("Não encontrei lançamentos no formato esperado.");
  }

  state.records = records;
  state.sourceName = name;
  state.page = 1;
}

function summarize(records) {
  const sum = (key) => records.reduce((total, item) => total + item[key], 0);
  const receivable = sum("receivable");
  const paid = sum("paid");
  const credited = sum("credited");
  const fees = records.reduce((total, item) => total + Math.abs(item.fee), 0);
  const open = Math.max(receivable - paid, 0);
  const awaitingCredit = records
    .filter((item) => item.paid > 0 && item.credited === 0)
    .reduce((total, item) => total + item.paid, 0);
  const people = new Set(records.map((item) => item.donor)).size;
  const paymentRate = receivable ? (paid / receivable) * 100 : 0;
  const overdue = records.reduce((total, record) => {
    const dueDate = parseDate(record.dueDate);
    if (!dueDate || dueDate >= new Date()) return total;
    return total + Math.max(record.receivable - record.paid, 0);
  }, 0);
  const futureDue = Math.max(open - overdue, 0);

  return { receivable, paid, credited, fees, open, overdue, futureDue, awaitingCredit, people, paymentRate };
}

function auditAttributes(key, label) {
  return `data-audit-key="${safe(key)}" role="button" tabindex="0" aria-label="Ver detalhamento de ${safe(label)}"`;
}

function groupCashFlowByMonth(records, ministryRecords) {
  const groups = new Map();
  const ensureGroup = (date) => {
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!groups.has(key)) groups.set(key, { key, date, income: 0, outgoing: 0 });
    return groups.get(key);
  };

  records.forEach((record) => {
    const date = parseDate(record.creditedAt);
    if (!date || record.credited <= 0) return;
    ensureGroup(date).income += record.credited;
  });

  ministryRecords.forEach((record) => {
    const type = record.type.toLocaleLowerCase("pt-BR");
    const kind = record.financeKind.toLocaleLowerCase("pt-BR");
    if (type !== "saida" || kind !== "aulas") return;
    const date = record.date ? new Date(`${record.date}T12:00:00`) : null;
    if (!date || Number.isNaN(date.getTime())) return;
    const group = ensureGroup(date);
    group.outgoing += record.amount;
  });

  return [...groups.values()].sort((a, b) => a.date - b.date);
}

function groupMinistryFlowByMonth(ministryRecords) {
  const groups = new Map();
  ministryRecords.forEach((record) => {
    const date = record.date ? new Date(`${record.date}T12:00:00`) : null;
    if (!date || Number.isNaN(date.getTime())) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!groups.has(key)) groups.set(key, { key, date, income: 0, outgoing: 0 });
    const group = groups.get(key);
    const type = record.type.toLocaleLowerCase("pt-BR");
    if (type === "entrada") group.income += record.amount;
    if (type === "saida") group.outgoing += record.amount;
  });
  return [...groups.values()].sort((a, b) => a.date - b.date);
}

function groupConsolidatedFlowByMonth(records, ministryRecords) {
  const groups = new Map();
  const ensureGroup = (date) => {
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!groups.has(key)) groups.set(key, { key, date, courseIncome: 0, administrativeIncome: 0, outgoing: 0 });
    return groups.get(key);
  };

  records.forEach((record) => {
    const date = parseDate(record.creditedAt);
    if (!date || record.credited <= 0) return;
    ensureGroup(date).courseIncome += record.credited;
  });

  ministryRecords.forEach((record) => {
    const date = record.date ? new Date(`${record.date}T12:00:00`) : null;
    if (!date || Number.isNaN(date.getTime())) return;
    const group = ensureGroup(date);
    const type = record.type.toLocaleLowerCase("pt-BR");
    if (type === "entrada") group.administrativeIncome += record.amount;
    if (type === "saida") group.outgoing += record.amount;
  });

  return [...groups.values()]
    .map((group) => ({
      ...group,
      income: group.courseIncome + group.administrativeIncome,
      net: group.courseIncome + group.administrativeIncome - group.outgoing,
    }))
    .sort((a, b) => a.date - b.date);
}

function groupByMethod(records) {
  const groups = new Map();
  records.forEach((record) => {
    if (!groups.has(record.method)) groups.set(record.method, 0);
    groups.set(record.method, groups.get(record.method) + record.paid);
  });
  return [...groups.entries()]
    .map(([name, value]) => ({ name, value }))
    .sort((a, b) => b.value - a.value);
}

const monthNames = [
  "Janeiro",
  "Fevereiro",
  "Março",
  "Abril",
  "Maio",
  "Junho",
  "Julho",
  "Agosto",
  "Setembro",
  "Outubro",
  "Novembro",
  "Dezembro",
];

// Para este relatório, agosto está encerrado e a projeção começa em setembro de 2026.
const forecastStartsAt = new Date(2026, 8, 1, 12);

function buildCreditTimeline(records) {
  const requestedMonths = [4, 5, 6, 7, 8, 9, 10];
  return requestedMonths.map((month) => {
    const credited = records.reduce((total, record) => {
      const creditedDate = parseDate(record.creditedAt);
      if (!creditedDate || creditedDate.getMonth() + 1 !== month) return total;
      return total + record.credited;
    }, 0);
    const toReconcile = records.reduce((total, record) => {
      const expectedDate = parseDate(record.expectedCreditAt);
      if (
        record.credited > 0 ||
        record.paid <= 0 ||
        !expectedDate ||
        expectedDate >= forecastStartsAt ||
        expectedDate.getMonth() + 1 !== month
      ) return total;
      return total + record.paid;
    }, 0);
    const forecast = records.reduce((total, record) => {
      const expectedDate = parseDate(record.expectedCreditAt);
      if (
        record.credited > 0 ||
        !expectedDate ||
        expectedDate < forecastStartsAt ||
        expectedDate.getMonth() + 1 !== month
      ) return total;
      return total + (record.paid > 0 ? record.paid : record.receivable);
    }, 0);
    return { month, name: monthNames[month - 1], credited, toReconcile, forecast };
  });
}

function buildDefaulters(records) {
  const people = new Map();
  records.forEach((record) => {
    const dueDate = parseDate(record.dueDate);
    const balance = Math.max(record.receivable - record.paid, 0);
    if (!dueDate || dueDate >= new Date() || balance <= 0) return;
    if (!people.has(record.donor)) {
      people.set(record.donor, { name: record.donor, months: new Set(), amount: 0, installments: 0, overdueRecords: [] });
    }
    const person = people.get(record.donor);
    if (dueDate) person.months.add(dueDate.getMonth() + 1);
    person.amount += balance;
    person.installments += 1;
    person.overdueRecords.push({
      installment: record.installment,
      method: record.method,
      dueDate: record.dueDate,
      dueMonth: dueDate.getMonth() + 1,
      paid: record.paid,
      balance,
    });
  });
  return [...people.values()]
    .map((person) => {
      const planGroups = new Map();
      person.overdueRecords.forEach((record) => {
        const { total } = installmentParts(record.installment);
        const key = `${record.method}::${total}`;
        if (!planGroups.has(key)) planGroups.set(key, { method: record.method, total, records: [] });
        planGroups.get(key).records.push(record);
      });
      return {
        ...person,
        months: [...person.months].sort((a, b) => a - b),
        plans: [...planGroups.values()],
      };
    })
    .sort((a, b) => b.amount - a.amount || a.name.localeCompare(b.name, "pt-BR"));
}

function installmentParts(value) {
  const match = String(value || "").match(/(\d+)\s*\/\s*(\d+)/);
  return match ? { current: Number(match[1]), total: Number(match[2]) } : { current: 0, total: 0 };
}

function overdueRecordDescription(record) {
  const { total } = installmentParts(record.installment);
  if (total === 1) return `pagamento único ${record.installment} (${record.method}) com vencimento em ${record.dueDate}`;
  return `parcela ${record.installment} (${record.method}) com vencimento em ${record.dueDate}`;
}

function sumDefaulters(records) {
  return buildDefaulters(records).reduce((total, person) => total + person.amount, 0);
}

function renderCreditTimeline(records) {
  const timeline = buildCreditTimeline(records);
  return timeline
    .map(
      (item) => `<tr>
        <td><div class="month-cell"><span>${String(item.month).padStart(2, "0")}</span><strong>${item.name}</strong></div></td>
        <td><div class="credit-split-value credit-split-real"><span>Recebido</span><strong>${currency.format(item.credited)}</strong></div></td>
        <td><div class="credit-split-value credit-split-overdue"><span>A conciliar</span><strong>${currency.format(item.toReconcile)}</strong></div></td>
        <td><div class="credit-split-value credit-split-forecast"><span>Previsão futura</span><strong>${currency.format(item.forecast)}</strong></div></td>
      </tr>`,
    )
    .join("");
}

function renderDefaulters(records) {
  const defaulters = buildDefaulters(records);
  if (!defaulters.length) return `<tr><td colspan="4" class="empty-state">Nenhuma parcela sem pagamento.</td></tr>`;
  return defaulters
    .map(
      (person) => `<tr>
        <td><div class="defaulter-person"><span class="avatar">${initials(person.name)}</span><div><strong>${safe(person.name)}</strong><small>${person.installments} ${person.installments === 1 ? "parcela vencida" : "parcelas vencidas"}</small></div></div></td>
        <td><div class="delinquency-type-list">${person.plans
          .map((plan) => {
            const description = plan.total === 1
              ? plan.records.every((record) => record.paid === 0) ? "Pagamento único (1/1) nunca pago" : "Pagamento único (1/1) com saldo em atraso"
              : plan.total > 1
                ? `${plan.records.length} ${plan.records.length === 1 ? "parcela" : "parcelas"} de ${plan.total} em atraso`
                : "Parcela vencida — plano não informado";
            return `<div><strong>${safe(plan.method)}</strong><span>${description}</span></div>`;
          })
          .join("")}</div></td>
        <td><div class="overdue-detail-list">${person.overdueRecords
          .map((record) => `<span><strong>${safe(record.installment)}</strong> · ${safe(monthNames[record.dueMonth - 1])} · ${safe(record.dueDate)}</span>`)
          .join("")}</div></td>
        <td class="money money-pending">${currency.format(person.amount)}</td>
      </tr>`,
    )
    .join("");
}

function groupByParticipant(records) {
  const people = new Map();
  records.forEach((record) => {
    if (!people.has(record.donor)) {
      people.set(record.donor, {
        name: record.donor,
        records: [],
        paid: 0,
        credited: 0,
        awaitingCredit: 0,
        toPay: 0,
        overdue: 0,
        remaining: 0,
        paidInstallments: 0,
        creditedInstallments: 0,
        declaredInstallments: 0,
        totalInstallments: 0,
        nextDate: null,
        nextDateLabel: "—",
      });
    }

    const person = people.get(record.donor);
    person.records.push(record);
    person.totalInstallments += 1;
    const declaredTotal = Number.parseInt(record.installment.split("/")[1], 10) || 0;
    person.declaredInstallments = Math.max(person.declaredInstallments, declaredTotal);
    person.paid += record.paid;
    person.credited += record.credited;
    if (record.paid > 0) person.paidInstallments += 1;
    const dueDate = parseDate(record.dueDate);
    const unpaidBalance = Math.max(record.receivable - record.paid, 0);
    if (dueDate && dueDate < new Date()) person.overdue += unpaidBalance;

    if (record.credited > 0) {
      person.creditedInstallments += 1;
      return;
    }

    if (record.paid > 0) person.awaitingCredit += record.paid;
    person.toPay += Math.max(record.receivable - record.paid, 0);
    person.remaining = person.awaitingCredit + person.toPay;

    const forecastLabel = record.expectedCreditAt || record.dueDate;
    const forecastDate = parseDate(forecastLabel);
    if (forecastDate && (!person.nextDate || forecastDate < person.nextDate)) {
      person.nextDate = forecastDate;
      person.nextDateLabel = forecastLabel;
    }
  });

  return [...people.values()]
    .map((person) => {
      person.totalInstallments = Math.max(person.totalInstallments, person.declaredInstallments);
      if (person.overdue > 0) person.status = "Inadimplente";
      else if (person.awaitingCredit > 0) person.status = "Aguardando crédito";
      else if (person.toPay > 0) person.status = "Pagamento futuro";
      else person.status = "Em dia";
      return person;
    })
    .sort((a, b) => b.overdue - a.overdue || b.awaitingCredit - a.awaitingCredit || a.name.localeCompare(b.name, "pt-BR"));
}

function participantStatusClass(status) {
  return {
    "Em dia": "status-credited",
    "Aguardando crédito": "status-waiting",
    "Pagamento futuro": "status-pending",
    Inadimplente: "status-overdue",
  }[status];
}

function initials(name) {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  return `${parts[0]?.[0] || ""}${parts.length > 1 ? parts.at(-1)[0] : ""}`.toUpperCase();
}

function safe(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function cashFlowComparisonChart(records, ministryRecords) {
  const groups = groupCashFlowByMonth(records, ministryRecords);
  const max = Math.max(...groups.flatMap((item) => [item.income, item.outgoing]), 1);
  const monthFormatter = new Intl.DateTimeFormat("pt-BR", { month: "short" });
  return `
    <div class="monthly-chart" style="--columns:${groups.length}">
      ${groups
        .map((item) => {
          const incomeHeight = Math.max((item.income / max) * 100, item.income ? 2 : 0);
          const outgoingHeight = Math.max((item.outgoing / max) * 100, item.outgoing ? 2 : 0);
          return `<div class="month-group">
            <div class="bar-wrap" style="--height:${incomeHeight}%">
              <div class="bar bar-income" style="height:${incomeHeight}%;opacity:${item.income ? 1 : 0}"></div>
              <div class="bar-tooltip">Receita líquida do curso<br><strong>${currency.format(item.income)}</strong></div>
            </div>
            <div class="bar-wrap" style="--height:${outgoingHeight}%">
              <div class="bar bar-outgoing" style="height:${outgoingHeight}%;opacity:${item.outgoing ? 1 : 0}"></div>
              <div class="bar-tooltip">Despesas das aulas<br><strong>${currency.format(item.outgoing)}</strong></div>
            </div>
          </div>`;
        })
        .join("")}
    </div>
    <div class="chart-labels" style="--columns:${groups.length}">
      ${groups.map((item) => `<span>${monthFormatter.format(item.date).replace(".", "")}</span>`).join("")}
    </div>`;
}

function consolidatedFlowComparisonChart(records, ministryRecords) {
  const groups = groupConsolidatedFlowByMonth(records, ministryRecords);
  const max = Math.max(...groups.flatMap((item) => [item.income, item.outgoing]), 1);
  const monthFormatter = new Intl.DateTimeFormat("pt-BR", { month: "short" });
  return `
    <div class="monthly-chart consolidated-monthly-chart" style="--columns:${groups.length}">
      ${groups.map((item) => {
        const incomeHeight = Math.max((item.income / max) * 100, item.income ? 2 : 0);
        const outgoingHeight = Math.max((item.outgoing / max) * 100, item.outgoing ? 2 : 0);
        const courseShare = item.income ? (item.courseIncome / item.income) * 100 : 0;
        const administrativeShare = item.income ? (item.administrativeIncome / item.income) * 100 : 0;
        return `<div class="month-group">
          <div class="bar-wrap" style="--height:${incomeHeight}%">
            <div class="bar bar-stacked-income" style="height:${incomeHeight}%;opacity:${item.income ? 1 : 0}">
              <span class="bar-income-administrative" style="height:${administrativeShare}%"></span>
              <span class="bar-income-course" style="height:${courseShare}%"></span>
            </div>
            <div class="bar-tooltip">Crédito do curso: <strong>${currency.format(item.courseIncome)}</strong><br>Entrada administrativa: <strong>${currency.format(item.administrativeIncome)}</strong><br>Total que entrou: <strong>${currency.format(item.income)}</strong></div>
          </div>
          <div class="bar-wrap" style="--height:${outgoingHeight}%">
            <div class="bar bar-outgoing" style="height:${outgoingHeight}%;opacity:${item.outgoing ? 1 : 0}"></div>
            <div class="bar-tooltip">Todas as saídas do Ministério<br><strong>${currency.format(item.outgoing)}</strong></div>
          </div>
        </div>`;
      }).join("")}
    </div>
    <div class="chart-labels" style="--columns:${groups.length}">
      ${groups.map((item) => `<span>${monthFormatter.format(item.date).replace(".", "")}</span>`).join("")}
    </div>
    <div class="presentation-monthly-net" style="--columns:${groups.length}">
      ${groups.map((item) => `<span class="${item.net < 0 ? "negative" : "positive"}" title="Resultado registrado de ${safe(monthFormatter.format(item.date))}">${item.net >= 0 ? "+" : ""}${currency.format(item.net)}</span>`).join("")}
    </div>`;
}

function groupSeriesByMonth(records, dateField, valueSelector, predicate = () => true) {
  const groups = new Map();
  records.forEach((record) => {
    if (!predicate(record)) return;
    const date = parseDate(record[dateField]);
    if (!date) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!groups.has(key)) groups.set(key, { key, date, value: 0 });
    groups.get(key).value += valueSelector(record);
  });
  return [...groups.values()].sort((a, b) => a.date - b.date);
}

function forecastCreditSeries(records) {
  return groupSeriesByMonth(
    records,
    "expectedCreditAt",
    (record) => (record.paid > 0 ? record.paid : record.receivable),
    (record) => {
      const expectedDate = parseDate(record.expectedCreditAt);
      return Boolean(record.credited === 0 && expectedDate && expectedDate >= forecastStartsAt);
    },
  );
}

function delinquencySeries(records) {
  return groupSeriesByMonth(
    records,
    "dueDate",
    (record) => Math.max(record.receivable - record.paid, 0),
    (record) => {
      const dueDate = parseDate(record.dueDate);
      return Boolean(dueDate && dueDate < new Date() && record.receivable > record.paid);
    },
  );
}

function singleSeriesChart(groups, tone = "green") {
  if (!groups.length) return `<div class="chart-empty">Sem valores para este gráfico.</div>`;
  const maximum = Math.max(...groups.map((item) => item.value), 1);
  const monthFormatter = new Intl.DateTimeFormat("pt-BR", { month: "short" });
  return `<div class="single-series">
    <div class="single-series-plot" style="--columns:${groups.length}">
      ${groups
        .map((item) => {
          const height = Math.max((item.value / maximum) * 100, item.value ? 3 : 0);
          return `<div class="single-series-column">
            <div class="single-series-value">${currency.format(item.value)}</div>
            <div class="single-series-bar single-series-${tone}" style="height:${height}%"></div>
          </div>`;
        })
        .join("")}
    </div>
    <div class="single-series-labels" style="--columns:${groups.length}">
      ${groups.map((item) => `<span>${monthFormatter.format(item.date).replace(".", "")}</span>`).join("")}
    </div>
  </div>`;
}

function methodsChart(records) {
  const methods = groupByMethod(records);
  const total = methods.reduce((sum, item) => sum + item.value, 0) || 1;
  const colors = ["#006cfc", "#55a9ff", "#24303c", "#8cc8ff"];
  let cursor = 0;
  const stops = methods.map((item, index) => {
    const start = cursor;
    cursor += (item.value / total) * 100;
    return `${colors[index % colors.length]} ${start}% ${cursor}%`;
  });
  const leading = methods[0] ? (methods[0].value / total) * 100 : 0;

  return `
    <div class="method-layout">
      <div class="donut" style="--donut:conic-gradient(${stops.join(",")})">
        <div class="donut-copy"><strong>${leading.toFixed(0)}%</strong><span>principal meio</span></div>
      </div>
      <div class="method-list">
        ${methods
          .map(
            (item, index) => `<div class="method-row">
              <i class="method-swatch" style="background:${colors[index % colors.length]}"></i>
              <span class="method-name">${safe(item.name)}</span>
              <strong class="method-value">${currency.format(item.value)}</strong>
            </div>`,
          )
          .join("")}
      </div>
    </div>`;
}

const compactCurrency = new Intl.NumberFormat("pt-BR", {
  style: "currency",
  currency: "BRL",
  notation: "compact",
  maximumFractionDigits: 1,
});

function monthShort(date) {
  return new Intl.DateTimeFormat("pt-BR", { month: "short" }).format(date).replace(".", "");
}

function operationalMonthlyChart(groups, series, options = {}) {
  if (!groups.length) return `<div class="chart-empty">Sem valores para este gráfico.</div>`;
  const maximum = Math.max(...groups.flatMap((group) => series.map((item) => Math.abs(item.value(group)))), 1);
  return `<div class="audit-chart-scroll"><div class="audit-month-chart" style="--columns:${groups.length}">
    ${groups.map((group) => {
      const result = options.result?.(group);
      const auditKey = options.auditKey?.(group);
      const tag = auditKey ? "button" : "div";
      return `<${tag} class="audit-month-column ${auditKey ? "audit-clickable" : ""}" ${auditKey ? `${auditAttributes(auditKey, `movimentações de ${monthShort(group.date)}`)} type="button"` : ""}>
        <div class="audit-month-bars">
          ${series.map((item) => {
            const value = item.value(group);
            const height = Math.max((Math.abs(value) / maximum) * 100, value ? 3 : 0);
            return `<div class="audit-bar-slot" title="${safe(item.label)}: ${currency.format(value)}">
              <span class="audit-bar-value">${value ? compactCurrency.format(value) : "—"}</span>
              <i class="audit-data-bar ${safe(item.className)}" style="height:${height}%"></i>
            </div>`;
          }).join("")}
        </div>
        <strong class="audit-month-label">${safe(monthShort(group.date))}</strong>
        ${typeof result === "number" ? `<span class="audit-month-result ${result < 0 ? "negative" : "positive"}">${result >= 0 ? "+" : ""}${currency.format(result)}</span>` : ""}
      </${tag}>`;
    }).join("")}
  </div></div>`;
}

function operationalConsolidatedFlowChart(records, ministryRecords) {
  const groups = groupConsolidatedFlowByMonth(records, ministryRecords);
  return operationalMonthlyChart(groups, [
    { label: "Entradas totais", className: "audit-bar-income", value: (item) => item.income },
    { label: "Saídas totais", className: "audit-bar-expense", value: (item) => item.outgoing },
  ], { result: (item) => item.net, auditKey: (item) => `flow-general:${item.key}` });
}

function operationalCourseFlowChart(records, ministryRecords) {
  const groups = groupCashFlowByMonth(records, ministryRecords).map((item) => ({ ...item, net: item.income - item.outgoing }));
  return operationalMonthlyChart(groups, [
    { label: "Creditado pelo curso", className: "audit-bar-course", value: (item) => item.income },
    { label: "Despesas das aulas", className: "audit-bar-classes", value: (item) => item.outgoing },
  ], { result: (item) => item.net, auditKey: (item) => `flow-course:${item.key}` });
}

function buildPendingAuditByMonth(records) {
  const groups = new Map();
  const ensure = (date) => {
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!groups.has(key)) groups.set(key, { key, date, operator: 0, overdue: 0, future: 0 });
    return groups.get(key);
  };
  records.forEach((record) => {
    const expectedDate = parseDate(record.expectedCreditAt);
    const dueDate = parseDate(record.dueDate);
    const unpaid = Math.max(record.receivable - record.paid, 0);
    if (record.paid > 0 && record.credited === 0 && expectedDate) ensure(expectedDate).operator += record.paid;
    if (unpaid > 0 && dueDate) {
      if (dueDate < new Date()) ensure(dueDate).overdue += unpaid;
      else ensure(dueDate).future += unpaid;
    }
  });
  return [...groups.values()].sort((a, b) => a.date - b.date);
}

function pendingAuditChart(records) {
  const groups = buildPendingAuditByMonth(records);
  return operationalMonthlyChart(groups, [
    { label: "Aguardando operadora", className: "audit-bar-operator", value: (item) => item.operator },
    { label: "Vencido com aluno", className: "audit-bar-overdue", value: (item) => item.overdue },
    { label: "Pagamento futuro", className: "audit-bar-future", value: (item) => item.future },
  ], { auditKey: (item) => `flow-pending:${item.key}` });
}

function groupMinistryExpensesByMonth(records) {
  const groups = new Map();
  records.filter((record) => record.type.toLocaleLowerCase("pt-BR") === "saida").forEach((record) => {
    const date = record.date ? new Date(`${record.date}T12:00:00`) : null;
    if (!date || Number.isNaN(date.getTime())) return;
    const key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
    if (!groups.has(key)) groups.set(key, { key, date, classes: 0, general: 0 });
    const group = groups.get(key);
    if (record.financeKind.toLocaleLowerCase("pt-BR") === "aulas") group.classes += record.amount;
    else group.general += record.amount;
  });
  return [...groups.values()].map((group) => ({ ...group, total: group.classes + group.general })).sort((a, b) => a.date - b.date);
}

function ministryExpenseTrendChart(records) {
  const groups = groupMinistryExpensesByMonth(records);
  return operationalMonthlyChart(groups, [
    { label: "Despesas das aulas", className: "audit-bar-classes", value: (item) => item.classes },
    { label: "Outras áreas", className: "audit-bar-general", value: (item) => item.general },
  ], { result: (item) => -item.total, auditKey: (item) => `flow-expense:${item.key}` });
}

function expenseCategoryAuditChart(records) {
  const groups = groupMinistryExpenses(records);
  const maximum = Math.max(...groups.map((group) => group.total), 1);
  return `<div class="expense-ranking-chart">
    ${groups.map((group, index) => {
      const percent = (group.total / maximum) * 100;
      const totalExpenses = groups.reduce((total, item) => total + item.total, 0) || 1;
      const share = (group.total / totalExpenses) * 100;
      const auditKey = `ministry-category:${encodeURIComponent(group.financeKind || group.kind)}:${encodeURIComponent(group.category)}`;
      const kindClass = group.kind.toLocaleLowerCase("pt-BR") === "aulas" ? "expense-kind-classes" : "expense-kind-general";
      return `<button class="expense-ranking-row ${kindClass} audit-clickable" ${auditAttributes(auditKey, `despesas de ${group.category}`)} type="button">
        <span class="expense-ranking-index">${String(index + 1).padStart(2, "0")}</span>
        <span class="expense-ranking-copy"><strong>${safe(group.category)}</strong><small>${safe(group.kind)} · ${number.format(group.records.length)} ${group.records.length === 1 ? "lançamento" : "lançamentos"}</small><i><b style="width:${percent}%"></b></i></span>
        <span class="expense-ranking-total"><strong>${currency.format(group.total)}</strong><small>${share.toFixed(1).replace(".", ",")}% das saídas</small></span>
      </button>`;
    }).join("")}
  </div>`;
}

function projectionScenarioChart() {
  const plan = planningTotals();
  const steps = [
    { label: "Entrada certa", value: yearEndPlan.certainCourseCredit, className: "positive", key: "plan-certain-recovery" },
    { label: "Despesas planejadas", value: -plan.totalExpenses, className: "negative", key: "plan-year-expenses" },
    { label: "Cenário confirmado", value: plan.certainDeficit, className: "result", key: "plan-deficit-base" },
    { label: "Possível Alex", value: plan.alexContribution, className: "uncertain", key: "plan-alex" },
    { label: "Inadimplência incerta", value: yearEndPlan.uncertainDelinquency, className: "uncertain", key: "plan-uncertain-recovery" },
  ];
  const maximum = Math.max(...steps.map((step) => Math.abs(step.value)), 1);
  return `<div class="projection-waterfall">
    ${steps.map((step) => `<button class="projection-step projection-${step.className} audit-clickable" ${auditAttributes(step.key, step.label)} type="button">
      <span>${safe(step.label)}</span>
      <div class="projection-axis"><i style="height:${Math.max((Math.abs(step.value) / maximum) * 100, 5)}%"></i></div>
      <strong>${step.value > 0 ? "+" : ""}${currency.format(step.value)}</strong>
    </button>`).join("")}
  </div>`;
}

function normalizeFinancialText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("pt-BR");
}

function ministryRecordText(record) {
  return normalizeFinancialText([
    record.category,
    record.eventTitle,
    record.name,
    record.specification,
    record.description,
    record.paymentDetails,
  ].filter(Boolean).join(" "));
}

function isInterpretationExpense(record) {
  return record.type.toLocaleLowerCase("pt-BR") === "saida" && ministryRecordText(record).includes("interpret");
}

function isMixedInterpretationExpense(record) {
  const text = ministryRecordText(record);
  return isInterpretationExpense(record) && (text.includes("alimentacao") || text.includes("oferta"));
}

function groupMinistryExpenses(records) {
  const groups = new Map();
  records
    .filter((record) => record.type.toLocaleLowerCase("pt-BR") === "saida")
    .forEach((record) => {
      const key = `${record.financeKind}::${record.category}`;
      if (!groups.has(key)) groups.set(key, { kind: record.financeKind, category: record.category, total: 0, records: [], reasons: new Set() });
      const group = groups.get(key);
      const reason = record.specification || record.description || record.name || record.eventTitle || record.category;
      group.total += record.amount;
      group.records.push(record);
      if (reason) group.reasons.add(reason.replace(/\s+/g, " ").trim());
    });
  return [...groups.values()].sort((a, b) => b.total - a.total || a.category.localeCompare(b.category, "pt-BR"));
}

function planningTotals() {
  const months = yearEndPlan.months.length;
  const fixedGeneral = yearEndPlan.fixedExpenses.reduce((total, expense) => total + expense.monthly * months, 0);
  const alexContribution = yearEndPlan.alexMonthlyContribution * months;
  const totalExpenses = yearEndPlan.courseExpenses + fixedGeneral;
  const certainDeficit = yearEndPlan.certainCourseCredit - totalExpenses;
  const deficitWithAlex = certainDeficit + alexContribution;
  const potentialDeficit = deficitWithAlex + yearEndPlan.uncertainDelinquency;
  return { months, fixedGeneral, alexContribution, totalExpenses, certainDeficit, deficitWithAlex, potentialDeficit };
}

function renderPlanningView(summary, ministrySummary) {
  const directInterpretation = state.ministryRecords.filter((record) => isInterpretationExpense(record) && !isMixedInterpretationExpense(record));
  const mixedInterpretation = state.ministryRecords.filter(isMixedInterpretationExpense);
  const directInterpretationTotal = directInterpretation.reduce((total, record) => total + record.amount, 0);
  const mixedInterpretationTotal = mixedInterpretation.reduce((total, record) => total + record.amount, 0);
  const expenseGroups = groupMinistryExpenses(state.ministryRecords);
  const plan = planningTotals();
  const courseResult = summary.credited - ministrySummary.classes;
  const ministryResult = ministrySummary.entries - ministrySummary.exits;
  const interpretationRows = directInterpretation.map((record) => {
    const reason = record.specification || record.description || record.name || "Interpretação registrada";
    return `<tr><td>${safe(formatIsoDate(record.date))}</td><td><strong>${safe(record.name || record.category)}</strong></td><td>${safe(reason)}</td><td>${safe(record.paymentMethod)}</td><td class="money money-pending">${currency.format(record.amount)}</td></tr>`;
  }).join("");
  const expenseRows = expenseGroups.map((group) => {
    const reasons = [...group.reasons];
    const visibleReasons = reasons.slice(0, 3).join(" · ");
    const remaining = reasons.length - 3;
    return `<tr>
      <td><span class="ministry-kind">${safe(group.kind)}</span></td>
      <td><strong>${safe(group.category)}</strong></td>
      <td class="planning-reason">${safe(visibleReasons)}${remaining > 0 ? ` <b>+ ${remaining} outros motivos</b>` : ""}</td>
      <td>${number.format(group.records.length)} ${group.records.length === 1 ? "lançamento" : "lançamentos"}</td>
      <td class="money money-pending">${currency.format(group.total)}</td>
    </tr>`;
  }).join("");

  return `
    <section class="planning-section app-view" id="planning" data-view="planning">
      <div class="planning-heading">
        <div><p class="eyebrow">Realizado e projetado</p><h2 class="section-title">Entradas, saídas e plano até dezembro</h2><p class="section-copy">Valores realizados vêm das bases. O plano de setembro a dezembro aparece separado e identificado como projeção.</p></div>
        <span class="planning-basis-badge">Auditoria por origem</span>
      </div>

      <div class="planning-scope-grid">
        <article class="planning-scope-card planning-income audit-clickable" ${auditAttributes("course-credited", "entrada geral realizada do curso")}><span>Entrada geral do curso</span><strong>${currency.format(summary.credited)}</strong><p>Valor Creditado líquido na planilha do curso.</p><small>Base: ${safe(state.sourceName)}</small></article>
        <article class="planning-scope-card planning-expense audit-clickable" ${auditAttributes("ministry-classes", "saída geral realizada do curso")}><span>Saída geral do curso</span><strong>${currency.format(ministrySummary.classes)}</strong><p>Despesas registradas no núcleo “aulas”.</p><small>Base: finançassinaisdoreino.json</small></article>
        <article class="planning-scope-card planning-income audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas do Ministério")}><span>Entradas administrativas</span><strong>${currency.format(ministrySummary.entries)}</strong><p>R$ 500 de verba missionária para Alex + R$ 1.000 do Ministério por mês.</p><small>R$ 1.500/mês · janeiro a setembro de 2026</small></article>
        <article class="planning-scope-card planning-expense audit-clickable" ${auditAttributes("ministry-exits", "saída geral registrada do Ministério")}><span>Saída geral do Ministério</span><strong>${currency.format(ministrySummary.exits)}</strong><p>Todas as saídas: aulas e despesas gerais.</p><small>O curso já está incluído neste total</small></article>
      </div>

      <div class="planning-result-strip">
        <div><span>Resultado realizado do curso</span><strong class="${courseResult < 0 ? "planning-negative" : "planning-positive"}">${currency.format(courseResult)}</strong><small>Entrada creditada − despesas das aulas</small></div>
        <div><span>Resultado da base do Ministério</span><strong class="${ministryResult < 0 ? "planning-negative" : "planning-positive"}">${currency.format(ministryResult)}</strong><small>Entradas JSON − todas as saídas JSON</small></div>
        <p><b>Não some as duas saídas:</b> os ${currency.format(ministrySummary.classes)} do curso já fazem parte dos ${currency.format(ministrySummary.exits)} do Ministério.</p>
      </div>

      <article class="card planning-interpretation-card">
        <div class="planning-card-head">
          <div><p class="eyebrow">Campo específico</p><h3 class="panel-title">Interpretação de culto</h3><p class="panel-subtitle">Somente pagamentos identificados diretamente como interpretação ou intérprete.</p></div>
          <button class="planning-focus-total audit-clickable" ${auditAttributes("ministry-interpretation", "pagamentos diretos de interpretação")} type="button"><span>Total identificado</span><strong>${currency.format(directInterpretationTotal)}</strong><small>${number.format(directInterpretation.length)} pagamentos</small></button>
        </div>
        <div class="table-wrap planning-table-wrap"><table class="planning-detail-table"><thead><tr><th>Data</th><th>Intérprete / destino</th><th>Por que foi pago</th><th>Meio</th><th>Valor</th></tr></thead><tbody>${interpretationRows || `<tr><td colspan="5" class="empty-state">Nenhum pagamento de interpretação identificado.</td></tr>`}</tbody></table></div>
        ${mixedInterpretation.length ? `<div class="planning-mixed-note audit-clickable" ${auditAttributes("ministry-interpretation-mixed", "lançamento misto que inclui interpretação")}><span>${icons.clock}</span><div><strong>${currency.format(mixedInterpretationTotal)} em lançamento misto</strong><p>“Alimentação, oferta e intérprete” inclui interpretação, mas o JSON não separa quanto desse valor foi pago ao intérprete. Por isso, ele não foi somado aos ${currency.format(directInterpretationTotal)}.</p></div></div>` : ""}
      </article>

      <article class="card planning-expense-detail-card">
        <div class="planning-card-head"><div><p class="eyebrow">Detalhamento realizado</p><h3 class="panel-title">Quanto foi gasto e por quê</h3><p class="panel-subtitle">Todas as saídas do JSON agrupadas por núcleo e categoria, com os motivos registrados.</p></div><strong class="planning-card-total">${currency.format(ministrySummary.exits)}</strong></div>
        <div class="table-wrap planning-table-wrap"><table class="planning-expense-table"><thead><tr><th>Núcleo</th><th>Categoria</th><th>Motivos registrados</th><th>Quantidade</th><th>Total</th></tr></thead><tbody>${expenseRows}</tbody></table></div>
      </article>

      <div class="planning-projection-label"><p class="eyebrow">Planejamento informado</p><h3 class="section-title">Plano de despesas fixas até dezembro</h3><p>Projeções gerenciais — não são lançamentos realizados no JSON.</p></div>
      <div class="planning-fixed-layout">
        <article class="card planning-fixed-card audit-clickable" ${auditAttributes("plan-fixed-general", "plano de despesas gerais fixas")}>
          <div class="planning-card-head"><div><h3 class="panel-title">Despesas gerais fixas</h3><p class="panel-subtitle">Setembro, outubro, novembro e dezembro</p></div><strong class="planning-card-total">${currency.format(plan.fixedGeneral)}</strong></div>
          <div class="planning-fixed-list">${yearEndPlan.fixedExpenses.map((expense) => `<div><span><strong>${safe(expense.label)}</strong><small>${safe(expense.reason)}</small></span><b>${currency.format(expense.monthly)}/mês</b><em>${currency.format(expense.monthly * plan.months)}</em></div>`).join("")}</div>
          <div class="planning-fixed-equation">${currency.format(850)} × ${plan.months} meses = <strong>${currency.format(plan.fixedGeneral)}</strong></div>
        </article>
        <article class="card planning-year-summary audit-clickable" ${auditAttributes("plan-year-expenses", "resumo das despesas até o final do ano")}>
          <span>Resumo de despesas até o final do ano</span>
          <div><small>Curso</small><strong>${currency.format(yearEndPlan.courseExpenses)}</strong></div>
          <div><small>Despesas gerais</small><strong>${currency.format(plan.fixedGeneral)}</strong></div>
          <footer><small>Total planejado</small><strong>${currency.format(plan.totalExpenses)}</strong></footer>
        </article>
      </div>

      <div class="planning-scenario-grid">
        <article class="card planning-recovery-card audit-clickable" ${auditAttributes("plan-certain-recovery", "crédito certo a receber do curso")}><span>Crédito certo a receber</span><strong>${currency.format(yearEndPlan.certainCourseCredit)}</strong><p>Entrada concreta informada para o curso.</p><small>Considerada no cenário real</small></article>
        <article class="card planning-recovery-card planning-uncertain audit-clickable" ${auditAttributes("plan-uncertain-recovery", "inadimplência que talvez seja recuperada")}><span>Recebimento incerto</span><strong>${currency.format(yearEndPlan.uncertainDelinquency)}</strong><p>Inadimplência do curso que talvez não seja recuperada.</p><small>Não entra no cenário real</small></article>
        <article class="card planning-recovery-card audit-clickable" ${auditAttributes("plan-alex", "possível entrada mensal explicada por Alex")}><span>Possível entrada — Alex</span><strong>${currency.format(plan.alexContribution)}</strong><p>${currency.format(yearEndPlan.alexMonthlyContribution)} por mês durante quatro meses.</p><small>A confirmar</small></article>
      </div>

      <article class="planning-scenario-card">
        <div class="planning-scenario-head"><div><p class="eyebrow">Cenário real</p><h3>Projeção do saldo até dezembro</h3></div><span>Valores informados</span></div>
        <div class="planning-scenario-flow">
          <div><span>Despesas planejadas</span><strong>− ${currency.format(plan.totalExpenses)}</strong><small>Curso + despesas gerais</small></div>
          <i>+</i><div><span>Entrada concreta</span><strong>${currency.format(yearEndPlan.certainCourseCredit)}</strong><small>Crédito certo do curso</small></div>
          <i>=</i><div class="planning-scenario-result audit-clickable" ${auditAttributes("plan-deficit-base", "saldo negativo do cenário real")}><span>Saldo real projetado</span><strong>${currency.format(plan.certainDeficit)}</strong><small>Déficit sem contar Alex e inadimplência</small></div>
        </div>
        <div class="planning-scenario-alternatives">
          <div class="audit-clickable" ${auditAttributes("plan-deficit-alex", "saldo com a possível entrada de Alex")}><span>Se Alex acrescentar ${currency.format(plan.alexContribution)}</span><strong>${currency.format(plan.deficitWithAlex)}</strong><small>Déficit ainda previsto</small></div>
          <div class="audit-clickable" ${auditAttributes("plan-deficit-potential", "saldo potencial recuperando também a inadimplência")}><span>Se também recuperar os ${currency.format(yearEndPlan.uncertainDelinquency)}</span><strong>${currency.format(plan.potentialDeficit)}</strong><small>Cenário potencial, não garantido</small></div>
        </div>
      </article>
    </section>`;
}

function buildLegacyPresentationPages() {
  const summary = summarize(state.records);
  const ministry = summarizeMinistry(state.ministryRecords);
  const timeline = buildCreditTimeline(state.records);
  const defaulters = buildDefaulters(state.records);
  const overdueInstallments = defaulters.reduce((total, person) => total + person.installments, 0);
  const toReconcile = timeline.reduce((total, item) => total + item.toReconcile, 0);
  const futureForecast = timeline.reduce((total, item) => total + item.forecast, 0);
  const ministryBalance = ministry.entries - ministry.exits;
  const ministryCourseExits = state.ministryRecords.filter((record) => record.financeKind === "aulas" && record.type.toLocaleLowerCase("pt-BR") === "saida");
  const directInterpretation = state.ministryRecords.filter((record) => isInterpretationExpense(record) && !isMixedInterpretationExpense(record));
  const mixedInterpretation = state.ministryRecords.filter(isMixedInterpretationExpense);
  const directInterpretationTotal = directInterpretation.reduce((total, record) => total + record.amount, 0);
  const mixedInterpretationTotal = mixedInterpretation.reduce((total, record) => total + record.amount, 0);
  const expenseGroups = groupMinistryExpenses(state.ministryRecords);
  const interpretationItems = directInterpretation.map((record) => {
    const reason = record.specification || record.description || record.eventTitle || "Interpretação registrada";
    return `<li>
      <time>${safe(formatIsoDate(record.date))}</time>
      <span><strong>${safe(record.name || record.category)}</strong><small>${safe(reason)}</small></span>
      <b>${currency.format(record.amount)}</b>
    </li>`;
  }).join("");
  const expenseGroupRows = expenseGroups.map((group) => {
    const reasons = [...group.reasons];
    const visibleReasons = reasons.slice(0, 3).join(" · ");
    const remaining = reasons.length - 3;
    return `<tr>
      <td><span class="presentation-expense-kind">${safe(group.kind)}</span></td>
      <td><strong>${safe(group.category)}</strong></td>
      <td class="presentation-expense-reason">${safe(visibleReasons)}${remaining > 0 ? ` <b>+ ${remaining} outros motivos</b>` : ""}</td>
      <td>${number.format(group.records.length)}</td>
      <td>${currency.format(group.total)}</td>
    </tr>`;
  }).join("");

  const pages = [
    {
      theme: "sources",
      kicker: "Página 1 · Fontes",
      title: "De onde vêm os dados desta auditoria?",
      content: `
        <div class="presentation-intro">
          <p>Esta apresentação usa duas bases independentes. Elas são comparadas quando necessário, mas não são tratadas como um extrato bancário completo.</p>
        </div>
        <div class="presentation-source-grid">
          <article class="audit-clickable" ${auditAttributes("course-all", "lançamentos do curso")}>
            <span class="presentation-source-icon">${icons.people}</span>
            <p class="presentation-source-label">Recebimentos do curso</p>
            <h4>Libras.xlsx</h4>
            <strong>${number.format(state.records.length)} lançamentos · ${number.format(summary.people)} alunos</strong>
            <p>Origina valores pagos, créditos líquidos, tarifas, parcelas, previsões e inadimplência.</p>
          </article>
          <article class="audit-clickable" ${auditAttributes("ministry-all", "lançamentos administrativos")}>
            <span class="presentation-source-icon presentation-source-expense">${icons.money}</span>
            <p class="presentation-source-label">Finanças administrativas</p>
            <h4>finançassinaisdoreino.json</h4>
            <strong>${number.format(ministry.count)} lançamentos</strong>
            <p>Reúne as saídas do JSON e as entradas administrativas recorrentes cadastradas no sistema.</p>
          </article>
        </div>
        <div class="presentation-source-note">${icons.check}<p><strong>Regra da auditoria:</strong> cada número apresentado informa sua coluna, fórmula e base de origem.</p></div>`,
    },
    {
      theme: "income",
      kicker: "Página 2 · Recebimentos",
      title: "Quanto realmente entrou na conta Curso de libras",
      content: `
        <div class="presentation-main-metric presentation-main-income audit-clickable" ${auditAttributes("course-credited", "recebido líquido pelo Ministério")}>
          <span>Recebido líquido pelo Ministério</span>
          <strong>${currency.format(summary.credited)}</strong>
          <p>Fonte: <b>Libras.xlsx</b> · soma da coluna <b>Valor Creditado</b>.</p>
        </div>
        <div class="presentation-metric-grid">
          <article class="audit-clickable" ${auditAttributes("course-paid", "pago pelos alunos")}><span>Pago pelos alunos</span><strong>${currency.format(summary.paid)}</strong><p>Coluna Valor Pago. É informativo e está sem os descontos da operadora de cartão.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-fees", "tarifas e taxas")}><span>Tarifas e taxas</span><strong>${currency.format(summary.fees)}</strong><p>Valor das tarifas e taxas cobradas pela operadora de cartão.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-awaiting", "pago sem crédito registrado")}><span>Pago sem crédito registrado</span><strong>${currency.format(summary.awaitingCredit)}</strong><p>Pagamentos feitos no cartão que serão creditados posteriormente pela operadora na conta do Ministério.</p></article>
        </div>
        <div class="presentation-formula presentation-formula-featured">
          <span>Conferência dos dados atuais</span>
          <div class="presentation-formula-equation">
            <div><small>Pago pelos alunos</small><strong>${currency.format(summary.paid)}</strong></div>
            <b aria-hidden="true">−</b>
            <div><small>Tarifas da operadora</small><strong>${currency.format(summary.fees)}</strong></div>
            <b aria-hidden="true">−</b>
            <div><small>Ainda sem crédito</small><strong>${currency.format(summary.awaitingCredit)}</strong></div>
            <b aria-hidden="true">=</b>
            <div class="presentation-formula-result"><small>Líquido recebido</small><strong>${currency.format(summary.credited)}</strong></div>
          </div>
          <p>O valor líquido é o total pago, descontando as taxas do cartão e as parcelas que a operadora ainda não creditou.</p>
        </div>`,
    },
    {
      theme: "credit",
      kicker: "Página 3 · Competência",
      title: "Quando o dinheiro deveria entrar?",
      content: `
        <div class="presentation-metric-grid presentation-metric-grid-three">
          <article class="audit-clickable" ${auditAttributes("course-awaiting", "pago aguardando crédito")}><span>Pago, aguardando crédito</span><strong>${currency.format(summary.awaitingCredit)}</strong><p>Pagamentos no cartão que a operadora creditará posteriormente.</p></article>
          <article class="presentation-alert-metric audit-clickable" ${auditAttributes("course-reconcile", "crédito vencido")}><span>Crédito vencido</span><strong>${currency.format(toReconcile)}</strong><p>Pago, com previsão encerrada e sem crédito efetivo.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-future", "previsão futura")}><span>Previsão futura</span><strong>${currency.format(futureForecast)}</strong><p>Créditos previstos a partir de setembro.</p></article>
        </div>
        <div class="presentation-table-wrap">
          <table class="compact-table credit-month-table">
            <thead><tr><th>Mês</th><th>Creditado líquido</th><th>Crédito vencido</th><th>Previsão futura</th></tr></thead>
            <tbody>${renderCreditTimeline(state.records)}</tbody>
          </table>
        </div>
        <div class="presentation-data-source">Fonte: Libras.xlsx · Valor Creditado agrupado pela Data de Crédito; pendências agrupadas pela Data de Crédito Previsto.</div>`,
    },
    {
      theme: "delinquency",
      kicker: "Página 4 · Cobrança",
      title: "Quem está inadimplente e como foi calculado?",
      content: `
        <div class="presentation-main-metric presentation-main-danger audit-clickable" ${auditAttributes("course-overdue", "inadimplência vencida")}>
          <span>Inadimplência vencida</span>
          <strong>${currency.format(sumDefaulters(state.records))}</strong>
          <p>${number.format(defaulters.length)} pessoas · ${number.format(overdueInstallments)} parcelas vencidas.</p>
        </div>
        <div class="presentation-defaulter-grid">
          ${defaulters.map((person) => `<article class="audit-clickable" ${auditAttributes(`defaulter:${encodeURIComponent(person.name)}`, `inadimplência de ${person.name}`)}>
            <div><span class="avatar">${initials(person.name)}</span><strong>${safe(person.name)}</strong></div>
            <p>${person.plans.map((plan) => plan.total === 1 ? `Pagamento único 1/1 não pago` : `${plan.records.length} ${plan.records.length === 1 ? "parcela" : "parcelas"} de ${plan.total} em atraso`).join(" · ")}</p>
            <b>${currency.format(person.amount)}</b>
          </article>`).join("")}
        </div>
        <div class="presentation-data-source">Fórmula por parcela: Valor a Receber − Valor Pago. Entram somente vencimentos anteriores à data da auditoria com saldo positivo.</div>`,
    },
    {
      theme: "expenses",
      kicker: "Página 5 · Despesas",
      title: "Entradas, saídas e para onde foi o dinheiro",
      content: `
        <div class="presentation-scope-grid">
          <article class="presentation-scope-card presentation-scope-income audit-clickable" ${auditAttributes("course-credited", "entrada geral realizada do curso")}>
            <span>Entrada geral do curso</span><strong>${currency.format(summary.credited)}</strong>
            <p>Valor líquido realmente creditado na conta.</p><small>Fonte: Valor Creditado · ${safe(state.sourceName)}</small>
          </article>
          <article class="presentation-scope-card presentation-scope-expense audit-clickable" ${auditAttributes("ministry-classes", "saída geral realizada do curso")}>
            <span>Saída geral do curso</span><strong>${currency.format(ministry.classes)}</strong>
            <p>Quanto foi gasto diretamente com o núcleo de aulas.</p><small>${number.format(ministryCourseExits.length)} saídas · finançassinaisdoreino.json</small>
          </article>
          <article class="presentation-scope-card presentation-scope-income audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas do Ministério")}>
            <span>Entradas administrativas</span><strong>${currency.format(ministry.entries)}</strong>
            <p>R$ 500 para Alex + R$ 1.000 do Ministério por mês.</p><small>R$ 1.500/mês · janeiro a setembro de 2026</small>
          </article>
          <article class="presentation-scope-card presentation-scope-expense audit-clickable" ${auditAttributes("ministry-exits", "saída geral registrada do Ministério")}>
            <span>Saída geral do Ministério</span><strong>${currency.format(ministry.exits)}</strong>
            <p>Todas as saídas registradas: aulas e despesas gerais.</p><small>O gasto do curso já está incluído neste total</small>
          </article>
        </div>

        <div class="presentation-scope-note">
          <span>${icons.check}</span>
          <p><strong>Como ler:</strong> não some as duas saídas. Os ${currency.format(ministry.classes)} gastos com o curso já fazem parte dos ${currency.format(ministry.exits)} de saída geral do Ministério.</p>
        </div>

        <div class="presentation-expense-focus">
          <section class="presentation-detail-panel presentation-interpretation-panel">
            <div class="presentation-detail-head">
              <div><span>Campo específico</span><h4>Interpretação de culto</h4><p>Pagamentos identificados diretamente como interpretação ou intérprete.</p></div>
              <button class="presentation-interpretation-total audit-clickable" ${auditAttributes("ministry-interpretation", "pagamentos diretos de interpretação")} type="button"><span>Total direto</span><strong>${currency.format(directInterpretationTotal)}</strong><small>${number.format(directInterpretation.length)} pagamentos</small></button>
            </div>
            <ul class="presentation-interpretation-list">${interpretationItems || `<li class="presentation-empty-detail">Nenhum pagamento direto de interpretação foi identificado.</li>`}</ul>
            ${mixedInterpretation.length ? `<button class="presentation-mixed-expense audit-clickable" ${auditAttributes("ministry-interpretation-mixed", "lançamento misto que inclui interpretação")} type="button"><span>${icons.clock}</span><p><strong>${currency.format(mixedInterpretationTotal)} em lançamento misto</strong> “Alimentação, oferta e intérprete” não separa a parte da interpretação; por isso, não entra no total direto acima.</p></button>` : ""}
          </section>

          <section class="presentation-detail-panel presentation-spending-panel">
            <div class="presentation-detail-head">
              <div><span>Detalhamento realizado</span><h4>Quanto foi gasto e por quê</h4><p>Cada saída agrupada por núcleo e categoria, com os motivos registrados.</p></div>
              <strong class="presentation-spending-total">${currency.format(ministry.exits)}</strong>
            </div>
            <div class="presentation-spending-table-wrap">
              <table class="presentation-spending-table"><thead><tr><th>Núcleo</th><th>Categoria</th><th>Por que foi gasto</th><th>Qtd.</th><th>Total</th></tr></thead><tbody>${expenseGroupRows}</tbody></table>
            </div>
            <p class="presentation-spending-help">Clique em “Saída geral do Ministério” para abrir a auditoria completa de cada lançamento.</p>
          </section>
        </div>`,
    },
    {
      theme: "flow",
      kicker: "Página 6 · Fluxo mensal",
      title: "Entradas administrativas x saídas gerais",
      content: `
        <div class="presentation-chart-card">
          <div class="presentation-chart-head">
            <div><div class="legend"><span><i style="background:#438b77"></i>Entradas administrativas</span><span><i style="background:#b85543"></i>Saídas gerais do Ministério</span></div><p class="presentation-chart-explanation">Comparação mensal de todas as entradas e saídas administrativas registradas.</p></div>
            <div class="presentation-excluded-expense audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas")}>
              <span>Entrada recorrente mensal</span><strong>${currency.format(1500)}</strong><small>R$ 500 para Alex + R$ 1.000 do Ministério</small>
            </div>
          </div>
          ${consolidatedFlowComparisonChart(state.records, state.ministryRecords)}
        </div>
        <div class="presentation-metric-grid presentation-flow-summary">
          <article class="audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas")}><span>Entradas administrativas</span><strong>${currency.format(ministry.entries)}</strong><p>R$ 1.500 por mês, de janeiro a setembro.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-exits", "saídas gerais do Ministério")}><span>Saídas gerais do Ministério</span><strong>${currency.format(ministry.exits)}</strong><p>Inclui ${currency.format(ministry.classes)} das aulas e ${currency.format(ministry.general)} das demais áreas.</p></article>
          <article class="${ministryBalance < 0 ? "presentation-alert-metric" : ""}"><span>Resultado administrativo</span><strong>${currency.format(ministryBalance)}</strong><p>Entradas administrativas menos todas as saídas do Ministério.</p></article>
        </div>`,
    },
    {
      theme: "method",
      kicker: "Página 7 · Metodologia",
      title: "Como ler os números sem misturar conceitos",
      content: `
        <div class="presentation-rules">
          <article><span>01</span><div><strong>Dinheiro recebido</strong><p>Somente Valor Creditado representa entrada real do curso.</p></div></article>
          <article><span>02</span><div><strong>Pagamento do aluno</strong><p>Valor Pago mostra o que saiu da conta do aluno; pode ainda estar aguardando crédito.</p></div></article>
          <article><span>03</span><div><strong>Inadimplência</strong><p>É saldo vencido não pago pelo aluno. Não é a mesma coisa que crédito vencido.</p></div></article>
          <article><span>04</span><div><strong>Finanças do Ministério</strong><p>As despesas vêm do JSON; as entradas administrativas recorrentes cadastradas no sistema aparecem separadamente.</p></div></article>
        </div>
        <div class="presentation-audit-footnote">
          <strong>Ajustes aplicados à base</strong>
          <p>Luciana Caroline Correia da Silva foi excluída dos cálculos. Jaqueline Pereira Fachone também foi excluída após a desistência do curso e a devolução do pagamento. Thaise Almeida e Yasmine Delefrati foram adicionadas conforme solicitado. Os dois lançamentos de Caio são mantidos porque correspondem a ele e à esposa.</p>
        </div>`,
    },
  ];

  return [pages.at(-1), pages[1], pages[3], pages[4], pages[5]].map((page, index) => ({
    ...page,
    kicker: page.kicker.replace(/Página \d+/, `Página ${index + 1}`),
  }));
}

function buildPresentationPages() {
  const summary = summarize(state.records);
  const ministry = summarizeMinistry(state.ministryRecords);
  const timeline = buildCreditTimeline(state.records);
  const defaulters = buildDefaulters(state.records);
  const overdueInstallments = defaulters.reduce((total, person) => total + person.installments, 0);
  const toReconcile = timeline.reduce((total, item) => total + item.toReconcile, 0);
  const futureForecast = timeline.reduce((total, item) => total + item.forecast, 0);
  const registeredInflows = summary.credited + ministry.entries;
  const registeredNet = registeredInflows - ministry.exits;
  const directInterpretation = state.ministryRecords.filter((record) => isInterpretationExpense(record) && !isMixedInterpretationExpense(record));
  const mixedInterpretation = state.ministryRecords.filter(isMixedInterpretationExpense);
  const directInterpretationTotal = directInterpretation.reduce((total, record) => total + record.amount, 0);
  const mixedInterpretationTotal = mixedInterpretation.reduce((total, record) => total + record.amount, 0);
  const expenseRecords = state.ministryRecords.filter((record) => record.type.toLocaleLowerCase("pt-BR") === "saida");
  const missingStructuredReason = expenseRecords.filter((record) => !record.specification && !record.description).length;
  const registeredEntryCount = state.ministryRecords.filter((record) => record.type.toLocaleLowerCase("pt-BR") === "entrada" && record.createdBy === "Cadastro solicitado").length;
  const expenseGroups = groupMinistryExpenses(state.ministryRecords);
  const leadingExpenseGroups = expenseGroups.slice(0, 6);
  const remainingExpenseTotal = expenseGroups.slice(6).reduce((total, group) => total + group.total, 0);
  const expenseRanking = remainingExpenseTotal > 0
    ? [...leadingExpenseGroups, { kind: "geral", category: "Demais categorias", total: remainingExpenseTotal, records: expenseGroups.slice(6).flatMap((group) => group.records) }]
    : leadingExpenseGroups;
  const expenseRankingMax = Math.max(...expenseRanking.map((group) => group.total), 1);
  const interpretationItems = directInterpretation.map((record) => {
    const reason = record.specification || record.description || record.eventTitle || "Interpretação registrada";
    return `<li><time>${safe(formatIsoDate(record.date))}</time><span><strong>${safe(record.name || record.category)}</strong><small>${safe(reason)}</small></span><b>${currency.format(record.amount)}</b></li>`;
  }).join("");
  const pendingMonths = timeline.filter((item) => item.toReconcile > 0 || item.forecast > 0);

  return [
    {
      theme: "method",
      kicker: "Página 1 · Resumo executivo",
      title: "O que os registros financeiros mostram",
      content: `
        <div class="presentation-audit-status-row" aria-label="Estados usados na auditoria">
          <span class="presentation-audit-status status-confirmed">Confirmado no sistema</span>
          <span class="presentation-audit-status status-registered">Cadastrado · comprovante pendente</span>
          <span class="presentation-audit-status status-recorded">Saída registrada</span>
          <span class="presentation-audit-status status-unreconciled">Não conciliado com extrato</span>
        </div>
        <div class="presentation-executive-grid">
          <article class="presentation-executive-card audit-clickable" ${auditAttributes("course-credited", "crédito líquido do curso")}><span>Curso · líquido creditado</span><strong>${currency.format(summary.credited)}</strong><p>Soma da coluna Valor Creditado.</p><small class="presentation-audit-status status-confirmed">Confirmado no sistema do curso</small></article>
          <article class="presentation-executive-card audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas cadastradas")}><span>Entradas administrativas</span><strong>${currency.format(ministry.entries)}</strong><p>R$ 1.500 por mês, de janeiro a setembro.</p><small class="presentation-audit-status status-registered">${number.format(registeredEntryCount)} registros retroativos</small></article>
          <article class="presentation-executive-card presentation-executive-expense audit-clickable" ${auditAttributes("ministry-exits", "todas as saídas do Ministério")}><span>Todas as saídas</span><strong>${currency.format(ministry.exits)}</strong><p>Aulas e demais despesas do Ministério.</p><small class="presentation-audit-status status-recorded">${number.format(expenseRecords.length)} pagamentos registrados</small></article>
          <article class="presentation-executive-card ${registeredNet < 0 ? "presentation-executive-expense" : "presentation-executive-result"} audit-clickable" ${auditAttributes("registered-net", "resultado registrado no período")}><span>Resultado registrado</span><strong>${currency.format(registeredNet)}</strong><p>Entradas registradas menos todas as saídas.</p><small class="presentation-audit-status status-unreconciled">Não representa saldo bancário</small></article>
        </div>
        <div class="presentation-formula presentation-formula-featured presentation-executive-formula">
          <span>Conta do resultado registrado</span>
          <div class="presentation-formula-equation">
            <div><small>Crédito do curso</small><strong>${currency.format(summary.credited)}</strong></div><b aria-hidden="true">+</b>
            <div><small>Entradas administrativas</small><strong>${currency.format(ministry.entries)}</strong></div><b aria-hidden="true">−</b>
            <div><small>Todas as saídas</small><strong>${currency.format(ministry.exits)}</strong></div><b aria-hidden="true">=</b>
            <div class="presentation-formula-result"><small>Resultado registrado</small><strong>${currency.format(registeredNet)}</strong></div>
          </div>
        </div>
        <div class="presentation-audit-caveat">${icons.clock}<p><strong>Limite da auditoria:</strong> este resultado só poderá ser chamado de saldo real depois da conferência com extrato, saldo inicial e comprovantes. As entradas administrativas usam datas mensais cadastradas, não datas bancárias comprovadas.</p></div>`,
    },
    {
      theme: "income",
      kicker: "Página 2 · Receita do curso",
      title: "Como chegamos ao dinheiro líquido do curso",
      content: `
        <div class="presentation-main-metric presentation-main-income audit-clickable" ${auditAttributes("course-credited", "recebido líquido pelo Ministério")}>
          <span>Valor líquido creditado</span><strong>${currency.format(summary.credited)}</strong>
          <p>Dinheiro registrado na coluna <b>Valor Creditado</b> da planilha do curso.</p>
          <small class="presentation-audit-status status-confirmed">Fórmula conferida linha a linha</small>
        </div>
        <div class="presentation-metric-grid">
          <article class="audit-clickable" ${auditAttributes("course-paid", "pago pelos alunos")}><span>Pago pelos alunos</span><strong>${currency.format(summary.paid)}</strong><p>Valor bruto, antes dos descontos da operadora.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-fees", "tarifas da operadora")}><span>Tarifas da operadora</span><strong>${currency.format(summary.fees)}</strong><p>Taxas cobradas no processamento do cartão.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-awaiting", "pago ainda não creditado")}><span>Pago, ainda sem crédito</span><strong>${currency.format(summary.awaitingCredit)}</strong><p>Valor pago no cartão que ainda não entrou.</p></article>
        </div>
        <div class="presentation-formula presentation-formula-featured">
          <span>Reconciliação do curso</span>
          <div class="presentation-formula-equation">
            <div><small>Pago pelos alunos</small><strong>${currency.format(summary.paid)}</strong></div><b aria-hidden="true">−</b>
            <div><small>Tarifas</small><strong>${currency.format(summary.fees)}</strong></div><b aria-hidden="true">−</b>
            <div><small>Ainda sem crédito</small><strong>${currency.format(summary.awaitingCredit)}</strong></div><b aria-hidden="true">=</b>
            <div class="presentation-formula-result"><small>Líquido creditado</small><strong>${currency.format(summary.credited)}</strong></div>
          </div>
          <p>Não foram encontrados créditos sem pagamento, créditos sem data ou diferenças entre Valor Pago − Taxa e Valor Creditado.</p>
        </div>`,
    },
    {
      theme: "delinquency",
      kicker: "Página 3 · Pendências",
      title: "O que ainda não entrou — e por qual motivo",
      content: `
        <div class="presentation-pending-grid">
          <article class="presentation-pending-card presentation-pending-operator audit-clickable" ${auditAttributes("course-awaiting", "pagamentos aguardando a operadora")}>
            <span class="presentation-audit-status status-pending">Pendente da operadora</span><h4>Aluno pagou, Ministério aguarda</h4><strong>${currency.format(summary.awaitingCredit)}</strong>
            <p>Não é inadimplência: o pagamento já ocorreu no cartão.</p>
            <div><span><b>${currency.format(toReconcile)}</b><small>previsão já vencida</small></span><span><b>${currency.format(futureForecast)}</b><small>previsão setembro e outubro</small></span></div>
          </article>
          <article class="presentation-pending-card presentation-pending-student audit-clickable" ${auditAttributes("course-overdue", "inadimplência vencida dos alunos")}>
            <span class="presentation-audit-status status-overdue">Cobrança necessária</span><h4>Aluno ainda não pagou</h4><strong>${currency.format(summary.overdue)}</strong>
            <p>${number.format(defaulters.length)} pessoas · ${number.format(overdueInstallments)} parcelas vencidas.</p>
            <div><span><b>${currency.format(summary.open)}</b><small>total em aberto</small></span><span><b>${currency.format(summary.futureDue)}</b><small>parcelas futuras</small></span></div>
          </article>
        </div>
        <div class="presentation-pending-details">
          <section><div class="presentation-detail-head"><div><span>Repasse do cartão</span><h4>Quando deveria cair</h4></div></div><div class="presentation-pending-months">${pendingMonths.map((item) => `<div><span>${safe(item.name)}</span><strong>${currency.format(item.toReconcile + item.forecast)}</strong><small>${item.toReconcile > 0 ? "Vencido, precisa conciliar" : "Previsão informada"}</small></div>`).join("")}</div></section>
          <section><div class="presentation-detail-head"><div><span>Cobrança por pessoa</span><h4>Quem ainda deve</h4></div></div><div class="presentation-pending-people">${defaulters.map((person) => `<button class="audit-clickable" ${auditAttributes(`defaulter:${encodeURIComponent(person.name)}`, `inadimplência de ${person.name}`)} type="button"><span>${safe(person.name)}</span><strong>${currency.format(person.amount)}</strong><small>${number.format(person.installments)} ${person.installments === 1 ? "parcela" : "parcelas"}</small></button>`).join("")}</div></section>
        </div>
        <div class="presentation-data-source">Não some os dois valores como dinheiro a receber da mesma origem: um depende do repasse da operadora; o outro depende do pagamento do aluno.</div>`,
    },
    {
      theme: "expenses",
      kicker: "Página 4 · Despesas",
      title: "Quanto foi gasto e para onde o dinheiro foi",
      content: `
        <div class="presentation-metric-grid">
          <article class="audit-clickable" ${auditAttributes("ministry-exits", "total de saídas")}><span>Todas as saídas</span><strong>${currency.format(ministry.exits)}</strong><p>${number.format(expenseRecords.length)} pagamentos registrados.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-classes", "despesas das aulas")}><span>Aulas</span><strong>${currency.format(ministry.classes)}</strong><p>Professores e despesas diretamente classificadas no núcleo aulas.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-general", "demais despesas do Ministério")}><span>Demais despesas</span><strong>${currency.format(ministry.general)}</strong><p>Alimentação, transporte, interpretação e outras áreas.</p></article>
        </div>
        <div class="presentation-expense-audit-grid">
          <section class="presentation-detail-panel">
            <div class="presentation-detail-head"><div><span>Concentração dos gastos</span><h4>Maiores categorias</h4><p>Valores agrupados pela classificação registrada.</p></div></div>
            <div class="presentation-expense-ranking">${expenseRanking.map((group) => `<div><span><strong>${safe(group.category)}</strong><small>${safe(group.kind)} · ${number.format(group.records.length)} ${group.records.length === 1 ? "lançamento" : "lançamentos"}</small></span><i><b style="width:${(group.total / expenseRankingMax) * 100}%"></b></i><em>${currency.format(group.total)}</em></div>`).join("")}</div>
          </section>
          <section class="presentation-detail-panel presentation-interpretation-panel">
            <div class="presentation-detail-head"><div><span>Campo específico</span><h4>Interpretação de culto</h4><p>Pagamentos identificados diretamente como interpretação.</p></div><button class="presentation-interpretation-total audit-clickable" ${auditAttributes("ministry-interpretation", "pagamentos diretos de interpretação")} type="button"><span>Total direto</span><strong>${currency.format(directInterpretationTotal)}</strong><small>${number.format(directInterpretation.length)} pagamentos</small></button></div>
            <ul class="presentation-interpretation-list">${interpretationItems}</ul>
            ${mixedInterpretation.length ? `<button class="presentation-mixed-expense audit-clickable" ${auditAttributes("ministry-interpretation-mixed", "lançamento misto com interpretação")} type="button"><span>${icons.clock}</span><p><strong>${currency.format(mixedInterpretationTotal)} em lançamento misto.</strong> Alimentação, oferta e intérprete não foram separados; esse valor não entra no total direto.</p></button>` : ""}
          </section>
        </div>
        <div class="presentation-audit-caveat presentation-audit-warning">${icons.clock}<p><strong>Qualidade da documentação:</strong> ${number.format(missingStructuredReason)} das ${number.format(expenseRecords.length)} saídas não possuem especificação ou descrição estruturada. O nome ajuda a identificar o gasto, mas ainda faltam comprovante, conta de origem e situação da conciliação.</p></div>`,
    },
    {
      theme: "flow",
      kicker: "Página 5 · Fluxo consolidado",
      title: "Tudo que entrou x tudo que saiu por mês",
      content: `
        <div class="presentation-chart-card presentation-consolidated-chart-card">
          <div class="presentation-chart-head">
            <div><div class="legend"><span><i style="background:#55a9ff"></i>Entradas administrativas</span><span><i style="background:#438b77"></i>Crédito do curso</span><span><i style="background:#b85543"></i>Todas as saídas</span></div><p class="presentation-chart-explanation">A barra de entrada é composta pelas duas origens. A linha abaixo mostra o resultado registrado de cada mês.</p></div>
            <div class="presentation-excluded-expense audit-clickable" ${auditAttributes("combined-entries", "todas as entradas registradas")}><span>Total que entrou</span><strong>${currency.format(registeredInflows)}</strong><small>Curso + entradas administrativas</small></div>
          </div>
          ${consolidatedFlowComparisonChart(state.records, state.ministryRecords)}
        </div>
        <div class="presentation-metric-grid presentation-flow-summary">
          <article class="audit-clickable" ${auditAttributes("combined-entries", "todas as entradas registradas")}><span>Entradas registradas</span><strong>${currency.format(registeredInflows)}</strong><p>${currency.format(summary.credited)} do curso + ${currency.format(ministry.entries)} administrativas.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-exits", "todas as saídas do Ministério")}><span>Todas as saídas</span><strong>${currency.format(ministry.exits)}</strong><p>${currency.format(ministry.classes)} das aulas + ${currency.format(ministry.general)} das demais áreas.</p></article>
          <article class="${registeredNet < 0 ? "presentation-alert-metric" : ""} audit-clickable" ${auditAttributes("registered-net", "resultado registrado no período")}><span>Resultado registrado</span><strong>${currency.format(registeredNet)}</strong><p>Fluxo do período; não é saldo bancário conciliado.</p></article>
        </div>
        <div class="presentation-audit-caveat">${icons.check}<p><strong>Leitura correta:</strong> o resultado de ${currency.format(registeredNet)} só é válido para os lançamentos presentes nas bases. Saldo inicial, transferências internas e movimentos ausentes podem alterar o saldo real da conta.</p></div>`,
    },
  ];
}

function renderPresentation() {
  const region = document.querySelector("#presentation-region");
  if (!region) return;
  const pages = buildPresentationPages();
  state.presentationPage = Math.min(Math.max(state.presentationPage, 0), pages.length - 1);
  const page = pages[state.presentationPage];

  region.innerHTML = `
    <div class="presentation-experience presentation-theme-${page.theme}">
      <div class="presentation-launch-row">
        <button class="button button-primary presentation-fullscreen-button" data-presentation-action="fullscreen" type="button">${icons.spark}<span class="fullscreen-enter-label">Apresentar</span><span class="fullscreen-exit-label">Encerrar apresentação</span></button>
      </div>
      <div class="presentation-stage">
        <span class="presentation-orb presentation-orb-one"></span>
        <span class="presentation-orb presentation-orb-two"></span>
        <img class="presentation-brand-crown" src="${crownLogo}" alt="" aria-hidden="true" />
        <button class="presentation-stage-arrow presentation-stage-arrow-prev" data-presentation-action="prev" type="button" aria-label="Página anterior" ${state.presentationPage === 0 ? "disabled" : ""}>‹</button>
        <article class="presentation-slide">
        <div class="presentation-slide-head">
          <div><p class="presentation-kicker">${page.kicker}</p><h3>${page.title}</h3></div>
          <span class="presentation-page-number">${String(state.presentationPage + 1).padStart(2, "0")}</span>
        </div>
        <div class="presentation-slide-content">${page.content}</div>
        </article>
        <button class="presentation-stage-arrow presentation-stage-arrow-next" data-presentation-action="next" type="button" aria-label="Próxima página" ${state.presentationPage === pages.length - 1 ? "disabled" : ""}>›</button>
      </div>
      <footer class="presentation-controls">
        <button class="button presentation-control" data-presentation-action="prev" type="button" ${state.presentationPage === 0 ? "disabled" : ""}>‹ Anterior</button>
        <div class="presentation-progress" aria-label="Páginas da apresentação">
          ${pages.map((_, index) => `<button class="presentation-dot ${index === state.presentationPage ? "active" : ""}" data-presentation-page="${index}" type="button" aria-label="Ir para a página ${index + 1}">${index + 1}</button>`).join("")}
        </div>
        <span class="presentation-position">Página ${state.presentationPage + 1} de ${pages.length}</span>
        <button class="button button-primary presentation-control" data-presentation-action="next" type="button" ${state.presentationPage === pages.length - 1 ? "disabled" : ""}>Próxima ›</button>
      </footer>
    </div>`;

  region.querySelectorAll("[data-presentation-page]").forEach((button) => {
    button.addEventListener("click", () => {
      state.presentationPage = Number(button.dataset.presentationPage);
      renderPresentation();
    });
  });
  region.querySelectorAll("[data-presentation-action]").forEach((button) => {
    button.addEventListener("click", () => {
      const action = button.dataset.presentationAction;
      if (action === "fullscreen") {
        const presentation = document.querySelector("#presentation");
        if (document.fullscreenElement) document.exitFullscreen?.();
        else presentation?.requestFullscreen?.();
        return;
      }
      const nextPage = state.presentationPage + (action === "next" ? 1 : -1);
      state.presentationPage = Math.min(Math.max(nextPage, 0), pages.length - 1);
      renderPresentation();
    });
  });
}

function renderMinistryAuditContent(ministrySummary) {
  const directInterpretation = state.ministryRecords.filter((record) => isInterpretationExpense(record) && !isMixedInterpretationExpense(record));
  const mixedInterpretation = state.ministryRecords.filter(isMixedInterpretationExpense);
  const directTotal = directInterpretation.reduce((total, record) => total + record.amount, 0);
  const mixedTotal = mixedInterpretation.reduce((total, record) => total + record.amount, 0);
  const undocumented = state.ministryRecords.filter((record) =>
    record.type.toLocaleLowerCase("pt-BR") === "saida" && !record.specification && !record.description,
  );
  const interpretationRows = directInterpretation.map((record) => `<tr>
    <td>${safe(formatIsoDate(record.date))}</td>
    <td><strong>${safe(record.name || record.category)}</strong></td>
    <td>${safe(record.specification || record.description || record.eventTitle || "Interpretação registrada")}</td>
    <td>${safe(record.paymentMethod)}</td>
    <td class="money money-pending">${currency.format(record.amount)}</td>
  </tr>`).join("");

  return `<div class="ministry-audit-content">
    <article class="card operational-panel">
      <div class="operational-panel-head">
        <div><p class="eyebrow">Destino das saídas</p><h3 class="panel-title">Onde o Ministério gastou</h3><p class="panel-subtitle">Categorias ordenadas pelo total. Clique em uma barra para conferir os lançamentos que formam o valor.</p></div>
        <div class="audit-inline-definition"><strong>${currency.format(ministrySummary.exits)}</strong><span>${number.format(state.ministryRecords.filter((record) => record.type.toLocaleLowerCase("pt-BR") === "saida").length)} saídas registradas</span></div>
      </div>
      <div class="operational-legend"><span><i class="legend-classes"></i>Núcleo de aulas</span><span><i class="legend-general"></i>Outras áreas</span></div>
      ${expenseCategoryAuditChart(state.ministryRecords)}
      <p class="graph-reading"><strong>Como ler:</strong> o comprimento compara as categorias; o percentual mostra quanto cada uma representa nas saídas. O valor de aulas já faz parte do total geral e não deve ser somado novamente.</p>
    </article>

    <article class="card operational-panel interpretation-audit-panel">
      <div class="operational-panel-head">
        <div><p class="eyebrow">Controle específico</p><h3 class="panel-title">Interpretação de culto</h3><p class="panel-subtitle">Somente pagamentos identificados diretamente como interpretação ou intérprete.</p></div>
        <button class="audit-inline-definition audit-clickable" ${auditAttributes("ministry-interpretation", "pagamentos diretos de interpretação")} type="button"><strong>${currency.format(directTotal)}</strong><span>${number.format(directInterpretation.length)} pagamentos identificados</span></button>
      </div>
      <div class="table-wrap planning-table-wrap"><table class="planning-detail-table"><thead><tr><th>Data</th><th>Intérprete / destino</th><th>Motivo registrado</th><th>Meio</th><th>Valor</th></tr></thead><tbody>${interpretationRows || `<tr><td colspan="5" class="empty-state">Nenhum pagamento direto identificado.</td></tr>`}</tbody></table></div>
      ${mixedInterpretation.length ? `<button class="ministry-mixed-warning audit-clickable" ${auditAttributes("ministry-interpretation-mixed", "lançamento misto com interpretação")} type="button"><span>${icons.clock}</span><div><strong>${currency.format(mixedTotal)} em lançamento misto</strong><p>O texto reúne alimentação, oferta e intérprete; como não há divisão, esse total não foi tratado como pagamento exclusivo de interpretação.</p></div></button>` : ""}
    </article>

    <div class="audit-documentation-note ${undocumented.length ? "has-warning" : ""}">
      <span>${icons.rows}</span><div><strong>Qualidade da documentação</strong><p>${number.format(undocumented.length)} de ${number.format(state.ministryRecords.filter((record) => record.type.toLocaleLowerCase("pt-BR") === "saida").length)} saídas não têm especificação ou descrição estruturada. O nome do destino continua disponível, mas comprovante, conta bancária e conciliação não fazem parte desta base.</p></div>
    </div>
  </div>`;
}

function renderOperationalCashflow(summary, ministrySummary) {
  const plan = planningTotals();
  const registeredEntries = summary.credited + ministrySummary.entries;
  const registeredNet = registeredEntries - ministrySummary.exits;
  return `<section class="flow-section app-view operational-flow" id="cashflow" data-view="cashflow">
    <div class="flow-heading operational-page-heading">
      <div><p class="eyebrow">Análise mensal auditável</p><h2 class="section-title">Fluxo e projeções</h2><p class="section-copy">Cada gráfico responde uma pergunta diferente. Realizado, registrado e projetado permanecem visualmente separados.</p></div>
      <span class="participant-total-badge">5 leituras</span>
    </div>

    <div class="operational-flow-stack">
      <article class="card operational-panel operational-chart-primary">
        <div class="operational-panel-head"><div><span class="chart-number">01</span><h3 class="panel-title">Quanto entrou e quanto saiu em cada mês?</h3><p class="panel-subtitle">Entradas = Valor Creditado do curso + entradas administrativas. Saídas = todos os lançamentos de saída do Ministério.</p></div><div class="audit-inline-definition"><strong>${currency.format(registeredNet)}</strong><span>resultado registrado no período</span></div></div>
        <div class="operational-legend"><span><i class="legend-income"></i>Entradas totais</span><span><i class="legend-expense"></i>Saídas totais</span><span><b>±</b>Resultado do mês</span></div>
        ${operationalConsolidatedFlowChart(state.records, state.ministryRecords)}
        <p class="graph-reading"><strong>Por que importa:</strong> identifica os meses em que as saídas foram maiores que as entradas. Este resultado não inclui ${currency.format(summary.awaitingCredit)} aguardando a operadora nem ${currency.format(summary.overdue)} de inadimplência. Clique em um mês para conferir a composição.</p>
      </article>

      <article class="card operational-panel">
        <div class="operational-panel-head"><div><span class="chart-number">02</span><h3 class="panel-title">O curso pagou as próprias aulas?</h3><p class="panel-subtitle">Compara exclusivamente o crédito líquido do curso com as despesas classificadas no núcleo “aulas”.</p></div><button class="audit-inline-definition audit-clickable" ${auditAttributes("comparison-difference", "resultado direto do curso")} type="button"><strong>${currency.format(summary.credited - ministrySummary.classes)}</strong><span>resultado direto do curso</span></button></div>
        <div class="operational-legend"><span><i class="legend-course"></i>Creditado do curso</span><span><i class="legend-classes"></i>Despesas das aulas</span><span class="legend-aside">Outras áreas excluídas: ${currency.format(ministrySummary.general)}</span></div>
        ${operationalCourseFlowChart(state.records, state.ministryRecords)}
        <p class="graph-reading"><strong>Critério:</strong> as demais despesas do Ministério ficam fora desta conta. Clique em um mês para abrir receita e custos diretos daquele período.</p>
      </article>

      <article class="card operational-panel">
        <div class="operational-panel-head"><div><span class="chart-number">03</span><h3 class="panel-title">Por que o dinheiro ainda não entrou?</h3><p class="panel-subtitle">Separa atraso da operadora, inadimplência do aluno e parcelas que ainda não venceram.</p></div><div class="pending-summary-mini"><span><b>Operadora</b>${currency.format(summary.awaitingCredit)}</span><span><b>Alunos</b>${currency.format(summary.overdue)}</span><span><b>Futuro</b>${currency.format(summary.futureDue)}</span></div></div>
        <div class="operational-legend"><span><i class="legend-operator"></i>Aguardando operadora</span><span><i class="legend-overdue"></i>Vencido com aluno</span><span><i class="legend-future"></i>Pagamento futuro</span></div>
        ${pendingAuditChart(state.records)}
        <p class="graph-reading"><strong>Leitura correta:</strong> pagamento no cartão aguardando repasse não é inadimplência. Clique no mês para saber quais alunos ou repasses formam o valor.</p>
      </article>

      <article class="card operational-panel">
        <div class="operational-panel-head"><div><span class="chart-number">04</span><h3 class="panel-title">Como as despesas evoluíram?</h3><p class="panel-subtitle">Mostra mensalmente quanto pertence às aulas e quanto pertence às demais atividades do Ministério.</p></div><div class="audit-inline-definition"><strong>${currency.format(ministrySummary.exits)}</strong><span>saídas acumuladas</span></div></div>
        <div class="operational-legend"><span><i class="legend-classes"></i>Despesas das aulas</span><span><i class="legend-general"></i>Outras áreas</span></div>
        ${ministryExpenseTrendChart(state.ministryRecords)}
        <p class="graph-reading"><strong>Por que importa:</strong> evidencia quando houve concentração de gastos. Clique no mês para conferir destinos e categorias.</p>
      </article>

      <article class="card operational-panel projection-panel">
        <div class="operational-panel-head"><div><span class="chart-number">05</span><h3 class="panel-title">Qual é o cenário planejado até dezembro?</h3><p class="panel-subtitle">Premissas informadas pelo Ministério. Não são lançamentos realizados nem créditos bancários confirmados.</p></div><span class="projection-badge">Projeção gerencial</span></div>
        ${projectionScenarioChart()}
        <div class="projection-details-grid">
          <button class="projection-detail audit-clickable" ${auditAttributes("plan-fixed-general", "despesas gerais fixas planejadas")} type="button"><span>Despesas gerais fixas</span><strong>${currency.format(plan.fixedGeneral)}</strong><small>R$ 850,00 × 4 meses</small></button>
          <button class="projection-detail audit-clickable" ${auditAttributes("plan-year-expenses", "despesas totais planejadas")} type="button"><span>Despesas totais planejadas</span><strong>${currency.format(plan.totalExpenses)}</strong><small>Curso + despesas gerais</small></button>
          <button class="projection-detail projection-detail-result audit-clickable" ${auditAttributes("plan-deficit-base", "cenário confirmado até dezembro")} type="button"><span>Cenário confirmado</span><strong>${currency.format(plan.certainDeficit)}</strong><small>Sem valores incertos</small></button>
        </div>
        <p class="graph-reading graph-reading-warning"><strong>Atenção de auditoria:</strong> a possível entrada de Alex e a recuperação da inadimplência ficam fora do cenário confirmado. A verba de Alex precisa ser conciliada para não duplicar a entrada recorrente já cadastrada.</p>
      </article>
    </div>
  </section>`;
}

function renderApp() {
  const summary = summarize(state.records);
  const ministrySummary = summarizeMinistry(state.ministryRecords);
  const forecastTotal = buildCreditTimeline(state.records)
    .filter((item) => item.month === 9 || item.month === 10)
    .reduce((total, item) => total + item.forecast, 0);
  const registeredEntries = summary.credited + ministrySummary.entries;
  const registeredNet = registeredEntries - ministrySummary.exits;
  const app = document.querySelector("#app");

  app.innerHTML = `
    <div class="app-shell">
      <aside class="sidebar" aria-label="Navegação principal">
        <div class="brand brand-with-logo">
          <div class="brand-logo-surface"><img class="brand-logo" src="${brandLogo}" alt="Ministério Sinais do Reino" /></div>
          <span class="brand-product">Financeiro</span>
        </div>
        <button class="mobile-menu-toggle" type="button" aria-expanded="false" aria-controls="side-navigation" aria-label="Abrir menu de navegação">
          ${icons.menu}<span class="mobile-menu-close-icon">${icons.close}</span><b>Menu</b>
        </button>
        <p class="side-label">Painel</p>
        <nav class="side-nav" id="side-navigation" aria-label="Seções do painel">
          <button class="nav-item ${state.activeView === "presentation" ? "active" : ""}" data-target="presentation" aria-current="${state.activeView === "presentation" ? "page" : "false"}">${icons.spark}<span>Apresentação</span></button>
          <button class="nav-item ${state.activeView === "overview" ? "active" : ""}" data-target="overview" aria-current="${state.activeView === "overview" ? "page" : "false"}">${icons.overview}<span>Visão geral</span></button>
          <button class="nav-item ${state.activeView === "participants" ? "active" : ""}" data-target="participants" aria-current="${state.activeView === "participants" ? "page" : "false"}">${icons.people}<span>Curso e alunos</span></button>
          <button class="nav-item ${state.activeView === "ministry" ? "active" : ""}" data-target="ministry" aria-current="${state.activeView === "ministry" ? "page" : "false"}">${icons.money}<span>Ministério e despesas</span></button>
          <button class="nav-item ${state.activeView === "cashflow" ? "active" : ""}" data-target="cashflow" aria-current="${state.activeView === "cashflow" ? "page" : "false"}">${icons.chart}<span>Fluxo e projeções</span></button>
        </nav>
        <button class="mobile-menu-backdrop" type="button" aria-label="Fechar menu"></button>
        <div class="side-source">
          <div class="source-line"><i class="source-dot"></i>Fonte conectada</div>
          <p class="source-file">${safe(state.sourceName)}</p>
          <p class="source-file source-file-secondary">finançassinaisdoreino.json + entradas cadastradas</p>
        </div>
      </aside>

      <main class="main-content">
        <header class="topbar">
          <div>
            <p class="eyebrow">Ministério Sinais do Reino</p>
            <h1 id="page-title">${viewTitles[state.activeView]}</h1>
          </div>
          <div class="top-actions">
            <input id="file-input" type="file" accept=".xlsx,.xls" hidden />
            <button class="button" id="import-button" type="button" title="Importar outra planilha">${icons.upload}<span>Importar</span></button>
            <button class="button button-primary" id="export-button" type="button" title="Exportar lançamentos filtrados">${icons.download}<span>Exportar CSV</span></button>
          </div>
        </header>

        <section class="presentation-section app-view" id="presentation" data-view="presentation" tabindex="-1" aria-label="Apresentação financeira">
          <div id="presentation-region"></div>
        </section>

        <section class="overview-executive app-view" data-view="overview" aria-label="Resumo executivo da auditoria">
          <div class="overview-executive-heading"><div><p class="eyebrow">Leitura executiva</p><h2 class="section-title">O essencial para entender o dinheiro</h2><p class="section-copy">Os cards resumem as duas bases. Clique em qualquer valor para conferir a conta e os lançamentos de origem.</p></div><span class="audit-basis-pill">Dados registrados · não conciliados com extrato</span></div>
          <div class="kpi-grid executive-kpi-grid">
            <article class="card kpi-card kpi-card-primary audit-clickable" ${auditAttributes("course-credited", "recebido líquido do curso")}><div class="kpi-head"><span class="kpi-label">Recebido líquido do curso</span><span class="kpi-icon">${icons.bank}</span></div><div class="kpi-value">${currency.format(summary.credited)}</div><p class="kpi-note">Somente Valor Creditado: dinheiro líquido registrado como recebido do curso.</p></article>
            <article class="card kpi-card audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas")}><div class="kpi-head"><span class="kpi-label">Entradas administrativas</span><span class="kpi-icon">${icons.money}</span></div><div class="kpi-value">${currency.format(ministrySummary.entries)}</div><p class="kpi-note">R$ 1.500,00 mensais cadastrados de janeiro a setembro; comprovação bancária pendente.</p></article>
            <article class="card kpi-card kpi-card-info audit-clickable" ${auditAttributes("combined-entries", "todas as entradas registradas")}><div class="kpi-head"><span class="kpi-label">Todas as entradas registradas</span><span class="kpi-icon">${icons.check}</span></div><div class="kpi-value">${currency.format(registeredEntries)}</div><p class="kpi-note">Crédito líquido do curso + entradas administrativas cadastradas.</p></article>
            <article class="card kpi-card kpi-card-danger audit-clickable" ${auditAttributes("ministry-exits", "todas as saídas registradas")}><div class="kpi-head"><span class="kpi-label">Todas as saídas registradas</span><span class="kpi-icon">${icons.money}</span></div><div class="kpi-value">${currency.format(ministrySummary.exits)}</div><p class="kpi-note">Inclui despesas das aulas e despesas das demais áreas do Ministério.</p></article>
            <article class="card kpi-card ${registeredNet < 0 ? "kpi-card-danger" : "kpi-card-result"} audit-clickable" ${auditAttributes("registered-net", "resultado registrado")}><div class="kpi-head"><span class="kpi-label">Resultado registrado</span><span class="kpi-icon">${icons.chart}</span></div><div class="kpi-value">${currency.format(registeredNet)}</div><p class="kpi-note">Entradas registradas − saídas registradas. Não inclui ${currency.format(summary.awaitingCredit)} aguardando a operadora nem ${currency.format(summary.overdue)} de inadimplência.</p></article>
            <article class="card kpi-card kpi-attention-card"><div class="kpi-head"><span class="kpi-label">Pendências e previsão</span><span class="kpi-icon">${icons.clock}</span></div><div class="attention-lines"><button class="audit-clickable" ${auditAttributes("course-awaiting", "aguardando operadora")} type="button"><span>Aguardando operadora</span><strong>${currency.format(summary.awaitingCredit)}</strong></button><button class="audit-clickable" ${auditAttributes("course-overdue", "inadimplência dos alunos")} type="button"><span>Inadimplência vencida</span><strong>${currency.format(summary.overdue)}</strong></button><button class="audit-clickable" ${auditAttributes("course-future", "previsão de setembro e outubro")} type="button"><span>Repasse previsto set. + out.</span><strong>${currency.format(forecastTotal)}</strong></button></div><small class="attention-overlap-note">Não some os três valores: a previsão por mês pode fazer parte do total aguardando a operadora.</small></article>
          </div>
          <div class="overview-reconciliation-note"><span>${icons.bank}</span><p><strong>Limite desta auditoria:</strong> os totais refletem os registros disponíveis. Saldo inicial, extrato bancário, comprovantes e possíveis transferências internas ainda precisam ser conciliados para afirmar o saldo da conta.</p></div>
        </section>

        <section class="participants-section app-view" id="participants" data-view="participants">
          <div class="participant-heading">
            <div>
              <p class="eyebrow">Auditoria individual</p>
              <h2 class="section-title">Curso e alunos</h2>
              <p class="section-copy">Pessoas, inadimplência e parcelas ficam reunidas em uma única sequência, sem repetir os cards da Visão geral.</p>
            </div>
            <div class="participant-total-badge">${number.format(summary.people)} alunos</div>
          </div>

          <article class="card people-card">
            <div class="people-toolbar">
              <div>
                <h3 class="panel-title">Situação financeira individual</h3>
                <p class="panel-subtitle">Clique no nome para abrir a auditoria completa das parcelas</p>
              </div>
              <div class="people-tools">
                <label class="search-box" aria-label="Buscar pessoa">${icons.search}<input id="participant-search" type="search" placeholder="Buscar pessoa…" /></label>
                <select class="select person-select" id="participant-name" aria-label="Selecionar uma pessoa">
                  <option>Todas as pessoas</option>
                  ${[...new Set(state.records.map((item) => item.donor))]
                    .sort((a, b) => a.localeCompare(b, "pt-BR"))
                    .map((name) => `<option value="${safe(name)}">${safe(name)}</option>`)
                    .join("")}
                </select>
                <select class="select" id="participant-status" aria-label="Filtrar situação da pessoa">
                  <option>Todas as situações</option>
                  <option>Em dia</option>
                  <option>Aguardando crédito</option>
                  <option>Pagamento futuro</option>
                  <option>Inadimplente</option>
                </select>
              </div>
            </div>
            <div id="participant-region"></div>
            <div class="calculation-note">
              <span>${icons.bank}</span>
              <p><strong>Leitura:</strong> “Creditado líquido” é o valor real recebido. “Pago pelo aluno” é informativo. “Ainda vai cair” reúne pagamentos com crédito zerado. “Devendo” considera parcelas vencidas não pagas.</p>
            </div>
          </article>

          <article class="card course-overdue-section">
            <div class="overview-data-head"><div><p class="eyebrow eyebrow-alert">Cobrança</p><h3 class="panel-title">Quem está inadimplente</h3><p class="panel-subtitle">A tabela diferencia pagamento único nunca pago de parcelas específicas em atraso.</p></div><button class="defaulter-count audit-clickable" ${auditAttributes("course-overdue", "soma dos inadimplentes")} type="button">${buildDefaulters(state.records).length} pessoas · ${currency.format(sumDefaulters(state.records))}</button></div>
            <div class="compact-table-wrap"><table class="compact-table defaulter-table"><thead><tr><th>Participante</th><th>Tipo da inadimplência</th><th>Parcelas e vencimentos</th><th>Em aberto</th></tr></thead><tbody>${renderDefaulters(state.records)}</tbody></table></div>
          </article>

          <article class="card table-card course-ledger-card">
            <div class="table-top"><div><p class="eyebrow">Base da auditoria</p><h3 class="panel-title">Todas as parcelas do curso</h3><p class="panel-subtitle">Uma linha por parcela. Os filtros atualizam todos os resultados abaixo sem criar páginas.</p></div><button class="button" id="clear-filters" type="button">Limpar filtros</button></div>
            <div class="quick-filter-bar" aria-label="Filtros rápidos"><button class="quick-filter active" data-quick-status="Todos" type="button">Todos</button><button class="quick-filter quick-filter-danger" data-quick-status="Em atraso" type="button">Inadimplentes</button><button class="quick-filter quick-filter-orange" data-quick-status="A creditar" type="button">Aguardando crédito</button><button class="quick-filter quick-filter-green" data-quick-status="Creditado" type="button">Já creditados</button></div>
            <div class="advanced-filters">
              <label class="search-box" aria-label="Buscar participante">${icons.search}<input id="search-input" type="search" placeholder="Buscar nome ou parcela…" /></label>
              <label class="filter-field"><span>Aluno</span><select class="select" id="donor-filter"><option>Todas as pessoas</option>${[...new Set(state.records.map((item) => item.donor))].sort((a, b) => a.localeCompare(b, "pt-BR")).map((name) => `<option value="${safe(name)}">${safe(name)}</option>`).join("")}</select></label>
              <label class="filter-field"><span>Pagamento</span><select class="select" id="method-filter" aria-label="Filtrar por meio de pagamento"><option>Todos os meios</option>${[...new Set(state.records.map((item) => item.method))].sort().map((method) => `<option>${safe(method)}</option>`).join("")}</select></label>
              <label class="filter-field"><span>Situação</span><select class="select" id="status-filter" aria-label="Filtrar por status"><option>Todos os status</option><option>Creditado</option><option>A creditar</option><option>Em atraso</option><option>Pendente</option></select></label>
              <label class="filter-field"><span>Parcela</span><select class="select" id="installment-filter"><option>Todas as parcelas</option>${[...new Set(state.records.map((item) => item.installment))].sort((a, b) => a.localeCompare(b, "pt-BR", { numeric: true })).map((installment) => `<option>${safe(installment)}</option>`).join("")}</select></label>
              <label class="filter-field date-field"><span>Vencimento inicial</span><input class="date-input" id="date-from" type="date" /></label><label class="filter-field date-field"><span>Vencimento final</span><input class="date-input" id="date-to" type="date" /></label>
            </div>
            <div id="table-region"></div>
          </article>
        </section>

        <section class="ministry-section app-view" id="ministry" data-view="ministry">
          <div class="table-top ministry-table-top">
            <div>
              <p class="eyebrow">Auditoria das saídas</p>
              <h2 class="section-title">Ministério e despesas</h2>
              <p class="panel-subtitle">Destino dos gastos, interpretação e lançamentos administrativos em uma única área.</p>
            </div>
          </div>
          <div class="ministry-source-note">
            <span>${icons.money}</span>
            <p><strong>Escopo:</strong> entradas administrativas, despesas gerais e despesas das aulas. Os créditos dos alunos permanecem na área do curso.</p>
          </div>
          ${renderMinistryAuditContent(ministrySummary)}
          <article class="card ministry-ledger-card">
            <div class="table-top ministry-table-top"><div><p class="eyebrow">Base administrativa</p><h3 class="panel-title">Todos os lançamentos do Ministério</h3><p class="panel-subtitle">Use os filtros para conferir pessoas, categorias, datas e motivos.</p></div><button class="button" id="clear-ministry-filters" type="button">Limpar filtros</button></div>
            <div class="advanced-filters ministry-filters">
              <label class="search-box" aria-label="Buscar nas finanças do Ministério">${icons.search}<input id="ministry-search" type="search" placeholder="Buscar categoria, pessoa ou descrição…" /></label>
              <label class="filter-field"><span>Núcleo</span><select class="select" id="ministry-kind-filter"><option value="">Todos os núcleos</option>${[...new Set(state.ministryRecords.map((item) => item.financeKind))].sort().map((kind) => `<option value="${safe(kind)}">${safe(kind)}</option>`).join("")}</select></label>
              <label class="filter-field"><span>Tipo</span><select class="select" id="ministry-type-filter"><option value="">Todos os tipos</option>${[...new Set(state.ministryRecords.map((item) => item.type))].sort().map((type) => `<option value="${safe(type)}">${safe(type)}</option>`).join("")}</select></label>
              <label class="filter-field"><span>Categoria</span><select class="select" id="ministry-category-filter"><option value="">Todas as categorias</option>${[...new Set(state.ministryRecords.map((item) => item.category))].sort((a, b) => a.localeCompare(b, "pt-BR")).map((category) => `<option value="${safe(category)}">${safe(category)}</option>`).join("")}</select></label>
              <label class="filter-field date-field"><span>Data inicial</span><input class="date-input" id="ministry-date-from" type="date" /></label><label class="filter-field date-field"><span>Data final</span><input class="date-input" id="ministry-date-to" type="date" /></label>
            </div>
            <div id="ministry-table-region"></div>
          </article>
        </section>

        ${renderOperationalCashflow(summary, ministrySummary)}
      </main>
    </div>`;

  bindAppEvents();
  applyFilters();
  applyMinistryFilters();
  renderPresentation();
  activateView(state.activeView);
}

function bindAppEvents() {
  const appRoot = document.querySelector("#app");
  const sidebar = document.querySelector(".sidebar");
  const mobileMenuToggle = document.querySelector(".mobile-menu-toggle");
  const closeMobileMenu = () => {
    sidebar?.classList.remove("mobile-menu-open");
    mobileMenuToggle?.setAttribute("aria-expanded", "false");
    mobileMenuToggle?.setAttribute("aria-label", "Abrir menu de navegação");
  };
  mobileMenuToggle?.addEventListener("click", () => {
    const willOpen = !sidebar.classList.contains("mobile-menu-open");
    sidebar.classList.toggle("mobile-menu-open", willOpen);
    mobileMenuToggle.setAttribute("aria-expanded", String(willOpen));
    mobileMenuToggle.setAttribute("aria-label", willOpen ? "Fechar menu de navegação" : "Abrir menu de navegação");
  });
  document.querySelector(".mobile-menu-backdrop")?.addEventListener("click", closeMobileMenu);
  appRoot.addEventListener("keydown", (event) => {
    if (event.key !== "Escape" || !sidebar?.classList.contains("mobile-menu-open")) return;
    closeMobileMenu();
    mobileMenuToggle?.focus();
  });
  const openAuditFromEvent = (event) => {
    const trigger = event.target.closest?.("[data-audit-key]");
    if (!trigger || !appRoot.contains(trigger)) return;
    if (event.type === "keydown" && !["Enter", " "].includes(event.key)) return;
    if (event.type === "keydown") event.preventDefault();
    openAuditBreakdown(trigger.dataset.auditKey);
  };
  appRoot.addEventListener("click", openAuditFromEvent);
  appRoot.addEventListener("keydown", openAuditFromEvent);
  appRoot.addEventListener("keydown", (event) => {
    if (state.activeView !== "presentation" || !["ArrowLeft", "ArrowRight"].includes(event.key)) return;
    if (document.querySelector(".audit-breakdown-backdrop")) return;
    const pages = buildPresentationPages();
    const direction = event.key === "ArrowRight" ? 1 : -1;
    const nextPage = Math.min(Math.max(state.presentationPage + direction, 0), pages.length - 1);
    if (nextPage === state.presentationPage) return;
    event.preventDefault();
    state.presentationPage = nextPage;
    renderPresentation();
  });

  document.querySelectorAll(".nav-item").forEach((button) => {
    button.addEventListener("click", () => {
      activateView(button.dataset.target, true);
      closeMobileMenu();
    });
  });

  document.querySelector("#open-ministry-view")?.addEventListener("click", () => activateView("ministry", true));

  document.querySelector("#ministry-search").addEventListener("input", (event) => {
    state.ministrySearch = event.target.value.trim().toLocaleLowerCase("pt-BR");
    applyMinistryFilters();
  });

  document.querySelector("#ministry-kind-filter").addEventListener("change", (event) => {
    state.ministryKind = event.target.value || "Todos";
    applyMinistryFilters();
  });

  document.querySelector("#ministry-type-filter").addEventListener("change", (event) => {
    state.ministryType = event.target.value || "Todos";
    applyMinistryFilters();
  });

  document.querySelector("#ministry-category-filter").addEventListener("change", (event) => {
    state.ministryCategory = event.target.value || "Todas";
    applyMinistryFilters();
  });

  document.querySelector("#ministry-date-from").addEventListener("change", (event) => {
    state.ministryDateFrom = event.target.value;
    applyMinistryFilters();
  });

  document.querySelector("#ministry-date-to").addEventListener("change", (event) => {
    state.ministryDateTo = event.target.value;
    applyMinistryFilters();
  });

  document.querySelector("#clear-ministry-filters").addEventListener("click", resetMinistryFilters);

  document.querySelector("#search-input").addEventListener("input", (event) => {
    state.search = event.target.value.trim().toLocaleLowerCase("pt-BR");
    state.page = 1;
    applyFilters();
  });

  document.querySelector("#donor-filter").addEventListener("change", (event) => {
    state.donor = event.target.selectedIndex === 0 ? "Todos" : event.target.value;
    state.page = 1;
    applyFilters();
  });

  document.querySelector("#method-filter").addEventListener("change", (event) => {
    state.method = event.target.selectedIndex === 0 ? "Todos" : event.target.value;
    state.page = 1;
    applyFilters();
  });

  document.querySelector("#status-filter").addEventListener("change", (event) => {
    state.status = event.target.selectedIndex === 0 ? "Todos" : event.target.value;
    state.page = 1;
    document.querySelectorAll("[data-quick-status]").forEach((button) => button.classList.toggle("active", button.dataset.quickStatus === state.status));
    applyFilters();
  });

  document.querySelector("#installment-filter").addEventListener("change", (event) => {
    state.installment = event.target.selectedIndex === 0 ? "Todas" : event.target.value;
    state.page = 1;
    applyFilters();
  });

  document.querySelector("#date-from").addEventListener("change", (event) => {
    state.dateFrom = event.target.value;
    state.page = 1;
    applyFilters();
  });

  document.querySelector("#date-to").addEventListener("change", (event) => {
    state.dateTo = event.target.value;
    state.page = 1;
    applyFilters();
  });

  document.querySelectorAll("[data-quick-status]").forEach((button) => {
    button.addEventListener("click", () => {
      state.status = button.dataset.quickStatus;
      state.page = 1;
      document.querySelectorAll("[data-quick-status]").forEach((item) => item.classList.toggle("active", item === button));
      const statusSelect = document.querySelector("#status-filter");
      statusSelect.selectedIndex = [...statusSelect.options].findIndex((option, index) =>
        state.status === "Todos" ? index === 0 : option.value === state.status,
      );
      applyFilters();
    });
  });

  document.querySelector("#clear-filters").addEventListener("click", resetTableFilters);

  document.querySelector("#participant-search").addEventListener("input", (event) => {
    state.participantSearch = event.target.value.trim().toLocaleLowerCase("pt-BR");
    state.participantPage = 1;
    renderParticipants();
  });

  document.querySelector("#participant-name").addEventListener("change", (event) => {
    state.participantName = event.target.selectedIndex === 0 ? "Todos" : event.target.value;
    state.participantPage = 1;
    renderParticipants();
  });

  document.querySelector("#participant-status").addEventListener("change", (event) => {
    state.participantStatus = event.target.selectedIndex === 0 ? "Todos" : event.target.value;
    state.participantPage = 1;
    renderParticipants();
  });

  document.querySelector("#import-button").addEventListener("click", () => document.querySelector("#file-input").click());
  document.querySelector("#file-input").addEventListener("change", importWorkbook);
  document.querySelector("#export-button").addEventListener("click", exportCsv);
}

function activateView(viewId, resetScroll = false) {
  if (!viewTitles[viewId]) return;
  if (viewId !== "presentation" && document.fullscreenElement === document.querySelector("#presentation")) {
    document.exitFullscreen?.();
  }
  state.activeView = viewId;
  document.querySelector(".main-content")?.classList.toggle("presentation-mode", viewId === "presentation");
  document.querySelectorAll(".app-view").forEach((section) => {
    section.hidden = section.dataset.view !== viewId;
  });
  document.querySelectorAll(".nav-item").forEach((button) => {
    const isActive = button.dataset.target === viewId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-current", isActive ? "page" : "false");
  });
  const pageTitle = document.querySelector("#page-title");
  if (pageTitle) pageTitle.textContent = viewTitles[viewId];
  if (resetScroll) {
    document.documentElement.scrollTop = 0;
    document.body.scrollTop = 0;
  }
}

function applyMinistryFilters() {
  state.filteredMinistry = state.ministryRecords.filter((record) => {
    const searchText = [
      record.category,
      record.name,
      record.specification,
      record.description,
      record.eventTitle,
      record.createdBy,
    ].join(" ").toLocaleLowerCase("pt-BR");
    const matchesSearch = !state.ministrySearch || searchText.includes(state.ministrySearch);
    const matchesKind = state.ministryKind === "Todos" || record.financeKind === state.ministryKind;
    const matchesType = state.ministryType === "Todos" || record.type === state.ministryType;
    const matchesCategory = state.ministryCategory === "Todas" || record.category === state.ministryCategory;
    const matchesFrom = !state.ministryDateFrom || record.date >= state.ministryDateFrom;
    const matchesTo = !state.ministryDateTo || record.date <= state.ministryDateTo;
    return matchesSearch && matchesKind && matchesType && matchesCategory && matchesFrom && matchesTo;
  });
  renderMinistryTable();
}

function resetMinistryFilters() {
  state.ministrySearch = "";
  state.ministryKind = "Todos";
  state.ministryType = "Todos";
  state.ministryCategory = "Todas";
  state.ministryDateFrom = "";
  state.ministryDateTo = "";
  document.querySelector("#ministry-search").value = "";
  document.querySelector("#ministry-kind-filter").selectedIndex = 0;
  document.querySelector("#ministry-type-filter").selectedIndex = 0;
  document.querySelector("#ministry-category-filter").selectedIndex = 0;
  document.querySelector("#ministry-date-from").value = "";
  document.querySelector("#ministry-date-to").value = "";
  applyMinistryFilters();
}

function renderMinistryTable() {
  const region = document.querySelector("#ministry-table-region");
  if (!region) return;
  const rows = state.filteredMinistry;
  const totals = summarizeMinistry(state.filteredMinistry);
  const body = rows.length
    ? rows.map((record) => {
        const details = record.specification || record.description || record.eventTitle || "—";
        const destination = record.name || record.eventTitle || record.category;
        return `<tr>
          <td>${formatIsoDate(record.date)}</td>
          <td><span class="ministry-kind">${safe(record.financeKind)}</span></td>
          <td><span class="ministry-type ministry-type-${record.type.toLocaleLowerCase("pt-BR")}">${safe(record.type)}</span></td>
          <td><strong class="ministry-category">${safe(record.category)}</strong></td>
          <td><span class="ministry-destination" title="${safe(destination)}">${safe(destination)}</span></td>
          <td><span class="ministry-description" title="${safe(details)}">${safe(details)}</span></td>
          <td>${safe(record.paymentMethod)}</td>
          <td>${safe(record.createdBy)}</td>
          <td class="money money-pending">${currency.format(record.amount)}</td>
        </tr>`;
      }).join("")
    : `<tr><td colspan="9" class="empty-state">Nenhum lançamento corresponde aos filtros.</td></tr>`;

  region.innerHTML = `
    <div class="table-wrap ministry-table-wrap">
      <table class="ministry-table">
        <thead><tr><th>Data</th><th>Núcleo</th><th>Tipo</th><th>Categoria</th><th>Nome / destino</th><th>Detalhes</th><th>Pagamento</th><th>Registrado por</th><th>Valor</th></tr></thead>
        <tbody>${body}</tbody>
        ${rows.length ? `<tfoot><tr><td colspan="6">Totais do filtro</td><td colspan="2">Entradas: ${currency.format(totals.entries)} · Saídas: ${currency.format(totals.exits)}</td><td>${currency.format(totals.entries - totals.exits)}</td></tr></tfoot>` : ""}
      </table>
    </div>
    <footer class="table-footer ministry-all-rows-note">
      <span>Exibindo todos os ${number.format(state.filteredMinistry.length)} lançamentos filtrados, sem divisão em páginas.</span>
    </footer>`;
}

function applyFilters() {
  state.filtered = state.records.filter((record) => {
    const matchesSearch = !state.search || `${record.donor} ${record.installment}`.toLocaleLowerCase("pt-BR").includes(state.search);
    const matchesDonor = state.donor === "Todos" || record.donor === state.donor;
    const matchesMethod = state.method === "Todos" || record.method === state.method;
    const matchesStatus = state.status === "Todos" || record.status === state.status;
    const matchesInstallment = state.installment === "Todas" || record.installment === state.installment;
    const dueDate = parseDate(record.dueDate);
    const fromDate = state.dateFrom ? new Date(`${state.dateFrom}T00:00:00`) : null;
    const toDate = state.dateTo ? new Date(`${state.dateTo}T23:59:59`) : null;
    const matchesDateFrom = !fromDate || (dueDate && dueDate >= fromDate);
    const matchesDateTo = !toDate || (dueDate && dueDate <= toDate);
    return matchesSearch && matchesDonor && matchesMethod && matchesStatus && matchesInstallment && matchesDateFrom && matchesDateTo;
  });
  const totalPages = Math.max(Math.ceil(state.filtered.length / state.perPage), 1);
  state.page = Math.min(state.page, totalPages);
  renderTable();
  renderParticipants();
}

function resetTableFilters() {
  state.search = "";
  state.donor = "Todos";
  state.method = "Todos";
  state.status = "Todos";
  state.installment = "Todas";
  state.dateFrom = "";
  state.dateTo = "";
  state.page = 1;
  document.querySelector("#search-input").value = "";
  document.querySelector("#donor-filter").selectedIndex = 0;
  document.querySelector("#method-filter").selectedIndex = 0;
  document.querySelector("#status-filter").selectedIndex = 0;
  document.querySelector("#installment-filter").selectedIndex = 0;
  document.querySelector("#date-from").value = "";
  document.querySelector("#date-to").value = "";
  document.querySelectorAll("[data-quick-status]").forEach((button) => button.classList.toggle("active", button.dataset.quickStatus === "Todos"));
  applyFilters();
}

function renderParticipants() {
  const region = document.querySelector("#participant-region");
  if (!region) return;
  const participants = groupByParticipant(state.records).filter((person) => {
    const matchesSearch = !state.participantSearch || person.name.toLocaleLowerCase("pt-BR").includes(state.participantSearch);
    const matchesName = state.participantName === "Todos" || person.name === state.participantName;
    const matchesStatus = state.participantStatus === "Todos" || person.status === state.participantStatus;
    return matchesSearch && matchesName && matchesStatus;
  });
  const pageRows = participants;
  const filteredTotals = participants.length
    ? participants.reduce(
        (totals, person) => {
          totals.paid += person.paid;
          totals.credited += person.credited;
          totals.awaitingCredit += person.awaitingCredit;
          totals.overdue += person.overdue;
          return totals;
        },
        { paid: 0, credited: 0, awaitingCredit: 0, overdue: 0 },
      )
    : { paid: 0, credited: 0, awaitingCredit: 0, overdue: 0 };

  const body = pageRows.length
    ? pageRows
        .map(
          (person) => `<tr>
            <td><div class="person-cell"><span class="avatar">${initials(person.name)}</span><button class="person-name-button" data-participant="${safe(person.name)}" type="button">${safe(person.name)}</button></div></td>
            <td><div class="installment-progress"><strong>${person.paidInstallments} de ${person.totalInstallments}</strong><span>pagas</span></div></td>
            <td class="money paid-informative">${currency.format(person.paid)}</td>
            <td class="money money-positive credited-highlight">${currency.format(person.credited)}</td>
            <td class="money">${currency.format(person.awaitingCredit)}</td>
            <td class="money ${person.overdue ? "money-pending" : ""}">${currency.format(person.overdue)}</td>
            <td><span class="status-pill ${participantStatusClass(person.status)}">${person.status}</span></td>
            <td><button class="row-action" data-participant="${safe(person.name)}" type="button" aria-label="Ver parcelas de ${safe(person.name)}">${icons.arrow}</button></td>
          </tr>`,
        )
        .join("")
    : `<tr><td colspan="8" class="empty-state">Nenhuma pessoa corresponde aos filtros.</td></tr>`;

  region.innerHTML = `
    <div class="table-wrap people-table-wrap">
      <table class="people-table">
        <thead><tr><th>Aluno</th><th>Parcelas pagas</th><th>Pago pelo aluno</th><th>Creditado líquido</th><th>Ainda vai cair</th><th>Devendo</th><th>Situação</th><th></th></tr></thead>
        <tbody>${body}</tbody>
        ${pageRows.length ? `<tfoot><tr><td colspan="2">Total do filtro</td><td>${currency.format(filteredTotals.paid)}</td><td>${currency.format(filteredTotals.credited)}</td><td>${currency.format(filteredTotals.awaitingCredit)}</td><td>${currency.format(filteredTotals.overdue)}</td><td colspan="2"></td></tr></tfoot>` : ""}
      </table>
    </div>
    <footer class="table-footer">
      <span>Exibindo todas as ${number.format(participants.length)} pessoas que correspondem aos filtros.</span>
    </footer>`;

  region.querySelectorAll("[data-participant]").forEach((button) => {
    button.addEventListener("click", () => openParticipantDrawer(button.dataset.participant));
  });
}

function getParticipantPageButtons(totalPages) {
  if (totalPages <= 4) return Array.from({ length: totalPages }, (_, index) => index + 1);
  if (state.participantPage <= 2) return [1, 2, 3, "…", totalPages];
  if (state.participantPage >= totalPages - 1) return [1, "…", totalPages - 2, totalPages - 1, totalPages];
  return [1, "…", state.participantPage, "…", totalPages];
}

function renderTable() {
  const region = document.querySelector("#table-region");
  const rows = state.filtered;

  const body = rows.length
    ? rows
        .map(
          (record) => `<tr>
            <td><div class="person-cell"><span class="avatar">${initials(record.donor)}</span><span class="person-name" title="${safe(record.donor)}">${safe(record.donor)}</span></div></td>
            <td>${safe(record.method)}</td>
            <td>${safe(record.dueDate || "—")}</td>
            <td>${safe(record.receivedAt || "—")}</td>
            <td>${safe(record.expectedCreditAt || "—")}</td>
            <td>${safe(record.creditedAt || "—")}</td>
            <td>${safe(record.installment)}</td>
            <td class="money">${currency.format(record.receivable)}</td>
            <td class="money paid-informative">${currency.format(record.paid)}</td>
            <td class="money money-positive credited-highlight">${currency.format(record.credited)}</td>
            <td class="money">${currency.format(Math.abs(record.fee))}</td>
            <td><span class="status-pill ${statusClass(record.status)}">${record.status}</span></td>
            <td><button class="row-action" data-record-id="${record.id}" type="button" aria-label="Ver detalhes de ${safe(record.donor)}">${icons.arrow}</button></td>
          </tr>`,
        )
        .join("")
    : `<tr><td colspan="13" class="empty-state">Nenhum lançamento corresponde aos filtros.</td></tr>`;

  region.innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Aluno</th><th>Pagamento</th><th>Vencimento</th><th>Data do pagamento</th><th>Crédito previsto</th><th>Crédito efetivo</th><th>Parcela</th><th>Valor a receber</th><th>Valor pago</th><th>Creditado líquido</th><th>Tarifa</th><th>Status</th><th></th></tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
    <footer class="table-footer">
      <span>Exibindo todos os ${number.format(state.filtered.length)} lançamentos que correspondem aos filtros.</span>
    </footer>`;

  region.querySelectorAll("[data-record-id]").forEach((button) => {
    button.addEventListener("click", () => openDrawer(button.dataset.recordId));
  });
}

function getPageButtons(totalPages) {
  if (totalPages <= 5) return Array.from({ length: totalPages }, (_, index) => index + 1);
  if (state.page <= 3) return [1, 2, 3, 4, "…", totalPages];
  if (state.page >= totalPages - 2) return [1, "…", totalPages - 3, totalPages - 2, totalPages - 1, totalPages];
  return [1, "…", state.page - 1, state.page, state.page + 1, "…", totalPages];
}

function courseAuditBreakdown(key) {
  const now = new Date();
  const requestedMonths = new Set([4, 5, 6, 7, 8, 9, 10]);
  const definitions = {
    "course-all": {
      title: "Base completa do Curso de Libras",
      valueLabel: "Valor cadastrado nas linhas",
      formula: "Todos os lançamentos válidos da planilha, após as inclusões e exclusões solicitadas.",
      relevance: "Define o tamanho financeiro total da turma e ajuda a conferir se todas as parcelas contratadas foram cadastradas. Não representa dinheiro disponível na conta.",
      predicate: () => true,
      value: (record) => record.receivable,
    },
    "course-credited": {
      title: "Recebido líquido pelo Ministério",
      valueLabel: "Valor Creditado",
      formula: "Valor Pago − tarifas da operadora de cartão − parcelas pagas ainda não creditadas = Valor Creditado líquido.",
      relevance: "É o dinheiro líquido realmente disponível para o Ministério. Por isso, este é o número correto para comparar a receita do curso com o custo das aulas.",
      predicate: (record) => record.credited > 0,
      value: (record) => record.credited,
    },
    "course-awaiting": {
      title: "Pago, aguardando crédito",
      valueLabel: "Valor Pago sem crédito",
      formula: "Soma do Valor Pago quando Valor Pago > 0 e Valor Creditado = 0. São pagamentos no cartão que serão creditados posteriormente pela operadora.",
      relevance: "Mostra dinheiro que o aluno já pagou, mas que a operadora ainda não repassou. Ele ajuda a prever o caixa sem confundir pagamento com saldo em conta.",
      predicate: (record) => record.paid > 0 && record.credited === 0,
      value: (record) => record.paid,
    },
    "course-future": {
      title: "Previsão de crédito de setembro e outubro",
      valueLabel: "Valor previsto",
      formula: "Soma do Valor Pago — ou do Valor a Receber quando ainda não houve pagamento — com crédito previsto para setembro ou outubro e Valor Creditado = 0.",
      relevance: "Ajuda a planejar as entradas dos próximos meses, mantendo a previsão separada do dinheiro já recebido.",
      predicate: (record) => {
        const date = parseDate(record.expectedCreditAt);
        return Boolean(date && date >= forecastStartsAt && [9, 10].includes(date.getMonth() + 1) && record.credited === 0);
      },
      value: (record) => (record.paid > 0 ? record.paid : record.receivable),
    },
    "course-reconcile": {
      title: "Crédito vencido a conciliar",
      valueLabel: "Valor Pago sem crédito",
      formula: "Soma do Valor Pago com previsão de crédito já encerrada, Valor Creditado = 0 e competência entre abril e outubro.",
      relevance: "Sinaliza pagamentos cujo prazo de repasse terminou sem crédito registrado e que precisam ser conferidos com a operadora do cartão.",
      predicate: (record) => {
        const date = parseDate(record.expectedCreditAt);
        return Boolean(date && date < forecastStartsAt && requestedMonths.has(date.getMonth() + 1) && record.paid > 0 && record.credited === 0);
      },
      value: (record) => record.paid,
    },
    "course-overdue": {
      title: "Inadimplência vencida",
      valueLabel: "Saldo em atraso",
      formula: "Por parcela vencida: Valor a Receber − Valor Pago. Entram somente saldos positivos.",
      relevance: "Mede a receita vencida que não entrou por falta de pagamento e orienta as cobranças aos alunos.",
      predicate: (record) => {
        const date = parseDate(record.dueDate);
        return Boolean(date && date < now && record.receivable - record.paid > 0);
      },
      value: (record) => Math.max(record.receivable - record.paid, 0),
    },
    "course-receivable": {
      title: "Valor total cadastrado",
      valueLabel: "Valor a Receber",
      formula: "Soma da coluna Valor a Receber de todos os lançamentos válidos.",
      relevance: "Mostra o valor bruto contratado com os alunos e funciona como referência do potencial da turma, não como saldo bancário.",
      predicate: () => true,
      value: (record) => record.receivable,
    },
    "course-paid": {
      title: "Pago pelos alunos",
      valueLabel: "Valor Pago",
      formula: "Soma da coluna Valor Pago. É um valor informativo, antes dos descontos da operadora de cartão.",
      relevance: "Permite conferir o que os alunos efetivamente pagaram e separar esse momento do repasse líquido feito pela operadora.",
      predicate: (record) => record.paid > 0,
      value: (record) => record.paid,
    },
    "course-fees": {
      title: "Tarifas e taxas do cartão",
      valueLabel: "Tarifa cobrada",
      formula: "Soma absoluta da coluna Despesa Financeira: tarifas e taxas de manutenção e processamento dos pagamentos no cartão.",
      relevance: "Explica a diferença entre o valor pago pelo aluno e o líquido recebido pelo Ministério, além de revelar o custo de usar o cartão.",
      predicate: (record) => Math.abs(record.fee) > 0,
      value: (record) => Math.abs(record.fee),
    },
  };

  let definition = definitions[key];
  if (key.startsWith("defaulter:")) {
    let personName = "";
    try {
      personName = decodeURIComponent(key.slice("defaulter:".length));
    } catch {
      personName = key.slice("defaulter:".length);
    }
    definition = {
      title: `Inadimplência de ${personName}`,
      valueLabel: "Saldo em atraso",
      formula: "Parcelas desta pessoa com vencimento anterior à auditoria: Valor a Receber − Valor Pago.",
      relevance: "Mostra exatamente quais parcelas desta pessoa precisam de cobrança e evita confundir atraso do aluno com atraso de repasse do cartão.",
      predicate: (record) => {
        const date = parseDate(record.dueDate);
        return record.donor === personName && Boolean(date && date < now && record.receivable - record.paid > 0);
      },
      value: (record) => Math.max(record.receivable - record.paid, 0),
    };
  }
  if (!definition) return null;

  const records = state.records
    .filter(definition.predicate)
    .sort((a, b) => a.donor.localeCompare(b.donor, "pt-BR") || (parseDate(a.dueDate)?.getTime() || 0) - (parseDate(b.dueDate)?.getTime() || 0));
  const total = records.reduce((sum, record) => sum + definition.value(record), 0);
  const displayKey = key.startsWith("defaulter:") ? "course-overdue" : key;
  const tableFormats = {
    "course-all": {
      columns: ["Aluno", "Parcela", "Vencimento", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.dueDate || "—", definition.value(record)],
    },
    "course-credited": {
      columns: ["Aluno", "Parcela", "Data do crédito", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.creditedAt || "Sem data", definition.value(record)],
    },
    "course-awaiting": {
      columns: ["Aluno", "Parcela", "Pagamento", "Crédito previsto", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.method || "—", record.expectedCreditAt || "Sem previsão", definition.value(record)],
    },
    "course-future": {
      columns: ["Aluno", "Parcela", "Crédito previsto", "Base usada", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.expectedCreditAt || "Sem previsão", record.paid > 0 ? "Valor Pago" : "Valor a Receber", definition.value(record)],
    },
    "course-reconcile": {
      columns: ["Aluno", "Parcela", "Crédito que venceu", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.expectedCreditAt || "Sem previsão", definition.value(record)],
    },
    "course-overdue": {
      columns: ["Aluno", "Parcela", "Vencimento", "A receber", "Pago", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.dueDate || "—", record.receivable, record.paid, definition.value(record)],
    },
    "course-receivable": {
      columns: ["Aluno", "Parcela", "Vencimento", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.dueDate || "—", definition.value(record)],
    },
    "course-paid": {
      columns: ["Aluno", "Parcela", "Data do pagamento", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.receivedAt || "Sem data", definition.value(record)],
    },
    "course-fees": {
      columns: ["Aluno", "Parcela", "Valor Pago", definition.valueLabel],
      row: (record) => [record.donor, record.installment || "—", record.paid, definition.value(record)],
    },
  };
  const tableFormat = tableFormats[displayKey];
  const sumRecords = (selector) => records.reduce((sum, record) => sum + selector(record), 0);
  const calculations = {
    "course-all": [{ label: `${number.format(records.length)} parcelas cadastradas`, value: total }],
    "course-credited": [{ label: `${number.format(records.length)} créditos líquidos`, value: total }],
    "course-awaiting": [
      { label: "Pago no cartão", value: sumRecords((record) => record.paid) },
      { operator: "−", label: "Já creditado nessas linhas", value: sumRecords((record) => record.credited) },
    ],
    "course-future": [
      { label: "Parcelas já pagas", value: sumRecords((record) => record.paid > 0 ? record.paid : 0) },
      { operator: "+", label: "Parcelas ainda a pagar", value: sumRecords((record) => record.paid > 0 ? 0 : record.receivable) },
    ],
    "course-reconcile": [
      { label: "Pago com previsão vencida", value: sumRecords((record) => record.paid) },
      { operator: "−", label: "Creditado nessas linhas", value: sumRecords((record) => record.credited) },
    ],
    "course-overdue": [
      { label: "A receber nas parcelas vencidas", value: sumRecords((record) => record.receivable) },
      { operator: "−", label: "Pago nessas parcelas", value: sumRecords((record) => record.paid) },
    ],
    "course-receivable": [{ label: `${number.format(records.length)} valores cadastrados`, value: total }],
    "course-paid": [{ label: `${number.format(records.length)} pagamentos informados`, value: total }],
    "course-fees": [{ label: `${number.format(records.length)} tarifas cobradas`, value: total }],
  };
  const lineContexts = {
    "course-all": (record) => `Vencimento: ${record.dueDate || "não informado"}`,
    "course-credited": (record) => `Creditado em ${record.creditedAt || "data não informada"}`,
    "course-awaiting": (record) => `${record.method || "Pagamento"} · crédito previsto em ${record.expectedCreditAt || "data não informada"}`,
    "course-future": (record) => `Crédito previsto em ${record.expectedCreditAt || "data não informada"} · base: ${record.paid > 0 ? "Valor Pago" : "Valor a Receber"}`,
    "course-reconcile": (record) => `Previsão vencida em ${record.expectedCreditAt || "data não informada"}`,
    "course-overdue": (record) => `Venceu em ${record.dueDate || "data não informada"} · ${currency.format(record.receivable)} − ${currency.format(record.paid)}`,
    "course-receivable": (record) => `Vencimento: ${record.dueDate || "não informado"}`,
    "course-paid": (record) => `Pago em ${record.receivedAt || "data não informada"}`,
    "course-fees": (record) => `Tarifa sobre o pagamento de ${currency.format(record.paid)}`,
  };
  const groupingRecords = displayKey === "course-credited"
    ? state.records
      .filter((record) => record.paid > 0)
      .sort((a, b) => a.donor.localeCompare(b.donor, "pt-BR") || (parseDate(a.dueDate)?.getTime() || 0) - (parseDate(b.dueDate)?.getTime() || 0))
    : records;
  const groupMap = new Map();
  groupingRecords.forEach((record) => {
    if (!groupMap.has(record.donor)) {
      groupMap.set(record.donor, {
        label: record.donor,
        total: 0,
        paid: 0,
        fees: 0,
        awaitingCredit: 0,
        lines: [],
      });
    }
    const group = groupMap.get(record.donor);
    const value = displayKey === "course-credited" ? record.credited : definition.value(record);
    group.total += value;
    group.paid += record.paid;
    group.fees += Math.abs(record.fee);
    if (record.paid > 0 && record.credited === 0) group.awaitingCredit += record.paid;
    group.lines.push({
      label: `Parcela ${record.installment || "—"}`,
      context: displayKey === "course-credited"
        ? `${currency.format(record.paid)} pago − ${currency.format(Math.abs(record.fee))} de taxa − ${currency.format(record.credited === 0 ? record.paid : 0)} ainda não creditado · crédito em ${record.creditedAt || "aguardando"}`
        : lineContexts[displayKey](record),
      value,
    });
  });
  const groups = [...groupMap.values()]
    .map((group) => displayKey === "course-credited" ? {
      ...group,
      totalLabel: "Líquido recebido",
      calculation: [
        { label: "Total pago", value: group.paid },
        { operator: "−", label: "Taxas", value: group.fees },
        { operator: "−", label: "Ainda não creditado", value: group.awaitingCredit },
      ],
    } : group)
    .sort((a, b) => b.total - a.total || a.label.localeCompare(b.label, "pt-BR"));

  return {
    title: definition.title,
    source: state.sourceName,
    formula: definition.formula,
    relevance: definition.relevance,
    valueLabel: definition.valueLabel,
    total,
    calculation: calculations[displayKey],
    resultLabel: definition.title,
    groupTitle: "Composição por pessoa",
    groupHint: displayKey === "course-credited"
      ? "Em cada pessoa: total pago − taxas do cartão − parcelas ainda sem crédito = líquido que entrou na conta. Abra um nome para conferir as parcelas."
      : "Confira o subtotal de cada pessoa. Abra um nome somente se precisar ver as parcelas.",
    groups,
    columns: tableFormat.columns,
    rows: records.map((record) => tableFormat.row(record).map((cell, index, row) => {
      const isValue = typeof cell === "number";
      const content = isValue ? currency.format(cell) : safe(cell);
      return index === 0 || isValue && index === row.length - 1 ? `<strong>${content}</strong>` : content;
    })),
    factCount: groupingRecords.length,
    note: displayKey === "course-credited"
      ? `${number.format(groupingRecords.length)} pagamentos foram considerados para demonstrar o que já entrou e o que ainda aguarda crédito da operadora.`
      : `${number.format(records.length)} ${records.length === 1 ? "linha encontrada" : "linhas encontradas"} na base do curso.`,
  };
}

function ministryAuditBreakdown(key) {
  if (key.startsWith("ministry-category:")) {
    const [, encodedKind = "", encodedCategory = ""] = key.split(":");
    const kind = decodeURIComponent(encodedKind);
    const category = decodeURIComponent(encodedCategory);
    const records = state.ministryRecords.filter((record) =>
      record.type.toLocaleLowerCase("pt-BR") === "saida" &&
      record.financeKind === kind &&
      record.category === category,
    );
    const total = records.reduce((sum, record) => sum + record.amount, 0);
    return {
      title: `Despesas de ${category}`,
      source: "finançassinaisdoreino.json",
      formula: `Soma das saídas classificadas no núcleo “${kind}” e na categoria “${category}”.`,
      relevance: "O detalhamento permite conferir quem recebeu, quando o pagamento ocorreu e qual motivo foi registrado.",
      valueLabel: "Valor da despesa",
      total,
      calculation: [{ label: `${number.format(records.length)} lançamentos encontrados`, value: total }],
      resultLabel: `Total de ${category}`,
      groupTitle: "Lançamentos da categoria",
      groupHint: "Cada linha abaixo participa integralmente do total desta categoria.",
      groups: [{
        label: category,
        total,
        lines: records.map((record) => ({
          label: record.name || record.eventTitle || record.category,
          context: `${formatIsoDate(record.date)} · ${record.specification || record.description || "sem descrição complementar"}`,
          value: record.amount,
        })),
      }],
      columns: ["Data", "Núcleo", "Destino", "Motivo registrado", "Valor"],
      rows: records.map((record) => [
        formatIsoDate(record.date),
        record.financeKind,
        record.name || record.eventTitle || record.category,
        record.specification || record.description || "Sem descrição complementar",
        record.amount,
      ]),
      note: `${number.format(records.length)} ${records.length === 1 ? "lançamento" : "lançamentos"} formam o total da categoria.`,
    };
  }
  const definitions = {
    "ministry-all": {
      title: "Todos os lançamentos administrativos",
      valueLabel: "Valor das linhas",
      formula: "Todos os registros válidos do JSON e as entradas administrativas recorrentes cadastradas no sistema.",
      relevance: "Apresenta o volume total registrado na base do Ministério e permite conferir a integridade dos lançamentos administrativos.",
      source: "finançassinaisdoreino.json + entradas administrativas cadastradas",
      predicate: () => true,
    },
    "ministry-exits": {
      title: "Total de saídas registradas",
      valueLabel: "Valor da saída",
      formula: "Soma da coluna amount quando type = saída.",
      relevance: "Dá a visão completa das saídas do Ministério, mas nem toda essa despesa pertence ao curso e, por isso, não deve ser toda abatida da receita das aulas.",
      predicate: (record) => record.type.toLocaleLowerCase("pt-BR") === "saida",
    },
    "ministry-classes": {
      title: "Despesas do núcleo de aulas",
      valueLabel: "Valor da despesa",
      formula: "Soma da coluna amount quando finance_kind = aulas e type = saída.",
      relevance: "É o custo diretamente ligado às aulas. Por isso, este é o valor correto para comparar com o Valor Creditado do curso.",
      predicate: (record) => record.financeKind.toLocaleLowerCase("pt-BR") === "aulas" && record.type.toLocaleLowerCase("pt-BR") === "saida",
    },
    "ministry-general": {
      title: "Despesas gerais do Ministério",
      valueLabel: "Valor da despesa",
      formula: "Soma da coluna amount quando finance_kind = geral e type = saída.",
      relevance: "Pertence a outras áreas do Ministério. Continua visível para controle, mas fica fora do resultado direto do curso.",
      predicate: (record) => record.financeKind.toLocaleLowerCase("pt-BR") === "geral" && record.type.toLocaleLowerCase("pt-BR") === "saida",
    },
    "ministry-entries": {
      title: "Entradas administrativas",
      valueLabel: "Valor da entrada",
      formula: "De janeiro a setembro: R$ 500,00 mensais de verba missionária para Alex + R$ 1.000,00 mensais do Ministério.",
      relevance: "Registra R$ 1.500,00 de entrada por mês e mantém essas verbas administrativas separadas dos créditos pagos pelos alunos do curso.",
      source: "Entradas administrativas recorrentes cadastradas",
      predicate: (record) => record.type.toLocaleLowerCase("pt-BR") === "entrada",
    },
    "ministry-interpretation": {
      title: "Pagamentos diretos de interpretação",
      valueLabel: "Total identificado como interpretação",
      formula: "Soma de amount nas saídas cujo texto contém interpretação ou intérprete, excluindo lançamentos mistos com alimentação ou oferta.",
      relevance: "Mostra quanto foi pago especificamente por interpretação sem atribuir ao intérprete valores de despesas misturadas que não foram discriminadas.",
      predicate: (record) => isInterpretationExpense(record) && !isMixedInterpretationExpense(record),
    },
    "ministry-interpretation-mixed": {
      title: "Lançamentos mistos que incluem interpretação",
      valueLabel: "Valor total dos lançamentos mistos",
      formula: "Saídas que mencionam interpretação junto com alimentação ou oferta. O valor integral é mostrado, mas não é classificado como pagamento exclusivo de intérprete.",
      relevance: "Evita afirmar que toda a despesa foi paga ao intérprete quando o lançamento reúne finalidades diferentes e não apresenta a divisão.",
      predicate: (record) => isMixedInterpretationExpense(record),
    },
  };
  const definition = definitions[key];
  if (!definition) return null;
  const records = state.ministryRecords.filter(definition.predicate);
  const total = records.reduce((sum, record) => sum + record.amount, 0);
  const groupMap = new Map();
  records.forEach((record) => {
    if (!groupMap.has(record.category)) groupMap.set(record.category, { label: record.category, total: 0, lines: [] });
    const group = groupMap.get(record.category);
    const destination = record.name || record.eventTitle || "Não informado";
    group.total += record.amount;
    group.lines.push({
      label: destination,
      context: `${formatIsoDate(record.date)} · núcleo ${record.financeKind}`,
      value: record.amount,
    });
  });
  const groups = [...groupMap.values()].sort((a, b) => b.total - a.total || a.label.localeCompare(b.label, "pt-BR"));

  return {
    title: definition.title,
    source: definition.source || "finançassinaisdoreino.json",
    formula: definition.formula,
    relevance: definition.relevance,
    valueLabel: definition.valueLabel,
    total,
    calculation: [{ label: `${number.format(records.length)} lançamentos selecionados`, value: total }],
    resultLabel: definition.title,
    groupTitle: "Composição por categoria",
    groupHint: "Os subtotais abaixo formam o valor do card. Abra uma categoria para conferir os lançamentos.",
    groups,
    columns: ["Data", "Núcleo", "Categoria", "Nome ou destino", definition.valueLabel],
    rows: records.map((record) => {
      const destination = record.name || record.eventTitle || "Não informado";
      return [
        safe(formatIsoDate(record.date)),
        safe(record.financeKind),
        `<strong>${safe(record.category)}</strong>`,
        safe(destination),
        `<strong>${currency.format(record.amount)}</strong>`,
      ];
    }),
    note: `${number.format(records.length)} ${records.length === 1 ? "lançamento encontrado" : "lançamentos encontrados"} na base administrativa.`,
  };
}

function comparisonAuditBreakdown(key) {
  const summary = summarize(state.records);
  const ministry = summarizeMinistry(state.ministryRecords);
  if (key === "registered-net") {
    const total = summary.credited + ministry.entries - ministry.exits;
    return {
      title: "Resultado registrado no período",
      source: `${state.sourceName} + finançassinaisdoreino.json + entradas administrativas cadastradas`,
      formula: "Valor Creditado do curso + entradas administrativas − todas as saídas registradas do Ministério.",
      relevance: `Resume o fluxo presente nas bases sem confundi-lo com saldo bancário. Não inclui ${currency.format(summary.awaitingCredit)} aguardando a operadora nem ${currency.format(summary.overdue)} de inadimplência, porque esses valores ainda não foram creditados.`,
      valueLabel: "Resultado dos registros",
      total,
      calculation: [
        { label: "Crédito líquido do curso", value: summary.credited },
        { operator: "+", label: "Entradas administrativas", value: ministry.entries },
        { operator: "−", label: "Todas as saídas", value: ministry.exits },
      ],
      resultLabel: "Resultado registrado",
      groupTitle: "Composição por origem",
      groupHint: `Os três componentes abaixo formam o resultado. Os ${currency.format(summary.awaitingCredit)} da operadora e os ${currency.format(summary.overdue)} dos inadimplentes permanecem fora desta conta.`,
      groups: [
        { label: "Crédito líquido do curso", total: summary.credited, lines: [{ label: "Valor Creditado", context: state.sourceName, value: summary.credited }] },
        { label: "Entradas administrativas", total: ministry.entries, lines: [{ label: "R$ 1.500 por mês", context: "Janeiro a setembro de 2026", value: ministry.entries }] },
        { label: "Todas as saídas (subtração)", total: -ministry.exits, lines: [{ label: "Aulas + demais áreas", context: "Efeito negativo no resultado", value: -ministry.exits }] },
      ],
      columns: ["Componente", "Origem", "Efeito no resultado"],
      rows: [
        ["Crédito líquido do curso", state.sourceName, summary.credited],
        ["Entradas administrativas", "Cadastros recorrentes", ministry.entries],
        ["Todas as saídas", "finançassinaisdoreino.json", -ministry.exits],
      ],
      note: `Resultado gerencial dos registros disponíveis; não representa saldo bancário conciliado. Valores fora do resultado: ${currency.format(summary.awaitingCredit)} aguardando a operadora e ${currency.format(summary.overdue)} de inadimplência.`,
    };
  }
  if (key === "combined-entries") {
    const courseRows = state.records
      .filter((record) => record.credited > 0)
      .map((record) => ["Curso de Libras", record.creditedAt || "Sem data", record.donor, "Valor Creditado", record.credited]);
    const ministryRows = state.ministryRecords
      .filter((record) => record.type.toLocaleLowerCase("pt-BR") === "entrada")
      .map((record) => ["Finanças do Ministério", formatIsoDate(record.date), record.name || record.category, "Entrada administrativa", record.amount]);
    const rows = [...courseRows, ...ministryRows];
    const groups = [
      {
        label: "Créditos do Curso de Libras",
        total: summary.credited,
        lines: courseRows.map((row) => ({ label: row[2], context: `Creditado em ${row[1]}`, value: row[4] })),
      },
      {
        label: "Entradas administrativas",
        total: ministry.entries,
        lines: ministryRows.map((row) => ({ label: row[2], context: row[1], value: row[4] })),
      },
    ].filter((group) => group.lines.length || group.total > 0);
    return {
      title: "Todas as entradas registradas",
      source: `${state.sourceName} + entradas administrativas recorrentes cadastradas`,
      formula: "Valor Creditado do curso + registros administrativos cujo type = entrada.",
      relevance: "Mostra todas as entradas presentes nas bases sem afirmar que o total corresponde ao saldo bancário conciliado.",
      valueLabel: "Valor da entrada",
      total: summary.credited + ministry.entries,
      calculation: [
        { label: "Créditos líquidos do curso", value: summary.credited },
        { operator: "+", label: "Entradas administrativas", value: ministry.entries },
      ],
      resultLabel: "Entradas registradas",
      groupTitle: "Composição por base",
      groupHint: "Abra uma base para conferir os lançamentos que formam o subtotal.",
      groups,
      columns: ["Base", "Data", "Pessoa ou descrição", "Critério", "Valor da entrada"],
      rows: rows.map((row) => [safe(row[0]), safe(row[1]), `<strong>${safe(row[2])}</strong>`, safe(row[3]), `<strong>${currency.format(row[4])}</strong>`]),
      note: `${number.format(rows.length)} linhas formam a entrada comparada.`,
    };
  }
  if (key === "comparison-difference") {
    const result = summary.credited - ministry.classes;
    return {
      title: "Resultado direto do Curso de Libras",
      source: `${state.sourceName} + finançassinaisdoreino.json`,
      formula: "Valor Creditado do curso − despesas onde finance_kind = aulas e type = saída. Despesas gerais não entram nesta comparação.",
      relevance: "Mostra se a receita líquida do curso cobre os custos diretamente ligados às aulas. As demais despesas pertencem a outras áreas e permanecem fora desta conta.",
      valueLabel: "Resultado direto do curso",
      total: result,
      calculation: [
        { label: "Receita líquida do curso", value: summary.credited },
        { operator: "−", label: "Despesas das aulas", value: ministry.classes },
      ],
      resultLabel: "Resultado direto do curso",
      groupTitle: "Valores usados na comparação",
      groupHint: "Somente a receita do curso e os custos das aulas formam este resultado.",
      groups: [
        { label: "Receita líquida do curso", total: summary.credited, lines: [{ label: "Valor Creditado", context: state.sourceName, value: summary.credited }] },
        { label: "Despesas das aulas (subtração)", total: -ministry.classes, lines: [{ label: "Núcleo aulas", context: "finance_kind = aulas · type = saída", value: -ministry.classes }] },
      ],
      columns: ["Componente", "Base de origem", "Operação", "Efeito na diferença"],
      rows: [
        ["Receita líquida do curso", safe(state.sourceName), "Somar", `<strong>${currency.format(summary.credited)}</strong>`],
        ["Despesas das aulas", "finançassinaisdoreino.json", "Subtrair", `<strong>${currency.format(-ministry.classes)}</strong>`],
      ],
      note: `As despesas gerais de ${currency.format(ministry.general)} pertencem a outras áreas e foram mantidas fora desta comparação.`,
    };
  }
  return null;
}

function flowAuditBreakdown(key) {
  const match = key.match(/^flow-(general|course|pending|expense):(\d{4})-(\d{2})$/);
  if (!match) return null;
  const [, kind, yearText, monthText] = match;
  const year = Number(yearText);
  const month = Number(monthText);
  const periodLabel = new Intl.DateTimeFormat("pt-BR", { month: "long", year: "numeric" }).format(new Date(year, month - 1, 1, 12));
  const courseInMonth = (record, field) => {
    const date = parseDate(record[field]);
    return Boolean(date && date.getFullYear() === year && date.getMonth() + 1 === month);
  };
  const ministryInMonth = (record) => {
    const date = record.date ? new Date(`${record.date}T12:00:00`) : null;
    return Boolean(date && !Number.isNaN(date.getTime()) && date.getFullYear() === year && date.getMonth() + 1 === month);
  };
  const makeGroup = (label, records, value, context, sign = 1) => ({
    label,
    total: records.reduce((total, record) => total + value(record) * sign, 0),
    lines: records.map((record) => ({
      label: record.donor || record.name || record.eventTitle || record.category,
      context: context(record),
      value: value(record) * sign,
    })),
  });
  const credited = state.records.filter((record) => record.credited > 0 && courseInMonth(record, "creditedAt"));
  const administrative = state.ministryRecords.filter((record) => ministryInMonth(record) && record.type.toLocaleLowerCase("pt-BR") === "entrada");
  const exits = state.ministryRecords.filter((record) => ministryInMonth(record) && record.type.toLocaleLowerCase("pt-BR") === "saida");
  const classExits = exits.filter((record) => record.financeKind.toLocaleLowerCase("pt-BR") === "aulas");
  const generalExits = exits.filter((record) => record.financeKind.toLocaleLowerCase("pt-BR") !== "aulas");
  const courseTotal = credited.reduce((total, record) => total + record.credited, 0);
  const administrativeTotal = administrative.reduce((total, record) => total + record.amount, 0);
  const exitTotal = exits.reduce((total, record) => total + record.amount, 0);
  const classTotal = classExits.reduce((total, record) => total + record.amount, 0);

  if (kind === "general") {
    const groups = [
      makeGroup("Crédito líquido do curso", credited, (record) => record.credited, (record) => `Creditado em ${record.creditedAt}`),
      makeGroup("Entradas administrativas", administrative, (record) => record.amount, (record) => `${formatIsoDate(record.date)} · ${record.category}`),
      makeGroup("Saídas do Ministério", exits, (record) => record.amount, (record) => `${formatIsoDate(record.date)} · ${record.category}`, -1),
    ].filter((group) => group.lines.length);
    return {
      title: `Fluxo geral de ${periodLabel}`,
      source: `${state.sourceName} + finançassinaisdoreino.json + entradas cadastradas`,
      formula: "Valor Creditado do curso + entradas administrativas − todas as saídas do mês.",
      relevance: "Explica por que o mês terminou positivo ou negativo nos registros e quais lançamentos provocaram o resultado.",
      valueLabel: "Resultado registrado no mês",
      total: courseTotal + administrativeTotal - exitTotal,
      calculation: [{ label: "Curso", value: courseTotal }, { operator: "+", label: "Entradas administrativas", value: administrativeTotal }, { operator: "−", label: "Saídas", value: exitTotal }],
      resultLabel: "Resultado mensal",
      groupTitle: "Composição do mês",
      groupHint: "Entradas aparecem positivas; saídas aparecem negativas.",
      groups,
      rows: groups.flatMap((group) => group.lines),
      note: "Resultado mensal dos registros disponíveis; não representa saldo bancário conciliado.",
    };
  }

  if (kind === "course") {
    const groups = [
      makeGroup("Crédito líquido do curso", credited, (record) => record.credited, (record) => `Creditado em ${record.creditedAt}`),
      makeGroup("Despesas das aulas", classExits, (record) => record.amount, (record) => `${formatIsoDate(record.date)} · ${record.category}`, -1),
    ].filter((group) => group.lines.length);
    return {
      title: `Resultado do curso em ${periodLabel}`,
      source: `${state.sourceName} + finançassinaisdoreino.json`,
      formula: "Valor Creditado do curso − saídas do núcleo de aulas no mês.",
      relevance: "Mostra se a receita líquida daquele mês cobriu somente os custos diretamente associados às aulas.",
      valueLabel: "Resultado direto do curso",
      total: courseTotal - classTotal,
      calculation: [{ label: "Crédito do curso", value: courseTotal }, { operator: "−", label: "Despesas das aulas", value: classTotal }],
      resultLabel: "Resultado direto",
      groupTitle: "Receita e custo direto",
      groupHint: "Despesas de outras áreas não participam desta conta.",
      groups,
      rows: groups.flatMap((group) => group.lines),
      note: "Comparação gerencial do curso no mês selecionado.",
    };
  }

  if (kind === "pending") {
    const operator = state.records.filter((record) => record.paid > 0 && record.credited === 0 && courseInMonth(record, "expectedCreditAt"));
    const overdue = state.records.filter((record) => {
      const dueDate = parseDate(record.dueDate);
      return courseInMonth(record, "dueDate") && dueDate < new Date() && record.receivable > record.paid;
    });
    const future = state.records.filter((record) => {
      const dueDate = parseDate(record.dueDate);
      return courseInMonth(record, "dueDate") && dueDate >= new Date() && record.receivable > record.paid;
    });
    const groups = [
      makeGroup("Aguardando operadora", operator, (record) => record.paid, (record) => `${record.donor} · crédito previsto em ${record.expectedCreditAt}`),
      makeGroup("Vencido com o aluno", overdue, (record) => Math.max(record.receivable - record.paid, 0), (record) => `${record.donor} · parcela ${record.installment}`),
      makeGroup("Pagamento futuro", future, (record) => Math.max(record.receivable - record.paid, 0), (record) => `${record.donor} · vencimento ${record.dueDate}`),
    ].filter((group) => group.lines.length);
    const total = groups.reduce((sum, group) => sum + group.total, 0);
    return {
      title: `Valores pendentes de ${periodLabel}`,
      source: state.sourceName,
      formula: "Classificação por origem: repasse da operadora, parcela vencida com o aluno ou vencimento futuro.",
      relevance: "Evita tratar como inadimplência um pagamento que já saiu da conta do aluno e está apenas aguardando a operadora.",
      valueLabel: "Valores pendentes classificados",
      total,
      calculation: groups.map((group, index) => ({ operator: index ? "+" : undefined, label: group.label, value: group.total })),
      resultLabel: "Pendências do mês",
      groupTitle: "Origem das pendências",
      groupHint: "As categorias são independentes e mostram quem mantém o dinheiro pendente.",
      groups,
      rows: groups.flatMap((group) => group.lines),
      note: "O total serve para classificação operacional; não deve ser somado ao caixa já recebido.",
    };
  }

  const groups = [
    makeGroup("Despesas das aulas", classExits, (record) => record.amount, (record) => `${record.category} · ${record.name || record.eventTitle || "destino não informado"}`),
    makeGroup("Despesas de outras áreas", generalExits, (record) => record.amount, (record) => `${record.category} · ${record.name || record.eventTitle || "destino não informado"}`),
  ].filter((group) => group.lines.length);
  return {
    title: `Despesas de ${periodLabel}`,
    source: "finançassinaisdoreino.json",
    formula: "Soma de todas as saídas do mês, separadas entre núcleo de aulas e demais áreas.",
    relevance: "Mostra qual núcleo provocou o gasto do mês e permite abrir cada lançamento para conferir destino e categoria.",
    valueLabel: "Saídas registradas no mês",
    total: exitTotal,
    calculation: groups.map((group, index) => ({ operator: index ? "+" : undefined, label: group.label, value: group.total })),
    resultLabel: "Despesas do mês",
    groupTitle: "Composição por núcleo",
    groupHint: "O custo das aulas já está incluído no total e não deve ser somado novamente.",
    groups,
    rows: groups.flatMap((group) => group.lines),
    note: "Detalhamento mensal das saídas administrativas registradas.",
  };
}

function planningAuditBreakdown(key) {
  const plan = planningTotals();
  const fixedGroups = yearEndPlan.fixedExpenses.map((expense) => ({
    label: expense.label,
    total: expense.monthly * plan.months,
    lines: yearEndPlan.months.map((month) => ({ label: month, context: expense.reason, value: expense.monthly })),
  }));
  const courseGroup = {
    label: "Despesas planejadas do curso",
    total: yearEndPlan.courseExpenses,
    lines: [{ label: "Curso e Dia do Surdo", context: "Valor de planejamento informado", value: yearEndPlan.courseExpenses }],
  };
  const generalGroup = {
    label: "Despesas gerais fixas",
    total: plan.fixedGeneral,
    lines: yearEndPlan.fixedExpenses.map((expense) => ({ label: expense.label, context: `${currency.format(expense.monthly)} por mês × ${plan.months} meses`, value: expense.monthly * plan.months })),
  };
  const incomeGroup = (label, value, context) => ({ label, total: value, lines: [{ label, context, value }] });
  const expenseGroup = { label: "Despesas planejadas (subtração)", total: -plan.totalExpenses, lines: [{ label: "Curso + despesas gerais", context: "Efeito negativo no cenário", value: -plan.totalExpenses }] };
  const configs = {
    "plan-fixed-general": {
      title: "Plano de despesas gerais fixas até dezembro",
      total: plan.fixedGeneral,
      formula: "R$ 850,00 por mês × 4 meses: café e comida + Uber essencial + ajuda Éder.",
      relevance: "Mostra o custo mínimo recorrente que precisa ser reservado de setembro a dezembro, mesmo sem criar novas atividades.",
      calculation: yearEndPlan.fixedExpenses.map((expense, index) => ({ operator: index ? "+" : undefined, label: `${expense.label} · 4 meses`, value: expense.monthly * plan.months })),
      groups: fixedGroups,
      resultLabel: "Despesas gerais fixas",
    },
    "plan-year-expenses": {
      title: "Resumo de despesas planejadas até o fim do ano",
      total: plan.totalExpenses,
      formula: "R$ 4.000,00 do curso + R$ 3.400,00 de despesas gerais fixas.",
      relevance: "Define o valor total que o Ministério precisa cobrir até dezembro antes de considerar entradas certas ou possíveis recuperações.",
      calculation: [{ label: "Despesas do curso", value: yearEndPlan.courseExpenses }, { operator: "+", label: "Despesas gerais", value: plan.fixedGeneral }],
      groups: [courseGroup, generalGroup],
      resultLabel: "Total planejado",
    },
    "plan-certain-recovery": {
      title: "Crédito certo a receber do curso",
      total: yearEndPlan.certainCourseCredit,
      formula: "Valor de entrada concreta informado no planejamento: R$ 2.200,00.",
      relevance: "É a única recuperação tratada como certa no cenário real e reduz diretamente o déficit planejado.",
      calculation: [{ label: "Crédito certo informado", value: yearEndPlan.certainCourseCredit }],
      groups: [incomeGroup("Crédito certo do curso", yearEndPlan.certainCourseCredit, "Entrada concreta informada")],
      resultLabel: "Entrada certa",
    },
    "plan-uncertain-recovery": {
      title: "Inadimplência com recuperação incerta",
      total: yearEndPlan.uncertainDelinquency,
      formula: "Valor informado como inadimplência que talvez seja recuperada: R$ 1.600,00.",
      relevance: "Permanece fora do cenário real porque ainda não há garantia de recebimento; aparece somente no cenário potencial.",
      calculation: [{ label: "Possível recuperação", value: yearEndPlan.uncertainDelinquency }],
      groups: [incomeGroup("Inadimplência do curso", yearEndPlan.uncertainDelinquency, "Recebimento não garantido")],
      resultLabel: "Recuperação incerta",
    },
    "plan-alex": {
      title: "Possível entrada mensal explicada por Alex",
      total: plan.alexContribution,
      formula: "R$ 500,00 por mês × 4 meses = R$ 2.000,00.",
      relevance: "Se confirmada, essa entrada reduz o déficit de R$ 5.200,00 para R$ 3.200,00.",
      calculation: yearEndPlan.months.map((month, index) => ({ operator: index ? "+" : undefined, label: month, value: yearEndPlan.alexMonthlyContribution })),
      groups: [{ label: "Possível entrada de Alex", total: plan.alexContribution, lines: yearEndPlan.months.map((month) => ({ label: month, context: "R$ 500,00 mensais a confirmar", value: yearEndPlan.alexMonthlyContribution })) }],
      resultLabel: "Possível entrada",
    },
    "plan-deficit-base": {
      title: "Saldo negativo do cenário real",
      total: plan.certainDeficit,
      formula: "R$ 2.200,00 de entrada concreta − R$ 7.400,00 de despesas planejadas.",
      relevance: "Mostra o déficit que permanece mesmo considerando todo o crédito certo informado para o curso.",
      calculation: [{ label: "Entrada concreta", value: yearEndPlan.certainCourseCredit }, { operator: "−", label: "Despesas planejadas", value: plan.totalExpenses }],
      groups: [incomeGroup("Entrada concreta", yearEndPlan.certainCourseCredit, "Crédito certo do curso"), expenseGroup],
      resultLabel: "Saldo real projetado",
    },
    "plan-deficit-alex": {
      title: "Saldo projetado com a possível entrada de Alex",
      total: plan.deficitWithAlex,
      formula: "R$ 2.200,00 de crédito certo + R$ 2.000,00 de possível entrada de Alex − R$ 7.400,00 de despesas.",
      relevance: "Mesmo com a entrada adicional de Alex, ainda restariam R$ 3.200,00 negativos até dezembro.",
      calculation: [{ label: "Crédito certo", value: yearEndPlan.certainCourseCredit }, { operator: "+", label: "Possível entrada de Alex", value: plan.alexContribution }, { operator: "−", label: "Despesas", value: plan.totalExpenses }],
      groups: [incomeGroup("Crédito certo", yearEndPlan.certainCourseCredit, "Entrada concreta"), incomeGroup("Possível entrada de Alex", plan.alexContribution, "R$ 500,00 × 4 meses"), expenseGroup],
      resultLabel: "Saldo com Alex",
    },
    "plan-deficit-potential": {
      title: "Cenário potencial com recuperação da inadimplência",
      total: plan.potentialDeficit,
      formula: "R$ 2.200,00 certos + R$ 2.000,00 de Alex + R$ 1.600,00 incertos − R$ 7.400,00 de despesas.",
      relevance: "Este é o melhor cenário apresentado, mas ainda termina negativo e depende de duas entradas que não estão garantidas.",
      calculation: [{ label: "Crédito certo", value: yearEndPlan.certainCourseCredit }, { operator: "+", label: "Possível entrada de Alex", value: plan.alexContribution }, { operator: "+", label: "Inadimplência recuperada", value: yearEndPlan.uncertainDelinquency }, { operator: "−", label: "Despesas", value: plan.totalExpenses }],
      groups: [incomeGroup("Crédito certo", yearEndPlan.certainCourseCredit, "Entrada concreta"), incomeGroup("Possível entrada de Alex", plan.alexContribution, "Ainda não confirmada"), incomeGroup("Possível recuperação da inadimplência", yearEndPlan.uncertainDelinquency, "Recebimento incerto"), expenseGroup],
      resultLabel: "Saldo potencial",
    },
  };
  const config = configs[key];
  if (!config) return null;
  return {
    title: config.title,
    source: "Planejamento informado pelo Ministério",
    formula: config.formula,
    relevance: config.relevance,
    valueLabel: config.resultLabel,
    total: config.total,
    calculation: config.calculation,
    resultLabel: config.resultLabel,
    groupTitle: "Composição do planejamento",
    groupHint: "Estes valores são premissas de planejamento e permanecem separados dos lançamentos realizados.",
    groups: config.groups,
    columns: ["Componente", "Origem", "Valor"],
    rows: config.groups.map((group) => [group.label, "Planejamento informado", group.total]),
    note: "Projeção gerencial até dezembro; não representa lançamento já realizado no JSON ou crédito já recebido na conta.",
  };
}

function getAuditBreakdown(key) {
  if (key.startsWith("course-") || key.startsWith("defaulter:")) return courseAuditBreakdown(key);
  if (key.startsWith("ministry-")) return ministryAuditBreakdown(key);
  if (key.startsWith("plan-")) return planningAuditBreakdown(key);
  if (key.startsWith("flow-")) return flowAuditBreakdown(key);
  return comparisonAuditBreakdown(key);
}

function renderAuditGroups(breakdown) {
  if (!breakdown.groups?.length) {
    return `<div class="audit-groups-empty"><strong>Nenhum lançamento encontrado</strong><p>A regra deste card não encontrou valores na base atual. Por isso, o total é ${currency.format(0)}.</p></div>`;
  }
  return breakdown.groups.map((group, index) => {
    const personEquation = group.calculation?.length ? `
      <span class="audit-person-equation" aria-label="Composição do valor líquido de ${safe(group.label)}">
        ${group.calculation.map((term, termIndex) => `
          ${termIndex ? `<i aria-hidden="true">${safe(term.operator || "+")}</i>` : ""}
          <span><small>${safe(term.label)}</small><strong>${currency.format(term.value)}</strong></span>`).join("")}
        <i aria-hidden="true">=</i>
        <span class="audit-person-equation-result"><small>Valor líquido que entrou</small><strong>${currency.format(group.total)}</strong></span>
      </span>` : "";
    return `
    <details class="audit-composition-group">
      <summary class="${personEquation ? "audit-group-summary-with-equation" : ""}">
        <span class="audit-group-index">${String(index + 1).padStart(2, "0")}</span>
        <span class="audit-group-name"><strong>${safe(group.label)}</strong><small>${number.format(group.lines.length)} ${group.lines.length === 1 ? "lançamento" : "lançamentos"}</small></span>
        <span class="audit-group-total ${group.total < 0 ? "audit-negative-value" : ""}"><small>${safe(group.totalLabel || "Subtotal")}</small><strong>${currency.format(group.total)}</strong></span>
        <span class="audit-group-chevron" aria-hidden="true">⌄</span>
        ${personEquation}
      </summary>
      <div class="audit-group-lines">
        ${group.lines.map((line) => `
          <div class="audit-group-line">
            <div><strong>${safe(line.label)}</strong><span>${safe(line.context)}</span></div>
            <strong class="${line.value < 0 ? "audit-negative-value" : ""}">${currency.format(line.value)}</strong>
          </div>`).join("")}
      </div>
    </details>`;
  }).join("");
}

function openAuditBreakdown(key) {
  const breakdown = getAuditBreakdown(key);
  if (!breakdown) return;
  const trigger = document.activeElement;
  const backdrop = document.createElement("div");
  backdrop.className = "drawer-backdrop audit-breakdown-backdrop";
  backdrop.innerHTML = `
    <aside class="drawer audit-breakdown-drawer" role="dialog" aria-modal="true" aria-labelledby="audit-breakdown-title">
      <div class="drawer-head audit-breakdown-head">
        <div><p class="eyebrow">Auditoria do valor</p><h2 id="audit-breakdown-title" tabindex="-1">${safe(breakdown.title)}</h2><p>Confira a origem do total e abra somente as pessoas ou categorias que quiser verificar.</p></div>
        <button class="close-button" type="button" aria-label="Fechar detalhamento">${icons.close}</button>
      </div>
      <div class="audit-breakdown-content">
        <div class="audit-total-hero ${breakdown.total < 0 ? "audit-total-negative" : ""}">
          <div><span>${safe(breakdown.valueLabel)}</span><strong>${currency.format(breakdown.total)}</strong></div>
          <div class="audit-total-facts"><span><b>${number.format(breakdown.factCount ?? breakdown.rows.length)}</b> linhas analisadas</span><span><b>${number.format(breakdown.groups?.length || 0)}</b> subtotais</span></div>
        </div>
        <div class="audit-explanation-row">
          <div class="audit-criterion">
            <div><span>Critério aplicado</span><p>${safe(breakdown.formula)}</p></div>
            <small>Fonte: <b>${safe(breakdown.source)}</b></small>
          </div>
          <div class="audit-relevance">
            <span class="audit-relevance-icon">${icons.spark}</span>
            <div><span>Por que este número é relevante</span><p>${safe(breakdown.relevance || "Este valor ajuda a conferir a origem e o efeito financeiro dos lançamentos selecionados.")}</p></div>
          </div>
        </div>
        <section class="audit-composition" aria-labelledby="audit-composition-title">
          <div class="audit-section-heading">
            <div><p class="eyebrow">De onde veio</p><h3 id="audit-composition-title">${safe(breakdown.groupTitle)}</h3><p>${safe(breakdown.groupHint)}</p></div>
          </div>
          <div class="audit-composition-list">${renderAuditGroups(breakdown)}</div>
        </section>
        <p class="audit-breakdown-note">${safe(breakdown.note)}</p>
      </div>
    </aside>`;
  const host = document.fullscreenElement || document.body;
  host.appendChild(backdrop);
  document.body.classList.add("audit-modal-open");
  const close = () => {
    document.removeEventListener("keydown", escapeHandler);
    document.body.classList.remove("audit-modal-open");
    backdrop.remove();
    trigger?.focus?.();
  };
  const escapeHandler = (event) => {
    if (event.key === "Escape") {
      close();
      return;
    }
    if (event.key !== "Tab") return;
    const focusable = [...backdrop.querySelectorAll('button, summary, [href], [tabindex]:not([tabindex="-1"])')]
      .filter((element) => !element.hasAttribute("disabled"));
    if (!focusable.length) return;
    const first = focusable[0];
    const last = focusable.at(-1);
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  };
  backdrop.querySelector(".close-button").addEventListener("click", close);
  backdrop.addEventListener("click", (event) => { if (event.target === backdrop) close(); });
  document.addEventListener("keydown", escapeHandler);
  backdrop.querySelector("#audit-breakdown-title").focus();
}

function openDrawer(recordId) {
  const record = state.records.find((item) => item.id === recordId);
  if (!record) return;
  const backdrop = document.createElement("div");
  backdrop.className = "drawer-backdrop";
  backdrop.innerHTML = `
    <aside class="drawer" role="dialog" aria-modal="true" aria-labelledby="detail-title">
      <div class="drawer-head">
        <div><h2 id="detail-title">${safe(record.donor)}</h2><p>${safe(record.event)}</p></div>
        <button class="close-button" type="button" aria-label="Fechar detalhes">${icons.close}</button>
      </div>
      <div class="detail-summary"><span>Creditado líquido na conta</span><strong>${currency.format(record.credited)}</strong></div>
      <div class="detail-grid">
        ${detailItem("Status", record.status)}
        ${detailItem("Parcela", record.installment)}
        ${detailItem("Meio de pagamento", record.method)}
        ${detailItem("Vencimento", record.dueDate || "—")}
        ${detailItem("Recebimento", record.receivedAt || "Não recebido")}
        ${detailItem("Crédito efetivo", record.creditedAt || "Aguardando")}
        ${detailItem("Crédito previsto", record.expectedCreditAt || "Não informado")}
        ${detailItem("Valor previsto", currency.format(record.receivable))}
        ${detailItem("Valor pago pelo aluno", currency.format(record.paid))}
        ${detailItem("Líquido creditado", currency.format(record.credited))}
        ${detailItem("Tarifa", currency.format(Math.abs(record.fee)))}
      </div>
      <p class="data-note">${record.sourceRow ? `Registro ${record.id}, correspondente à linha ${record.sourceRow} da planilha importada.` : "Registro manual adicionado ao painel financeiro."}</p>
    </aside>`;
  document.body.appendChild(backdrop);
  const close = () => backdrop.remove();
  backdrop.querySelector(".close-button").addEventListener("click", close);
  backdrop.addEventListener("click", (event) => { if (event.target === backdrop) close(); });
  const escapeHandler = (event) => {
    if (event.key === "Escape") {
      close();
      document.removeEventListener("keydown", escapeHandler);
    }
  };
  document.addEventListener("keydown", escapeHandler);
  backdrop.querySelector(".close-button").focus();
}

function openParticipantDrawer(name) {
  const person = groupByParticipant(state.records).find((item) => item.name === name);
  if (!person) return;
  const backdrop = document.createElement("div");
  backdrop.className = "drawer-backdrop";
  const installments = [...person.records].sort((a, b) => (parseDate(a.dueDate)?.getTime() || 0) - (parseDate(b.dueDate)?.getTime() || 0));
  const overdueRecords = installments.filter((record) => {
    const dueDate = parseDate(record.dueDate);
    return Boolean(dueDate && dueDate < new Date() && record.receivable > record.paid);
  });
  const missingInstallments = Math.max(person.totalInstallments - installments.length, 0);
  backdrop.innerHTML = `
    <aside class="drawer participant-drawer" role="dialog" aria-modal="true" aria-labelledby="person-detail-title">
      <div class="drawer-head">
        <div><p class="eyebrow">Auditoria do aluno</p><h2 id="person-detail-title">${safe(person.name)}</h2><p>${person.paidInstallments} de ${person.totalInstallments} parcelas pagas • ${person.creditedInstallments} creditadas</p></div>
        <button class="close-button" type="button" aria-label="Fechar resumo">${icons.close}</button>
      </div>
      <div class="person-audit-grid">
        <div class="person-audit-metric person-audit-primary"><span>Creditado líquido</span><strong>${currency.format(person.credited)}</strong><small>Recebido pelo Ministério</small></div>
        <div class="person-audit-metric"><span>Pago pelo aluno</span><strong>${currency.format(person.paid)}</strong><small>Valor informativo</small></div>
        <div class="person-audit-metric person-audit-waiting"><span>Ainda vai cair</span><strong>${currency.format(person.awaitingCredit)}</strong><small>Pago, sem crédito</small></div>
        <div class="person-audit-metric ${person.overdue ? "person-audit-danger" : ""}"><span>Valor devendo</span><strong>${currency.format(person.overdue)}</strong><small>Parcelas vencidas</small></div>
      </div>
      ${overdueRecords.length ? `<div class="overdue-alert"><span>${icons.clock}</span><div><strong>Inadimplente</strong><p>Não pagou ${overdueRecords.map((record) => safe(overdueRecordDescription(record))).join("; ")}.</p></div></div>` : `<div class="student-ok-note">${icons.check}<span>Nenhuma parcela vencida em aberto.</span></div>`}
      ${missingInstallments ? `<div class="missing-data-note">${icons.clock}<span><strong>Atenção:</strong> ${missingInstallments} ${missingInstallments === 1 ? "parcela esperada não consta" : "parcelas esperadas não constam"} na exportação atual.</span></div>` : ""}
      <h3 class="drawer-section-title">Parcelas</h3>
      <div class="installment-list">
        ${installments
          .map((record) => {
            return `<div class="installment-audit-item">
              <div class="installment-audit-head"><div><span>Parcela</span><strong>${safe(record.installment)}</strong></div><span class="status-pill ${statusClass(record.status)}">${record.status}</span></div>
              <div class="installment-audit-grid">
                ${auditField("Vencimento", record.dueDate || "—")}
                ${auditField("Data do pagamento", record.receivedAt || "Não pago")}
                ${auditField("Crédito previsto", record.expectedCreditAt || "Sem previsão")}
                ${auditField("Crédito efetivo", record.creditedAt || "Não creditado")}
                ${auditField("Valor a receber", currency.format(record.receivable))}
                ${auditField("Pago pelo aluno", currency.format(record.paid))}
                ${auditField("Creditado líquido", currency.format(record.credited), "audit-value-highlight")}
                ${auditField("Tarifa", currency.format(Math.abs(record.fee)))}
              </div>
            </div>`;
          })
          .join("")}
      </div>
    </aside>`;
  document.body.appendChild(backdrop);
  const close = () => backdrop.remove();
  backdrop.querySelector(".close-button").addEventListener("click", close);
  backdrop.addEventListener("click", (event) => { if (event.target === backdrop) close(); });
  const escapeHandler = (event) => {
    if (event.key === "Escape") {
      close();
      document.removeEventListener("keydown", escapeHandler);
    }
  };
  document.addEventListener("keydown", escapeHandler);
  backdrop.querySelector(".close-button").focus();
}

function auditField(label, value, className = "") {
  return `<div class="installment-audit-field ${className}"><span>${safe(label)}</span><strong>${safe(value)}</strong></div>`;
}

function detailItem(label, value) {
  return `<div class="detail-item"><span>${safe(label)}</span><strong>${safe(value)}</strong></div>`;
}

async function importWorkbook(event) {
  const file = event.target.files?.[0];
  if (!file) return;
  try {
    await readWorkbook(await file.arrayBuffer(), file.name);
    state.search = "";
    state.donor = "Todos";
    state.method = "Todos";
    state.status = "Todos";
    state.installment = "Todas";
    state.dateFrom = "";
    state.dateTo = "";
    state.participantSearch = "";
    state.participantName = "Todos";
    state.participantStatus = "Todos";
    state.participantPage = 1;
    state.ministrySearch = "";
    state.ministryKind = "Todos";
    state.ministryType = "Todos";
    state.ministryCategory = "Todas";
    state.ministryDateFrom = "";
    state.ministryDateTo = "";
    renderApp();
    showToast(`${file.name} importada com sucesso.`);
  } catch (error) {
    showToast(error.message || "Não foi possível importar a planilha.");
  }
}

function exportCsv() {
  const headers = ["Participante", "Pagamento", "Vencimento", "Data do pagamento", "Crédito previsto", "Crédito efetivo", "Parcela", "Valor a receber", "Valor pago", "Líquido creditado", "Tarifa", "Status"];
  const lines = state.filtered.map((record) => [
    record.donor,
    record.method,
    record.dueDate,
    record.receivedAt,
    record.expectedCreditAt,
    record.creditedAt,
    record.installment,
    record.receivable.toFixed(2).replace(".", ","),
    record.paid.toFixed(2).replace(".", ","),
    record.credited.toFixed(2).replace(".", ","),
    Math.abs(record.fee).toFixed(2).replace(".", ","),
    record.status,
  ]);
  const csv = [headers, ...lines]
    .map((row) => row.map((cell) => `"${String(cell).replaceAll('"', '""')}"`).join(";"))
    .join("\n");
  const blob = new Blob([`\ufeff${csv}`], { type: "text/csv;charset=utf-8" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "lancamentos-financeiros.csv";
  link.click();
  URL.revokeObjectURL(link.href);
  showToast(`${state.filtered.length} lançamentos exportados.`);
}

function showToast(message) {
  document.querySelector(".toast")?.remove();
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = message;
  document.body.appendChild(toast);
  window.setTimeout(() => toast.remove(), 2800);
}

async function init() {
  try {
    const [workbookResponse, ministryResponse] = await Promise.all([
      fetch(sourceWorkbook),
      fetch(ministryFinanceJson),
    ]);
    if (!workbookResponse.ok) throw new Error("Não foi possível carregar Libras.xlsx.");
    if (!ministryResponse.ok) throw new Error("Não foi possível carregar finançassinaisdoreino.json.");
    await readWorkbook(await workbookResponse.arrayBuffer(), "Libras.xlsx");
    readMinistryFinance(await ministryResponse.json());
    renderApp();
  } catch (error) {
    document.querySelector("#app").innerHTML = `<div class="error-screen"><div><h1>Não foi possível abrir o painel</h1><p>${safe(error.message)}</p></div></div>`;
  }
}

init();
