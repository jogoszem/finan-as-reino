import { readFirstSheetRows } from "./xlsx-reader.js";
import { excludedParticipants, manualRecords } from "./manual-records.js";
import { parseCsv } from "./csv-reader.js";
import "./style.css";

const sourceWorkbook = new URL("../Libras.xlsx", import.meta.url);
const ministryFinanceCsv = new URL("../finançassinaisdoreino.csv", import.meta.url);

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
  participants: "Valores por aluno",
  cashflow: "Fluxo mensal",
  ministry: "Finanças do Ministério",
  transactions: "Tabela financeira",
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

function readMinistryFinance(csvText) {
  const records = parseCsv(csvText)
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

  if (!records.length) throw new Error("Não encontrei lançamentos válidos em finançassinaisdoreino.csv.");
  state.ministryRecords = records;
}

function summarizeMinistry(records) {
  const totalByType = (type) => records
    .filter((record) => record.type.toLocaleLowerCase("pt-BR") === type)
    .reduce((total, record) => total + record.amount, 0);
  const totalByKind = (kind) => records
    .filter((record) => record.financeKind.toLocaleLowerCase("pt-BR") === kind)
    .reduce((total, record) => total + record.amount, 0);
  const entries = totalByType("entrada");
  const exits = totalByType("saida");
  return {
    count: records.length,
    entries,
    exits,
    balance: entries - exits,
    classes: totalByKind("aulas"),
    general: totalByKind("geral"),
  };
}

function formatIsoDate(value) {
  const [year, month, day] = String(value || "").split("-");
  return year && month && day ? `${day}/${month}/${year}` : "—";
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
    const date = record.date ? new Date(`${record.date}T12:00:00`) : null;
    if (!date || Number.isNaN(date.getTime())) return;
    const group = ensureGroup(date);
    const type = record.type.toLocaleLowerCase("pt-BR");
    if (type === "entrada") group.income += record.amount;
    if (type === "saida") group.outgoing += record.amount;
  });

  return [...groups.values()].sort((a, b) => a.date - b.date);
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
              <div class="bar-tooltip">Entradas líquidas<br><strong>${currency.format(item.income)}</strong></div>
            </div>
            <div class="bar-wrap" style="--height:${outgoingHeight}%">
              <div class="bar bar-outgoing" style="height:${outgoingHeight}%;opacity:${item.outgoing ? 1 : 0}"></div>
              <div class="bar-tooltip">Saídas do Ministério<br><strong>${currency.format(item.outgoing)}</strong></div>
            </div>
          </div>`;
        })
        .join("")}
    </div>
    <div class="chart-labels" style="--columns:${groups.length}">
      ${groups.map((item) => `<span>${monthFormatter.format(item.date).replace(".", "")}</span>`).join("")}
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
  const colors = ["#307b67", "#e89555", "#d8e76a", "#5678a5"];
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

function buildPresentationPages() {
  const summary = summarize(state.records);
  const ministry = summarizeMinistry(state.ministryRecords);
  const timeline = buildCreditTimeline(state.records);
  const defaulters = buildDefaulters(state.records);
  const overdueInstallments = defaulters.reduce((total, person) => total + person.installments, 0);
  const toReconcile = timeline.reduce((total, item) => total + item.toReconcile, 0);
  const futureForecast = timeline.reduce((total, item) => total + item.forecast, 0);
  const comparisonDifference = summary.credited + ministry.entries - ministry.exits;

  return [
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
            <p class="presentation-source-label">Despesas administrativas</p>
            <h4>finançassinaisdoreino.csv</h4>
            <strong>${number.format(ministry.count)} lançamentos</strong>
            <p>Origina as saídas dos núcleos “aulas” e “geral”, com data, categoria e responsável.</p>
          </article>
        </div>
        <div class="presentation-source-note">${icons.check}<p><strong>Regra da auditoria:</strong> cada número apresentado informa sua coluna, fórmula e base de origem.</p></div>`,
    },
    {
      theme: "income",
      kicker: "Página 2 · Recebimentos",
      title: "Quanto realmente entrou na conta?",
      content: `
        <div class="presentation-main-metric presentation-main-income audit-clickable" ${auditAttributes("course-credited", "recebido líquido pelo Ministério")}>
          <span>Recebido líquido pelo Ministério</span>
          <strong>${currency.format(summary.credited)}</strong>
          <p>Fonte: <b>Libras.xlsx</b> · soma da coluna <b>Valor Creditado</b>.</p>
        </div>
        <div class="presentation-metric-grid">
          <article class="audit-clickable" ${auditAttributes("course-paid", "pago pelos alunos")}><span>Pago pelos alunos</span><strong>${currency.format(summary.paid)}</strong><p>Coluna Valor Pago. É informativo e não significa dinheiro em conta.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-fees", "tarifas e taxas")}><span>Tarifas e taxas</span><strong>${currency.format(summary.fees)}</strong><p>Valor absoluto da coluna Despesa Financeira.</p></article>
          <article class="audit-clickable" ${auditAttributes("course-awaiting", "pago sem crédito registrado")}><span>Pago sem crédito registrado</span><strong>${currency.format(summary.awaitingCredit)}</strong><p>Pagamentos feitos no cartão que serão creditados posteriormente pela operadora na conta do Ministério.</p></article>
        </div>
        <div class="presentation-formula"><span>Conferência dos dados atuais</span><strong>${currency.format(summary.paid)} − ${currency.format(summary.fees)} − ${currency.format(summary.awaitingCredit)} = ${currency.format(summary.credited)}</strong></div>`,
    },
    {
      theme: "credit",
      kicker: "Página 3 · Competência",
      title: "Quando o dinheiro entrou — ou deveria entrar?",
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
      title: "O que saiu nas finanças do Ministério?",
      content: `
        <div class="presentation-main-metric presentation-main-danger audit-clickable" ${auditAttributes("ministry-exits", "total de saídas registradas")}>
          <span>Total de saídas registradas</span>
          <strong>${currency.format(ministry.exits)}</strong>
          <p>Fonte: finançassinaisdoreino.csv · soma de amount onde type = “saida”.</p>
        </div>
        <div class="presentation-metric-grid">
          <article class="audit-clickable" ${auditAttributes("ministry-classes", "despesas do núcleo de aulas")}><span>Núcleo de aulas</span><strong>${currency.format(ministry.classes)}</strong><p>${state.ministryRecords.filter((record) => record.financeKind === "aulas").length} lançamentos com finance_kind = aulas.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-general", "despesas gerais")}><span>Despesas gerais</span><strong>${currency.format(ministry.general)}</strong><p>${state.ministryRecords.filter((record) => record.financeKind === "geral").length} lançamentos com finance_kind = geral.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-entries", "entradas administrativas")}><span>Entradas administrativas</span><strong>${currency.format(ministry.entries)}</strong><p>A base atual não possui registros com type = entrada.</p></article>
        </div>
        <div class="presentation-source-note presentation-source-warning">${icons.clock}<p><strong>Importante:</strong> essa base registra despesas administrativas e permanece separada das mensalidades do curso.</p></div>`,
    },
    {
      theme: "flow",
      kicker: "Página 6 · Fluxo mensal",
      title: "Comparativo do que entrou e do que saiu",
      content: `
        <div class="presentation-chart-card">
          <div class="legend"><span><i style="background:#307b67"></i>Entradas líquidas</span><span><i style="background:#b85543"></i>Saídas do Ministério</span></div>
          ${cashFlowComparisonChart(state.records, state.ministryRecords)}
        </div>
        <div class="presentation-metric-grid presentation-flow-summary">
          <article class="audit-clickable" ${auditAttributes("combined-entries", "entradas no recorte")}><span>Entradas no recorte</span><strong>${currency.format(summary.credited + ministry.entries)}</strong><p>Valor Creditado do curso + entradas administrativas.</p></article>
          <article class="audit-clickable" ${auditAttributes("ministry-exits", "saídas no recorte")}><span>Saídas no recorte</span><strong>${currency.format(ministry.exits)}</strong><p>Registros type = saída no CSV.</p></article>
          <article class="audit-clickable ${comparisonDifference < 0 ? "presentation-alert-metric" : ""}" ${auditAttributes("comparison-difference", "diferença comparada")}><span>Diferença comparada</span><strong>${currency.format(comparisonDifference)}</strong><p>Comparação entre as duas bases; não equivale ao saldo bancário completo.</p></article>
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
          <article><span>04</span><div><strong>Despesas do Ministério</strong><p>Vêm exclusivamente do finançassinaisdoreino.csv e são comparadas por data.</p></div></article>
        </div>
        <div class="presentation-audit-footnote">
          <strong>Ajustes aplicados à base</strong>
          <p>Luciana Caroline Correia da Silva foi excluída dos cálculos. Thaise Almeida e Yasmine Delefrati foram adicionadas conforme solicitado. Os dois lançamentos de Caio são mantidos porque correspondem a ele e à esposa.</p>
        </div>`,
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
        <article class="presentation-slide">
        <div class="presentation-slide-head">
          <div><p class="presentation-kicker">${page.kicker}</p><h3>${page.title}</h3></div>
          <span class="presentation-page-number">${String(state.presentationPage + 1).padStart(2, "0")}</span>
        </div>
        <div class="presentation-slide-content">${page.content}</div>
        </article>
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
      state.presentationPage += action === "next" ? 1 : -1;
      renderPresentation();
    });
  });
}

function renderApp() {
  const summary = summarize(state.records);
  const ministrySummary = summarizeMinistry(state.ministryRecords);
  const forecastTotal = buildCreditTimeline(state.records)
    .filter((item) => item.month === 9 || item.month === 10)
    .reduce((total, item) => total + item.forecast, 0);
  const app = document.querySelector("#app");

  app.innerHTML = `
    <div class="app-shell">
      <aside class="sidebar" aria-label="Navegação principal">
        <div class="brand">
          <div class="brand-mark">R</div>
          <div class="brand-copy"><strong>Reino</strong><span>Financeiro</span></div>
        </div>
        <p class="side-label">Painel</p>
        <nav class="side-nav">
          <button class="nav-item ${state.activeView === "presentation" ? "active" : ""}" data-target="presentation" aria-current="${state.activeView === "presentation" ? "page" : "false"}">${icons.spark}<span>Apresentação</span></button>
          <button class="nav-item ${state.activeView === "overview" ? "active" : ""}" data-target="overview" aria-current="${state.activeView === "overview" ? "page" : "false"}">${icons.overview}<span>Visão geral</span></button>
          <button class="nav-item ${state.activeView === "participants" ? "active" : ""}" data-target="participants" aria-current="${state.activeView === "participants" ? "page" : "false"}">${icons.people}<span>Valores por pessoa</span></button>
          <button class="nav-item ${state.activeView === "cashflow" ? "active" : ""}" data-target="cashflow" aria-current="${state.activeView === "cashflow" ? "page" : "false"}">${icons.chart}<span>Fluxo mensal</span></button>
          <button class="nav-item ${state.activeView === "ministry" ? "active" : ""}" data-target="ministry" aria-current="${state.activeView === "ministry" ? "page" : "false"}">${icons.money}<span>Finanças do Ministério</span></button>
          <button class="nav-item ${state.activeView === "transactions" ? "active" : ""}" data-target="transactions" aria-current="${state.activeView === "transactions" ? "page" : "false"}">${icons.rows}<span>Tabela</span></button>
        </nav>
        <div class="side-source">
          <div class="source-line"><i class="source-dot"></i>Fonte conectada</div>
          <p class="source-file">${safe(state.sourceName)}</p>
          <p class="source-file source-file-secondary">finançassinaisdoreino.csv</p>
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

        <section class="kpi-grid app-view" data-view="overview" aria-label="Indicadores financeiros">
          <article class="card kpi-card kpi-card-primary audit-clickable" ${auditAttributes("course-credited", "recebido líquido pelo Ministério")}>
            <div class="kpi-head"><span class="kpi-label">Recebido líquido pelo Ministério</span><span class="kpi-icon">${icons.bank}</span></div>
            <div class="kpi-value">${currency.format(summary.credited)}</div>
            <p class="kpi-note">Soma do Valor Creditado: dinheiro que realmente entrou na conta.</p>
          </article>
          <article class="card kpi-card audit-clickable" ${auditAttributes("course-awaiting", "pago aguardando crédito")}>
            <div class="kpi-head"><span class="kpi-label">Pago, aguardando crédito</span><span class="kpi-icon">${icons.clock}</span></div>
            <div class="kpi-value">${currency.format(summary.awaitingCredit)}</div>
            <p class="kpi-note">Pagamentos feitos no cartão que serão creditados posteriormente pela operadora na conta do Ministério.</p>
          </article>
          <article class="card kpi-card kpi-forecast audit-clickable" ${auditAttributes("course-future", "previsão de setembro e outubro")}>
            <div class="kpi-head"><span class="kpi-label">Previsão setembro + outubro</span><span class="kpi-icon">${icons.chart}</span></div>
            <div class="kpi-value">${currency.format(forecastTotal)}</div>
            <p class="kpi-note">Previsão bruta pelas datas informadas; o líquido pode sofrer tarifas.</p>
          </article>
          <article class="card kpi-card kpi-card-danger audit-clickable" ${auditAttributes("course-overdue", "inadimplência vencida")}>
            <div class="kpi-head"><span class="kpi-label">Inadimplência vencida</span><span class="kpi-icon">${icons.clock}</span></div>
            <div class="kpi-value">${currency.format(summary.overdue)}</div>
            <p class="kpi-note">Parcelas vencidas cujo Valor Pago ainda não cobre o previsto.</p>
          </article>
          <article class="card kpi-card audit-clickable" ${auditAttributes("course-receivable", "valor total cadastrado")}>
            <div class="kpi-head"><span class="kpi-label">Valor total cadastrado</span><span class="kpi-icon">${icons.money}</span></div>
            <div class="kpi-value">${currency.format(summary.receivable)}</div>
            <p class="kpi-note">Soma do Valor à Receber. Não representa dinheiro em conta.</p>
          </article>
          <article class="card kpi-card kpi-card-info audit-clickable" ${auditAttributes("course-paid", "pago pelos alunos")}>
            <div class="kpi-head"><span class="kpi-label">Pago pelos alunos</span><span class="kpi-icon">${icons.check}</span></div>
            <div class="kpi-value">${currency.format(summary.paid)}</div>
            <p class="kpi-note">Informativo: saiu da conta dos alunos, mas não significa que o Ministério recebeu.</p>
          </article>
          <article class="card kpi-card audit-clickable" ${auditAttributes("course-fees", "tarifas e taxas do cartão")}>
            <div class="kpi-head"><span class="kpi-label">Tarifas e taxas do cartão</span><span class="kpi-icon">${icons.money}</span></div>
            <div class="kpi-value">${currency.format(summary.fees)}</div>
            <p class="kpi-note">Soma das tarifas e taxas financeiras cobradas pela manutenção e pelo processamento dos pagamentos no cartão de crédito.</p>
          </article>
        </section>

        <section class="card ministry-overview-card app-view" data-view="overview" aria-label="Resumo das finanças do Ministério">
          <div class="ministry-overview-head">
            <div>
              <p class="eyebrow">Base finançassinaisdoreino</p>
              <h2 class="panel-title">Despesas registradas pelo Ministério</h2>
              <p class="panel-subtitle">Resumo separado da receita do Curso de Libras para não misturar bases com finalidades diferentes.</p>
            </div>
            <button class="button" id="open-ministry-view" type="button">Ver lançamentos ${icons.arrow}</button>
          </div>
          <div class="ministry-overview-metrics">
            <div class="ministry-overview-primary audit-clickable" ${auditAttributes("ministry-exits", "total de saídas")}><span>Total de saídas</span><strong>${currency.format(ministrySummary.exits)}</strong><small>${number.format(ministrySummary.count)} lançamentos</small></div>
            <div class="audit-clickable" ${auditAttributes("ministry-classes", "despesas do núcleo de aulas")}><span>Núcleo de aulas</span><strong>${currency.format(ministrySummary.classes)}</strong><small>Pagamentos e despesas de aulas</small></div>
            <div class="audit-clickable" ${auditAttributes("ministry-general", "despesas gerais")}><span>Despesas gerais</span><strong>${currency.format(ministrySummary.general)}</strong><small>Demais despesas do Ministério</small></div>
            <div class="audit-clickable" ${auditAttributes("ministry-all", "período da base administrativa")}><span>Período da base</span><strong>Jan — Ago/2026</strong><small>Todos os registros atuais são saídas via Pix</small></div>
          </div>
        </section>

        <section class="overview-tables-grid app-view" data-view="overview" aria-label="Créditos mensais e inadimplência">
          <article class="card overview-data-card">
            <div class="overview-data-head">
              <div>
                <p class="eyebrow">Crédito por competência</p>
                <h2 class="panel-title">Valores creditados por mês</h2>
                <p class="panel-subtitle">Cada mês separa o que entrou, o que venceu sem crédito e a previsão futura</p>
              </div>
              <div class="mini-legend"><span><i></i>Realizado</span><span><i></i>A conciliar</span><span><i></i>Previsão futura</span></div>
            </div>
            <div class="compact-table-wrap">
              <table class="compact-table credit-month-table">
                <thead><tr><th>Mês</th><th>Creditado líquido</th><th>Crédito vencido</th><th>Previsão futura</th></tr></thead>
                <tbody>${renderCreditTimeline(state.records)}</tbody>
              </table>
            </div>
            <p class="table-data-note">“Creditado líquido” é dinheiro em conta. “Crédito vencido” foi pago, mas a previsão terminou sem Valor Creditado. “Previsão futura” começa em setembro e poderá sofrer tarifas.</p>
          </article>

          <article class="card overview-data-card">
            <div class="overview-data-head">
              <div>
                <p class="eyebrow eyebrow-alert">Acompanhamento</p>
                <h2 class="panel-title">Inadimplentes</h2>
                <p class="panel-subtitle">Pessoas com parcelas vencidas e saldo ainda não pago</p>
              </div>
              <span class="defaulter-count">${buildDefaulters(state.records).length} pessoas</span>
            </div>
            <div class="defaulter-total-card audit-clickable" ${auditAttributes("course-overdue", "soma dos inadimplentes")}>
              <span class="defaulter-total-icon">${icons.clock}</span>
              <div><span>Soma dos inadimplentes</span><strong>${currency.format(sumDefaulters(state.records))}</strong><small>Saldo das parcelas vencidas não pagas</small></div>
            </div>
            <div class="compact-table-wrap">
              <table class="compact-table defaulter-table">
                <thead><tr><th>Participante</th><th>Tipo da inadimplência</th><th>Parcelas e vencimentos</th><th>Em aberto</th></tr></thead>
                <tbody>${renderDefaulters(state.records)}</tbody>
              </table>
            </div>
          </article>
        </section>

        <section class="flow-section app-view" id="cashflow" data-view="cashflow">
          <div class="flow-heading">
            <div><p class="eyebrow">Análise visual</p><h2 class="section-title">Fluxo mensal</h2><p class="section-copy">Quatro leituras financeiras, incluindo o comparativo mensal entre o dinheiro que entrou e o que saiu.</p></div>
            <span class="participant-total-badge">4 gráficos</span>
          </div>
          <div class="flow-grid">
            <article class="card panel flow-card-full">
              <div class="panel-head">
                <div><span class="chart-number">01</span><h3 class="panel-title">Entradas líquidas x saídas</h3><p class="panel-subtitle">Entradas usam o Valor Creditado dos alunos e eventuais entradas administrativas. Saídas usam os registros do finançassinaisdoreino.</p></div>
                <div class="legend"><span><i style="background:#307b67"></i>Entradas</span><span><i style="background:#b85543"></i>Saídas</span></div>
              </div>
              ${cashFlowComparisonChart(state.records, state.ministryRecords)}
            </article>
            <article class="card panel">
              <div class="panel-head"><div><span class="chart-number">02</span><h3 class="panel-title">Previsão futura de crédito</h3><p class="panel-subtitle">Somente parcelas previstas a partir de setembro; agosto não entra como previsão futura.</p></div><span class="chart-definition chart-definition-orange">Estimativa bruta</span></div>
              ${singleSeriesChart(forecastCreditSeries(state.records), "orange")}
            </article>
            <article class="card panel">
              <div class="panel-head"><div><span class="chart-number">03</span><h3 class="panel-title">Formas de pagamento</h3><p class="panel-subtitle">Distribuição do Valor Pago informado pelos alunos.</p></div><span class="chart-definition chart-definition-blue">Informativo</span></div>
              ${methodsChart(state.records)}
            </article>
            <article class="card panel flow-card-full">
              <div class="panel-head"><div><span class="chart-number">04</span><h3 class="panel-title">Inadimplência por vencimento</h3><p class="panel-subtitle">Saldo das parcelas vencidas que ainda não foram pagas.</p></div><span class="chart-definition chart-definition-red">Cobrança necessária</span></div>
              ${singleSeriesChart(delinquencySeries(state.records), "red")}
            </article>
          </div>
        </section>

        <section class="participants-section app-view" id="participants" data-view="participants">
          <div class="participant-heading">
            <div>
              <p class="eyebrow">Auditoria individual</p>
              <h2 class="section-title">Valores por aluno</h2>
              <p class="section-copy">Uma linha por pessoa. O valor creditado líquido é o dinheiro realmente recebido pelo Ministério.</p>
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
        </section>

        <section class="card ministry-section app-view" id="ministry" data-view="ministry">
          <div class="table-top ministry-table-top">
            <div>
              <p class="eyebrow">Base exclusiva</p>
              <h2 class="section-title">Finanças Sinais do Reino</h2>
              <p class="panel-subtitle">Lançamentos administrativos importados de finançassinaisdoreino.csv</p>
            </div>
            <button class="button" id="clear-ministry-filters" type="button">Limpar filtros</button>
          </div>
          <div class="ministry-source-note">
            <span>${icons.money}</span>
            <p><strong>Escopo desta aba:</strong> despesas gerais e de aulas do Ministério. Estes valores permanecem separados dos recebimentos dos alunos.</p>
          </div>
          <div class="advanced-filters ministry-filters">
            <label class="search-box" aria-label="Buscar nas finanças do Ministério">${icons.search}<input id="ministry-search" type="search" placeholder="Buscar categoria, pessoa ou descrição…" /></label>
            <label class="filter-field"><span>Núcleo</span><select class="select" id="ministry-kind-filter">
              <option value="">Todos os núcleos</option>
              ${[...new Set(state.ministryRecords.map((item) => item.financeKind))].sort().map((kind) => `<option value="${safe(kind)}">${safe(kind)}</option>`).join("")}
            </select></label>
            <label class="filter-field"><span>Tipo</span><select class="select" id="ministry-type-filter">
              <option value="">Todos os tipos</option>
              ${[...new Set(state.ministryRecords.map((item) => item.type))].sort().map((type) => `<option value="${safe(type)}">${safe(type)}</option>`).join("")}
            </select></label>
            <label class="filter-field"><span>Categoria</span><select class="select" id="ministry-category-filter">
              <option value="">Todas as categorias</option>
              ${[...new Set(state.ministryRecords.map((item) => item.category))].sort((a, b) => a.localeCompare(b, "pt-BR")).map((category) => `<option value="${safe(category)}">${safe(category)}</option>`).join("")}
            </select></label>
            <label class="filter-field date-field"><span>Data inicial</span><input class="date-input" id="ministry-date-from" type="date" /></label>
            <label class="filter-field date-field"><span>Data final</span><input class="date-input" id="ministry-date-to" type="date" /></label>
          </div>
          <div id="ministry-table-region"></div>
        </section>

        <section class="card table-card app-view" id="transactions" data-view="transactions">
          <div class="table-top">
            <div><p class="eyebrow">Base completa</p><h2 class="section-title">Tabela financeira</h2><p class="panel-subtitle">Todos os campos relevantes da planilha, uma linha por parcela</p></div>
            <button class="button" id="clear-filters" type="button">Limpar filtros</button>
          </div>
          <div class="quick-filter-bar" aria-label="Filtros rápidos">
            <button class="quick-filter active" data-quick-status="Todos" type="button">Todos</button>
            <button class="quick-filter quick-filter-danger" data-quick-status="Em atraso" type="button">Inadimplentes</button>
            <button class="quick-filter quick-filter-orange" data-quick-status="A creditar" type="button">Aguardando crédito</button>
            <button class="quick-filter quick-filter-green" data-quick-status="Creditado" type="button">Já creditados</button>
          </div>
          <div class="advanced-filters">
            <label class="search-box" aria-label="Buscar participante">${icons.search}<input id="search-input" type="search" placeholder="Buscar nome ou parcela…" /></label>
            <label class="filter-field"><span>Aluno</span><select class="select" id="donor-filter">
              <option>Todas as pessoas</option>
              ${[...new Set(state.records.map((item) => item.donor))].sort((a, b) => a.localeCompare(b, "pt-BR")).map((name) => `<option value="${safe(name)}">${safe(name)}</option>`).join("")}
            </select></label>
            <label class="filter-field"><span>Pagamento</span><select class="select" id="method-filter" aria-label="Filtrar por meio de pagamento">
              <option>Todos os meios</option>
              ${[...new Set(state.records.map((item) => item.method))].sort().map((method) => `<option>${safe(method)}</option>`).join("")}
            </select></label>
            <label class="filter-field"><span>Situação</span><select class="select" id="status-filter" aria-label="Filtrar por status">
              <option>Todos os status</option>
              <option>Creditado</option><option>A creditar</option><option>Em atraso</option><option>Pendente</option>
            </select></label>
            <label class="filter-field"><span>Parcela</span><select class="select" id="installment-filter">
              <option>Todas as parcelas</option>
              ${[...new Set(state.records.map((item) => item.installment))].sort((a, b) => a.localeCompare(b, "pt-BR", { numeric: true })).map((installment) => `<option>${safe(installment)}</option>`).join("")}
            </select></label>
            <label class="filter-field date-field"><span>Vencimento inicial</span><input class="date-input" id="date-from" type="date" /></label>
            <label class="filter-field date-field"><span>Vencimento final</span><input class="date-input" id="date-to" type="date" /></label>
          </div>
          <div id="table-region"></div>
        </section>
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
  const openAuditFromEvent = (event) => {
    const trigger = event.target.closest?.("[data-audit-key]");
    if (!trigger || !appRoot.contains(trigger)) return;
    if (event.type === "keydown" && !["Enter", " "].includes(event.key)) return;
    if (event.type === "keydown") event.preventDefault();
    openAuditBreakdown(trigger.dataset.auditKey);
  };
  appRoot.addEventListener("click", openAuditFromEvent);
  appRoot.addEventListener("keydown", openAuditFromEvent);

  document.querySelectorAll(".nav-item").forEach((button) => {
    button.addEventListener("click", () => activateView(button.dataset.target, true));
  });

  document.querySelector("#open-ministry-view").addEventListener("click", () => activateView("ministry", true));

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
  const totalPages = Math.max(Math.ceil(participants.length / state.participantPerPage), 1);
  state.participantPage = Math.min(state.participantPage, totalPages);
  const start = (state.participantPage - 1) * state.participantPerPage;
  const pageRows = participants.slice(start, start + state.participantPerPage);
  const filteredTotals = pageRows.length
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

  const pageButtons = getParticipantPageButtons(totalPages)
    .map((page) =>
      page === "…"
        ? `<span>…</span>`
        : `<button class="page-button ${page === state.participantPage ? "active" : ""}" data-participant-page="${page}" type="button">${page}</button>`,
    )
    .join("");

  region.innerHTML = `
    <div class="table-wrap people-table-wrap">
      <table class="people-table">
        <thead><tr><th>Aluno</th><th>Parcelas pagas</th><th>Pago pelo aluno</th><th>Creditado líquido</th><th>Ainda vai cair</th><th>Devendo</th><th>Situação</th><th></th></tr></thead>
        <tbody>${body}</tbody>
        ${pageRows.length ? `<tfoot><tr><td colspan="2">Total do filtro</td><td>${currency.format(filteredTotals.paid)}</td><td>${currency.format(filteredTotals.credited)}</td><td>${currency.format(filteredTotals.awaitingCredit)}</td><td>${currency.format(filteredTotals.overdue)}</td><td colspan="2"></td></tr></tfoot>` : ""}
      </table>
    </div>
    <footer class="table-footer">
      <span>Mostrando ${pageRows.length ? start + 1 : 0}–${Math.min(start + state.participantPerPage, participants.length)} de ${number.format(participants.length)} pessoas</span>
      <div class="pagination">
        <button class="page-button" data-participant-page="prev" type="button" ${state.participantPage === 1 ? "disabled" : ""} aria-label="Página anterior">‹</button>
        ${pageButtons}
        <button class="page-button" data-participant-page="next" type="button" ${state.participantPage === totalPages ? "disabled" : ""} aria-label="Próxima página">›</button>
      </div>
    </footer>`;

  region.querySelectorAll("[data-participant]").forEach((button) => {
    button.addEventListener("click", () => openParticipantDrawer(button.dataset.participant));
  });
  region.querySelectorAll("[data-participant-page]").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.participantPage === "prev") state.participantPage -= 1;
      else if (button.dataset.participantPage === "next") state.participantPage += 1;
      else state.participantPage = Number(button.dataset.participantPage);
      renderParticipants();
    });
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
  const totalPages = Math.max(Math.ceil(state.filtered.length / state.perPage), 1);
  const start = (state.page - 1) * state.perPage;
  const rows = state.filtered.slice(start, start + state.perPage);

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

  const pageButtons = getPageButtons(totalPages)
    .map((page) =>
      page === "…"
        ? `<span>…</span>`
        : `<button class="page-button ${page === state.page ? "active" : ""}" data-page="${page}" type="button">${page}</button>`,
    )
    .join("");

  region.innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Aluno</th><th>Pagamento</th><th>Vencimento</th><th>Data do pagamento</th><th>Crédito previsto</th><th>Crédito efetivo</th><th>Parcela</th><th>Valor a receber</th><th>Valor pago</th><th>Creditado líquido</th><th>Tarifa</th><th>Status</th><th></th></tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
    <footer class="table-footer">
      <span>Mostrando ${rows.length ? start + 1 : 0}–${Math.min(start + state.perPage, state.filtered.length)} de ${number.format(state.filtered.length)} lançamentos</span>
      <div class="pagination">
        <button class="page-button" data-page="prev" type="button" ${state.page === 1 ? "disabled" : ""} aria-label="Página anterior">‹</button>
        ${pageButtons}
        <button class="page-button" data-page="next" type="button" ${state.page === totalPages ? "disabled" : ""} aria-label="Próxima página">›</button>
      </div>
    </footer>`;

  region.querySelectorAll("[data-record-id]").forEach((button) => {
    button.addEventListener("click", () => openDrawer(button.dataset.recordId));
  });
  region.querySelectorAll("[data-page]").forEach((button) => {
    button.addEventListener("click", () => {
      if (button.dataset.page === "prev") state.page -= 1;
      else if (button.dataset.page === "next") state.page += 1;
      else state.page = Number(button.dataset.page);
      renderTable();
    });
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
      predicate: () => true,
      value: (record) => record.receivable,
    },
    "course-credited": {
      title: "Recebido líquido pelo Ministério",
      valueLabel: "Valor Creditado",
      formula: "Soma da coluna Valor Creditado nas linhas em que o crédito líquido é maior que zero.",
      predicate: (record) => record.credited > 0,
      value: (record) => record.credited,
    },
    "course-awaiting": {
      title: "Pago, aguardando crédito",
      valueLabel: "Valor Pago sem crédito",
      formula: "Soma do Valor Pago quando Valor Pago > 0 e Valor Creditado = 0. São pagamentos no cartão que serão creditados posteriormente pela operadora.",
      predicate: (record) => record.paid > 0 && record.credited === 0,
      value: (record) => record.paid,
    },
    "course-future": {
      title: "Previsão de crédito de setembro e outubro",
      valueLabel: "Valor previsto",
      formula: "Soma do Valor Pago — ou do Valor a Receber quando ainda não houve pagamento — com crédito previsto para setembro ou outubro e Valor Creditado = 0.",
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
      predicate: () => true,
      value: (record) => record.receivable,
    },
    "course-paid": {
      title: "Pago pelos alunos",
      valueLabel: "Valor Pago",
      formula: "Soma da coluna Valor Pago. Este valor é informativo e não representa necessariamente dinheiro creditado na conta do Ministério.",
      predicate: (record) => record.paid > 0,
      value: (record) => record.paid,
    },
    "course-fees": {
      title: "Tarifas e taxas do cartão",
      valueLabel: "Tarifa cobrada",
      formula: "Soma absoluta da coluna Despesa Financeira: tarifas e taxas de manutenção e processamento dos pagamentos no cartão.",
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
  const groupMap = new Map();
  records.forEach((record) => {
    if (!groupMap.has(record.donor)) groupMap.set(record.donor, { label: record.donor, total: 0, lines: [] });
    const group = groupMap.get(record.donor);
    const value = definition.value(record);
    group.total += value;
    group.lines.push({
      label: `Parcela ${record.installment || "—"}`,
      context: lineContexts[displayKey](record),
      value,
    });
  });
  const groups = [...groupMap.values()].sort((a, b) => b.total - a.total || a.label.localeCompare(b.label, "pt-BR"));

  return {
    title: definition.title,
    source: state.sourceName,
    formula: definition.formula,
    valueLabel: definition.valueLabel,
    total,
    calculation: calculations[displayKey],
    resultLabel: definition.title,
    groupTitle: "Composição por pessoa",
    groupHint: "Confira o subtotal de cada pessoa. Abra um nome somente se precisar ver as parcelas.",
    groups,
    columns: tableFormat.columns,
    rows: records.map((record) => tableFormat.row(record).map((cell, index, row) => {
      const isValue = typeof cell === "number";
      const content = isValue ? currency.format(cell) : safe(cell);
      return index === 0 || isValue && index === row.length - 1 ? `<strong>${content}</strong>` : content;
    })),
    note: `${number.format(records.length)} ${records.length === 1 ? "linha encontrada" : "linhas encontradas"} na base do curso.`,
  };
}

function ministryAuditBreakdown(key) {
  const definitions = {
    "ministry-all": {
      title: "Todos os lançamentos administrativos",
      valueLabel: "Valor das linhas",
      formula: "Todos os registros válidos da base administrativa do Ministério.",
      predicate: () => true,
    },
    "ministry-exits": {
      title: "Total de saídas registradas",
      valueLabel: "Valor da saída",
      formula: "Soma da coluna amount quando type = saída.",
      predicate: (record) => record.type.toLocaleLowerCase("pt-BR") === "saida",
    },
    "ministry-classes": {
      title: "Despesas do núcleo de aulas",
      valueLabel: "Valor da despesa",
      formula: "Soma da coluna amount quando finance_kind = aulas.",
      predicate: (record) => record.financeKind.toLocaleLowerCase("pt-BR") === "aulas",
    },
    "ministry-general": {
      title: "Despesas gerais do Ministério",
      valueLabel: "Valor da despesa",
      formula: "Soma da coluna amount quando finance_kind = geral.",
      predicate: (record) => record.financeKind.toLocaleLowerCase("pt-BR") === "geral",
    },
    "ministry-entries": {
      title: "Entradas administrativas",
      valueLabel: "Valor da entrada",
      formula: "Soma da coluna amount quando type = entrada.",
      predicate: (record) => record.type.toLocaleLowerCase("pt-BR") === "entrada",
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
    source: "finançassinaisdoreino.csv",
    formula: definition.formula,
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
      title: "Entradas no recorte comparado",
      source: `${state.sourceName} + finançassinaisdoreino.csv`,
      formula: "Valor Creditado do curso + registros administrativos cujo type = entrada.",
      valueLabel: "Valor da entrada",
      total: summary.credited + ministry.entries,
      calculation: [
        { label: "Créditos líquidos do curso", value: summary.credited },
        { operator: "+", label: "Entradas administrativas", value: ministry.entries },
      ],
      resultLabel: "Entradas no recorte",
      groupTitle: "Composição por base",
      groupHint: "Abra uma base para conferir os lançamentos que formam o subtotal.",
      groups,
      columns: ["Base", "Data", "Pessoa ou descrição", "Critério", "Valor da entrada"],
      rows: rows.map((row) => [safe(row[0]), safe(row[1]), `<strong>${safe(row[2])}</strong>`, safe(row[3]), `<strong>${currency.format(row[4])}</strong>`]),
      note: `${number.format(rows.length)} linhas formam a entrada comparada.`,
    };
  }
  if (key === "comparison-difference") {
    const entries = summary.credited + ministry.entries;
    return {
      title: "Diferença entre entradas e saídas",
      source: `${state.sourceName} + finançassinaisdoreino.csv`,
      formula: "Entradas líquidas do recorte − saídas administrativas do recorte. Não equivale ao saldo bancário completo.",
      valueLabel: "Efeito na diferença",
      total: entries - ministry.exits,
      calculation: [
        { label: "Entradas no recorte", value: entries },
        { operator: "−", label: "Saídas do Ministério", value: ministry.exits },
      ],
      resultLabel: "Diferença comparada",
      groupTitle: "Valores usados na comparação",
      groupHint: "Esta conta compara duas bases diferentes e não representa o saldo bancário completo.",
      groups: [
        { label: "Entradas no recorte", total: entries, lines: [{ label: "Curso + entradas administrativas", context: "Operação: somar", value: entries }] },
        { label: "Saídas do Ministério (subtração)", total: -ministry.exits, lines: [{ label: "Despesas administrativas", context: "Efeito negativo na comparação", value: -ministry.exits }] },
      ],
      columns: ["Componente", "Base de origem", "Operação", "Efeito na diferença"],
      rows: [
        ["Entradas no recorte", `${safe(state.sourceName)} + finanças administrativas`, "Somar", `<strong>${currency.format(entries)}</strong>`],
        ["Saídas do Ministério", "finançassinaisdoreino.csv", "Subtrair", `<strong>${currency.format(-ministry.exits)}</strong>`],
      ],
      note: "Esta comparação reúne duas bases de finalidades diferentes e serve apenas para leitura gerencial.",
    };
  }
  return null;
}

function getAuditBreakdown(key) {
  if (key.startsWith("course-") || key.startsWith("defaulter:")) return courseAuditBreakdown(key);
  if (key.startsWith("ministry-")) return ministryAuditBreakdown(key);
  return comparisonAuditBreakdown(key);
}

function renderAuditCalculation(breakdown) {
  const terms = (breakdown.calculation || []).map((term, index) => `
    ${index ? `<span class="audit-equation-operator" aria-hidden="true">${safe(term.operator || "+")}</span>` : ""}
    <div class="audit-equation-term">
      <span>${safe(term.label)}</span>
      <strong>${currency.format(term.value)}</strong>
    </div>`).join("");
  return `
    <section class="audit-calculation" aria-labelledby="audit-calculation-title">
      <div class="audit-section-heading">
        <div><p class="eyebrow">Conta usada</p><h3 id="audit-calculation-title">Como chegamos ao total</h3></div>
      </div>
      <div class="audit-equation">
        ${terms}
        <span class="audit-equation-operator audit-equation-equals" aria-hidden="true">=</span>
        <div class="audit-equation-result ${breakdown.total < 0 ? "audit-equation-negative" : ""}">
          <span>${safe(breakdown.resultLabel)}</span>
          <strong>${currency.format(breakdown.total)}</strong>
        </div>
      </div>
    </section>`;
}

function renderAuditGroups(breakdown) {
  if (!breakdown.groups?.length) {
    return `<div class="audit-groups-empty"><strong>Nenhum lançamento encontrado</strong><p>A regra deste card não encontrou valores na base atual. Por isso, o total é ${currency.format(0)}.</p></div>`;
  }
  return breakdown.groups.map((group, index) => `
    <details class="audit-composition-group">
      <summary>
        <span class="audit-group-index">${String(index + 1).padStart(2, "0")}</span>
        <span class="audit-group-name"><strong>${safe(group.label)}</strong><small>${number.format(group.lines.length)} ${group.lines.length === 1 ? "lançamento" : "lançamentos"}</small></span>
        <span class="audit-group-total ${group.total < 0 ? "audit-negative-value" : ""}"><small>Subtotal</small><strong>${currency.format(group.total)}</strong></span>
        <span class="audit-group-chevron" aria-hidden="true">⌄</span>
      </summary>
      <div class="audit-group-lines">
        ${group.lines.map((line) => `
          <div class="audit-group-line">
            <div><strong>${safe(line.label)}</strong><span>${safe(line.context)}</span></div>
            <strong class="${line.value < 0 ? "audit-negative-value" : ""}">${currency.format(line.value)}</strong>
          </div>`).join("")}
      </div>
    </details>`).join("");
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
        <div><p class="eyebrow">Auditoria do valor</p><h2 id="audit-breakdown-title" tabindex="-1">${safe(breakdown.title)}</h2><p>Veja primeiro a conta e depois abra somente os subtotais que quiser conferir.</p></div>
        <button class="close-button" type="button" aria-label="Fechar detalhamento">${icons.close}</button>
      </div>
      <div class="audit-breakdown-content">
        <div class="audit-total-hero ${breakdown.total < 0 ? "audit-total-negative" : ""}">
          <div><span>${safe(breakdown.valueLabel)}</span><strong>${currency.format(breakdown.total)}</strong></div>
          <div class="audit-total-facts"><span><b>${number.format(breakdown.rows.length)}</b> linhas usadas</span><span><b>${number.format(breakdown.groups?.length || 0)}</b> subtotais</span></div>
        </div>
        <div class="audit-criterion">
          <div><span>Critério aplicado</span><p>${safe(breakdown.formula)}</p></div>
          <small>Fonte: <b>${safe(breakdown.source)}</b></small>
        </div>
        ${renderAuditCalculation(breakdown)}
        <section class="audit-composition" aria-labelledby="audit-composition-title">
          <div class="audit-section-heading">
            <div><p class="eyebrow">De onde veio</p><h3 id="audit-composition-title">${safe(breakdown.groupTitle)}</h3><p>${safe(breakdown.groupHint)}</p></div>
          </div>
          <div class="audit-composition-list">${renderAuditGroups(breakdown)}</div>
        </section>
        <div class="audit-final-check"><span>${icons.check}</span><div><strong>Total conferido</strong><p>A soma dos subtotais acima fecha em <b>${currency.format(breakdown.total)}</b>.</p></div></div>
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
      fetch(ministryFinanceCsv),
    ]);
    if (!workbookResponse.ok) throw new Error("Não foi possível carregar Libras.xlsx.");
    if (!ministryResponse.ok) throw new Error("Não foi possível carregar finançassinaisdoreino.csv.");
    await readWorkbook(await workbookResponse.arrayBuffer(), "Libras.xlsx");
    readMinistryFinance(await ministryResponse.text());
    renderApp();
  } catch (error) {
    document.querySelector("#app").innerHTML = `<div class="error-screen"><div><h1>Não foi possível abrir o painel</h1><p>${safe(error.message)}</p></div></div>`;
  }
}

init();
