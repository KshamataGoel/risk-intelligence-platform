require('dotenv').config();
const express = require('express');
const path    = require('path');
const XLSX    = require('xlsx');
const fs      = require('fs');
const axios   = require('axios');

const app      = express();
const DATA_DIR = path.join(__dirname, 'data', 'insights');

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ─── Utilities ────────────────────────────────────────────────────────────────

function parseGBP(value) {
  if (value === null || value === undefined) return null;
  if (typeof value === 'number') return value;
  // Handle £, Â£ (mis-encoded), and comma formatting including Indian-style commas
  const cleaned = String(value)
    .replace(/Â£/g, '')
    .replace(/[££]/g, '')
    .replace(/,/g, '')
    .trim();
  const n = parseFloat(cleaned);
  return isNaN(n) ? null : n;
}

function cleanKey(k) {
  return String(k).trim().replace(/\s+/g, '_').replace(/[^\w]/g, '');
}

// ─── Data Quality & Insight Helpers ──────────────────────────────────────────

function cleanDataRows(rows, labelCol) {
  const JUNK = /^(total|grand total|subtotal|n\/a|unknown|not applicable|tbc|tbd|-)$/i;
  const clean = rows.filter(r => {
    const lv = String(r[labelCol] ?? '').trim();
    return lv && !JUNK.test(lv);
  });
  const excluded = rows.length - clean.length;
  const qualityNote = excluded > 0
    ? `${excluded} row(s) excluded — blank, N/A, or summary entries removed from analysis.`
    : null;
  return { clean, excluded, qualityNote };
}

function concentrationInsight(values, labels) {
  if (!values || values.length < 3) return null;
  const total = values.reduce((s, v) => s + v, 0);
  if (!total) return null;
  const sorted = labels
    ? [...labels.map((l, i) => ({ l, v: values[i] })).sort((a, b) => b.v - a.v)]
    : values.map((v, i) => ({ l: String(i), v })).sort((a, b) => b.v - a.v);
  const top3Sum = sorted.slice(0, 3).reduce((s, x) => s + x.v, 0);
  const pct = (top3Sum / total) * 100;
  if (pct >= 45) {
    const top3Names = sorted.slice(0, 3).map(x => x.l).join(', ');
    return `Material concentration dependency identified: top 3 segments (${top3Names}) account for ${pct.toFixed(0)}% of total.`;
  }
  return null;
}

function trendInsight(values, labels) {
  if (!values || values.length < 3) return null;
  const n = values.length;
  const firstHalf  = values.slice(0, Math.floor(n / 2));
  const secondHalf = values.slice(Math.ceil(n / 2));
  const avg1 = firstHalf.reduce((s, v) => s + v, 0) / (firstHalf.length || 1);
  const avg2 = secondHalf.reduce((s, v) => s + v, 0) / (secondHalf.length || 1);
  const chg  = ((avg2 - avg1) / (avg1 || 1)) * 100;
  const firstLabel = labels?.[0] || 'start';
  const lastLabel  = labels?.[n - 1] || 'end';
  if (chg > 10)  return `Upward trajectory: average activity in the latter half is ${chg.toFixed(0)}% higher than the first half (${firstLabel} → ${lastLabel}). Monitor for sustained growth pressure.`;
  if (chg < -10) return `Declining trend: activity has fallen ${Math.abs(chg).toFixed(0)}% from ${firstLabel} to ${lastLabel}. Investigate underlying drivers before next governance cycle.`;
  return `Transaction volumes are broadly stable across the reporting period with no material directional shift detected (${firstLabel} → ${lastLabel}).`;
}

function riskConcentrationInsight(labelCounts, total) {
  if (!labelCounts || !total) return null;
  const highRiskKeys  = Object.keys(labelCounts).filter(k => /(high|critical|very.?high)/i.test(k));
  const highRiskCount = highRiskKeys.reduce((s, k) => s + (labelCounts[k] || 0), 0);
  const pct = (highRiskCount / total) * 100;
  if (pct >= 30) return `Elevated risk concentration: ${pct.toFixed(0)}% of the customer population (${highRiskCount.toLocaleString()} customers) are classified as High or Critical risk. Board-level attention warranted.`;
  if (pct >= 15) return `Moderate risk concentration: ${pct.toFixed(0)}% of customers carry elevated risk ratings. Periodic deep-dive review recommended.`;
  return null;
}

// ─── Dataset Loading ──────────────────────────────────────────────────────────

function loadDatasets() {
  const result = {};
  if (!fs.existsSync(DATA_DIR)) return result;

  const files = fs.readdirSync(DATA_DIR).filter(f => /\.xlsx?$/i.test(f));
  if (!files.length) return result;

  for (const file of files) {
    try {
      const wb = XLSX.readFile(path.join(DATA_DIR, file));
      for (const sheetName of wb.SheetNames) {
        const ws = wb.Sheets[sheetName];
        // Detect if row 0 is a title and row 1 is the real header
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
        const cnt0 = (raw[0] || []).filter(c => c !== '').length;
        const cnt1 = (raw[1] || []).filter(c => c !== '').length;
        const headerRow = (raw.length >= 2 && (cnt1 > cnt0 || cnt0 <= 2)) ? 1 : 0;

        const rows = XLSX.utils.sheet_to_json(ws, { defval: null, range: headerRow });
        result[sheetName] = rows.map(row => {
          const r = {};
          for (const [k, v] of Object.entries(row)) r[cleanKey(k)] = v;
          return r;
        });
      }
    } catch (err) {
      console.error(`Error loading ${file}:`, err.message);
    }
  }
  return result;
}

// ─── Intent Router ────────────────────────────────────────────────────────────

function detectDataset(q) {
  const s = q.toLowerCase();

  // Parties — any mention of the word party/parties in any position
  if (/\bpart(y|ies)\b|party type|party count|party breakdown|party database|party.*risk|risk.*party/.test(s))
    return 'Parties';

  // Top Counterparty Exposures — counterpart, exposure (check early before other patterns)
  if (/counterpart|\bexposure(s)?\b|top.{0,15}exposure|largest.{0,15}exposure/.test(s))
    return 'Top10_Counterparty_Exposures';

  // Top Accounts / Markets — "top N" + market/banking/account signals; check BEFORE risk ratings
  // so "top 5 private banking market" routes here, not to Risk_Ratings
  if (/top\s*\d+.{0,40}(market|banking|bank|account|volume|transaction|financial|institution|country)|private\s*banking\s*market|(market|banking).{0,20}top\s*\d+|top.{0,15}account|account.{0,15}(volume|value|transaction|biggest|largest|highest)|highest.{0,15}(volume|account)|\baccount(s)?\b/.test(s))
    return 'Top10_Accounts_Volume';

  // Risk Ratings — risk level, tier, category, distribution, profile, score; explicit tier words
  if (/risk.{0,30}(rating|distribution|tier|level|categor|profile|score)|(rating|distribution|tier|level|categor|profile|score).{0,30}risk|customer.{0,10}risk|\b(high|medium|low|critical|not.?rated).{0,10}risk\b|\brisk.{0,10}(high|medium|low|critical|not.?rated)\b|\brisk\b.{0,30}\bpopulation\b|\bpopulation\b.{0,30}\brisk\b/.test(s))
    return 'Risk_Ratings';

  // Monthly Trends — monthly, trend, month, over time, transaction volume over time
  if (/monthly|month.{0,10}(trend|transaction|volume)|transaction.{0,10}(trend|month|over.{0,5}time)|\btrend(s)?\b|over time/.test(s))
    return 'Monthly_TXN_Trends';

  return null;
}

// Detect explicit chart type preference in the question
function detectChartType(q) {
  const s = q.toLowerCase();
  if (/(pie chart|pie graph|donut|doughnut|circular chart|as a pie)/.test(s)) return 'doughnut';
  if (/(bar chart|bar graph|horizontal bar|as a bar)/.test(s)) return 'bar';
  return null;
}

// Detect if question is asking about risk ranking within Parties
function isRiskRankQuestion(q) {
  return /(highest risk|most risk|riskiest|risk score|risk rank|carry.*risk|carry the highest|ranked by risk)/.test(q.toLowerCase());
}

// Detect if question is asking about the smallest/lowest value
function isLowestQuestion(q) {
  return /\b(lowest|smallest|least|minimum|min\b|fewest|tiniest|bottom|least.*common|rarest)\b/.test(q.toLowerCase());
}

// Find if any label is explicitly mentioned in the question — returns index or null
function detectMentionedLabel(question, labels) {
  const q = question.toLowerCase();
  // Try longest labels first so "Commercial - Large Corp" beats "Commercial"
  const sorted = labels.map((l, i) => ({ l, i })).sort((a, b) => b.l.length - a.l.length);
  for (const { l, i } of sorted) {
    if (q.includes(String(l).toLowerCase())) return i;
  }
  return null;
}

// Resolve which bar to highlight: explicit mention > min/max intent > default max
function resolveHighlight(question, labels, values) {
  const mentioned = detectMentionedLabel(question, labels);
  if (mentioned !== null) return { idx: mentioned, type: 'mention' };
  if (isLowestQuestion(question)) return { idx: values.indexOf(Math.min(...values)), type: 'min' };
  return { idx: values.indexOf(Math.max(...values)), type: 'max' };
}

// ─── Data Processors ─────────────────────────────────────────────────────────

function processParties(rows, opts = {}) {
  if (!rows.length) throw new Error('Parties sheet is empty');
  const cols  = Object.keys(rows[0]);
  const tCol  = cols.find(c => /type/i.test(c)) || cols[0];
  const cCol  = cols.find(c => /count/i.test(c));
  const rCol  = cols.find(c => /risk.?score|avg.?risk/i.test(c));

  const { clean: data, qualityNote } = cleanDataRows(rows, tCol);

  // ── Risk-ranked view ──────────────────────────────────────────────
  if (opts.rankByRisk && rCol) {
    const ranked = data
      .map(r => ({ type: String(r[tCol] ?? 'Unknown'), score: parseFloat(r[rCol]) || 0 }))
      .filter(d => d.score > 0)
      .sort((a, b) => b.score - a.score);

    return {
      answer_available: true,
      dataset_used: 'Parties',
      chart_type: 'horizontalBar',
      title: 'Party Types Ranked by Average Risk Score',
      chart_data: { labels: ranked.map(d => d.type), values: ranked.map(d => d.score) },
      key_metrics: {
        'Highest Risk Party Type': ranked[0]?.type  ?? 'N/A',
        'Risk Score (Top)':        ranked[0]?.score ?? 'N/A',
        'Party Types Assessed':    ranked.length,
        'Data Coverage':           `${data.length} active segments`
      },
      quality_note: qualityNote,
      columns: cols,
      preview: data.slice(0, 10)
    };
  }

  // ── Default: count by type ────────────────────────────────────────
  const counts = {};
  data.forEach(r => {
    const t = String(r[tCol] ?? 'Unknown');
    const n = cCol ? (parseFloat(r[cCol]) || 0) : 1;
    counts[t] = (counts[t] || 0) + n;
  });

  const labels = Object.keys(counts);
  const values = labels.map(l => counts[l]);
  const total  = values.reduce((s, v) => s + v, 0);
  const { idx: hlIdx, type: hlType } = resolveHighlight(opts.question || '', labels, values);

  const largestLabel = labels[values.indexOf(Math.max(...values))];
  const largestVal   = Math.max(...values);
  const largestPct   = total ? ((largestVal / total) * 100).toFixed(1) : '0';
  const concInsight  = concentrationInsight(values, labels);

  const metricLabel = hlType === 'min' ? 'Smallest Segment'
                    : hlType === 'mention' ? 'Highlighted Segment'
                    : 'Largest Segment';

  return {
    answer_available: true,
    dataset_used: 'Parties',
    chart_type: opts.chartType || 'bar',
    title: opts.chartType === 'doughnut' ? 'Party Type Distribution' : 'Party Count by Type',
    chart_data:            { labels, values },
    highlight_index:       hlIdx,
    highlight_type:        hlType,
    key_metrics: {
      'Total Registered Parties': total.toLocaleString(),
      'Distinct Party Types':     labels.length,
      'Dominant Segment':         `${largestLabel} (${largestPct}%)`,
      [metricLabel]:              labels[hlIdx]
    },
    concentration_insight: concInsight,
    quality_note:          qualityNote,
    columns: cols,
    preview: data.slice(0, 10)
  };
}

function processRiskRatings(rows, opts = {}) {
  if (!rows.length) throw new Error('Risk_Ratings sheet is empty');
  const cols = Object.keys(rows[0]);
  const rCol = cols.find(c => /risk.*rating|rating/i.test(c)) || cols[0];
  const cCol = cols.find(c => /count|customer/i.test(c));

  const { clean: data, qualityNote } = cleanDataRows(rows, rCol);

  const counts = {};
  data.forEach(r => {
    const t = String(r[rCol] ?? 'Unknown').trim();
    const n = cCol ? (parseFloat(r[cCol]) || 0) : 1;
    counts[t] = (counts[t] || 0) + n;
  });

  const ORDER = ['low', 'medium', 'high', 'very high', 'critical', 'not rated'];
  const labels = Object.keys(counts).sort((a, b) => {
    const ai = ORDER.findIndex(o => a.toLowerCase().includes(o));
    const bi = ORDER.findIndex(o => b.toLowerCase().includes(o));
    return (ai < 0 ? 99 : ai) - (bi < 0 ? 99 : bi);
  });
  const values = labels.map(l => counts[l]);
  const total  = values.reduce((s, v) => s + v, 0);
  const { idx: hlIdx, type: hlType } = resolveHighlight(opts.question || '', labels, values);

  const highRiskKeys  = labels.filter(l => /(high|critical|very.?high)/i.test(l));
  const highRiskCount = highRiskKeys.reduce((s, k) => s + (counts[k] || 0), 0);
  const highRiskPct   = total ? ((highRiskCount / total) * 100).toFixed(1) : '0';
  const notRatedKeys  = labels.filter(l => /not.?rated/i.test(l));
  const notRatedCount = notRatedKeys.reduce((s, k) => s + (counts[k] || 0), 0);
  const riskInsight   = riskConcentrationInsight(counts, total);

  const tierLabel = hlType === 'min' ? 'Least Common Tier'
                  : hlType === 'mention' ? 'Highlighted Tier'
                  : 'Most Common Tier';

  return {
    answer_available: true,
    dataset_used: 'Risk_Ratings',
    chart_type: 'doughnut',
    title: 'Risk Rating Distribution',
    chart_data:            { labels, values },
    highlight_index:       hlIdx,
    highlight_type:        hlType,
    key_metrics: {
      'Total Customers':      total.toLocaleString(),
      'High / Critical Risk': `${highRiskCount.toLocaleString()} (${highRiskPct}%)`,
      'Not Rated':            notRatedCount.toLocaleString(),
      [tierLabel]:            labels[hlIdx]
    },
    concentration_insight: riskInsight,
    quality_note:          qualityNote,
    columns: cols,
    preview: data.slice(0, 10)
  };
}

function processTopAccounts(rows, opts = {}) {
  if (!rows.length) throw new Error('Top10_Accounts_Volume sheet is empty');
  const cols = Object.keys(rows[0]);
  const nCol = cols.find(c => /party.?name|account.?name|\bname\b/i.test(c))
            || cols.find(c => /name|party|account/i.test(c))
            || cols[0];
  const vCol = cols.find(c => /volume|total|amount|value|txn/i.test(c)) || cols[cols.length - 1];

  const { clean: cleanRows, qualityNote } = cleanDataRows(rows, nCol);

  const data = cleanRows
    .map(r => ({ name: String(r[nCol] ?? 'Unknown'), volume: parseGBP(r[vCol]) ?? 0 }))
    .sort((a, b) => b.volume - a.volume)
    .slice(0, 10);

  const labels     = data.map(d => d.name);
  const values     = data.map(d => d.volume);
  const totalVol   = values.reduce((s, v) => s + v, 0);
  const topVolume  = values[0] || 0;
  const topShare   = totalVol ? ((topVolume / totalVol) * 100).toFixed(1) : '0';
  const concInsight = concentrationInsight(values, labels);
  const { idx: hlIdx, type: hlType } = resolveHighlight(opts.question || '', labels, values);

  return {
    answer_available: true,
    dataset_used: 'Top10_Accounts_Volume',
    chart_type: 'horizontalBar',
    title: 'Top 10 Accounts by Transaction Volume',
    chart_data:            { labels, values },
    highlight_index:       hlIdx,
    highlight_type:        hlType,
    key_metrics: {
      'Top Account':          labels[0] ?? 'N/A',
      'Top Account Volume':   `£${topVolume.toLocaleString()}`,
      'Top Account Share':    `${topShare}% of top 10`,
      'Top 10 Combined':      `£${totalVol.toLocaleString()}`
    },
    concentration_insight: concInsight,
    quality_note:          qualityNote,
    columns: cols,
    preview: cleanRows.slice(0, 10)
  };
}

function processMonthlyTrends(rows) {
  if (!rows.length) throw new Error('Monthly_TXN_Trends sheet is empty');
  const cols = Object.keys(rows[0]);
  const mCol = cols.find(c => /month|date|period/i.test(c)) || cols[0];
  const vCol = cols.find(c => /volume|total|amount|value|txn/i.test(c)) || cols[cols.length - 1];

  const { clean: cleanRows, qualityNote } = cleanDataRows(rows, mCol);
  const data   = cleanRows.map(r => ({ month: String(r[mCol] ?? ''), volume: parseGBP(r[vCol]) ?? 0 }));
  const labels = data.map(d => d.month);
  const values = data.map(d => d.volume);
  const total  = values.reduce((s, v) => s + v, 0);
  const avg    = total / (data.length || 1);
  const max    = Math.max(...values);
  const min    = Math.min(...values);
  const peak   = data.find(d => d.volume === max);
  const trough = data.find(d => d.volume === min);
  const trendNote = trendInsight(values, labels);

  return {
    answer_available: true,
    dataset_used: 'Monthly_TXN_Trends',
    chart_type: 'line',
    title: 'Monthly Transaction Volume Trends',
    chart_data:            { labels, values },
    key_metrics: {
      'Total Volume':    `£${total.toLocaleString()}`,
      'Monthly Average': `£${Math.round(avg).toLocaleString()}`,
      'Peak Month':      `${peak?.month ?? 'N/A'} (£${max.toLocaleString()})`,
      'Lowest Month':    `${trough?.month ?? 'N/A'} (£${min.toLocaleString()})`
    },
    concentration_insight: trendNote,
    quality_note:          qualityNote,
    columns: cols,
    preview: cleanRows.slice(0, 10)
  };
}

function processCounterparties(rows, opts = {}) {
  if (!rows.length) throw new Error('Top10_Counterparty_Exposures sheet is empty');
  const cols = Object.keys(rows[0]);
  const nCol = cols.find(c => /counterpart|name|party/i.test(c)) || cols[0];
  const eCol = cols.find(c => /exposure|gross|amount|value/i.test(c)) || cols[cols.length - 1];

  const { clean: cleanRows, qualityNote } = cleanDataRows(rows, nCol);

  const data = cleanRows
    .map(r => ({ name: String(r[nCol] ?? 'Unknown'), exposure: parseGBP(r[eCol]) ?? 0 }))
    .sort((a, b) => b.exposure - a.exposure)
    .slice(0, 10);

  const labels      = data.map(d => d.name);
  const values      = data.map(d => d.exposure);
  const totalExp    = values.reduce((s, v) => s + v, 0);
  const topExposure = values[0] || 0;
  const topShare    = totalExp ? ((topExposure / totalExp) * 100).toFixed(1) : '0';
  const concInsight = concentrationInsight(values, labels);
  const { idx: hlIdx, type: hlType } = resolveHighlight(opts.question || '', labels, values);

  return {
    answer_available: true,
    dataset_used: 'Top10_Counterparty_Exposures',
    chart_type: 'horizontalBar',
    title: 'Top 10 Counterparty Exposures',
    chart_data:            { labels, values },
    highlight_index:       hlIdx,
    highlight_type:        hlType,
    key_metrics: {
      'Largest Counterparty':   labels[0] ?? 'N/A',
      'Largest Exposure':       `£${topExposure.toLocaleString()}`,
      'Concentration Share':    `${topShare}% of top 10`,
      'Total Top-10 Exposure':  `£${totalExp.toLocaleString()}`
    },
    concentration_insight: concInsight,
    quality_note:          qualityNote,
    columns: cols,
    preview: cleanRows.slice(0, 10)
  };
}

// ─── Summaries ────────────────────────────────────────────────────────────────

function templateSummary(result) {
  const { dataset_used: ds, key_metrics: m, concentration_insight: ci, quality_note: qn } = result;

  const sections = {
    Parties: {
      A: `The party database contains ${m['Total Registered Parties']} registered counterparties segmented across ${m['Distinct Party Types']} distinct party types.`,
      B: `The dominant segment is ${m['Dominant Segment']}, representing the largest share of the registered population. Each segment carries distinct risk profiles and monitoring obligations.`,
      C: `High-volume or concentrated segments warrant enhanced due diligence review. Portfolio diversification across party types should be assessed against risk appetite limits.`,
      D: `Segment size directly influences onboarding capacity, KYC review cycles, and periodic refresh workloads. Imbalanced distributions may signal emerging concentration risk.`,
      E: ci || 'No material concentration dependency identified across party segments.',
      F: qn || 'All data rows processed. No quality issues identified.'
    },
    Risk_Ratings: {
      A: `The risk profile covers ${m['Total Customers']} customers with High or Critical risk representing ${m['High / Critical Risk']} of the population.`,
      B: `Not-rated customers stand at ${m['Not Rated']}, representing a blind spot in the risk management framework. The most prevalent rated tier is ${m['Most Common Tier']}.`,
      C: `Elevated high-risk concentration increases exposure to AML, sanctions, and reputational risk. Not-rated customers require prioritised review and risk assessment completion.`,
      D: `Risk tier distribution is the primary driver for KYC review frequency, enhanced due diligence triggers, and transaction monitoring calibration thresholds.`,
      E: ci || 'Risk distribution is within expected bounds for the current portfolio profile.',
      F: qn || 'All risk rating records processed without exclusions.'
    },
    Top10_Accounts_Volume: {
      A: `The top 10 accounts by transaction volume are identified, with ${m['Top Account']} leading at ${m['Top Account Volume']}.`,
      B: `The leading account represents ${m['Top Account Share']}, with combined top-10 volume at ${m['Top 10 Combined']}. This signals material single-name concentration.`,
      C: `Accounts with disproportionate transaction volumes require enhanced transaction monitoring and periodic business rationale reviews.`,
      D: `Volume-driven concentration risk may affect liquidity planning, limit utilisation, and correspondent banking relationships.`,
      E: ci || 'Transaction volume is distributed across the top 10 accounts without extreme single-name dependency.',
      F: qn || 'All account records processed. No quality issues identified.'
    },
    Monthly_TXN_Trends: {
      A: `Total transaction volume over the reporting period reached ${m['Total Volume']}, with a monthly average of ${m['Monthly Average']}.`,
      B: `Volume peaked in ${m['Peak Month']} and reached its lowest point in ${m['Lowest Month']}. This range indicates the scale of intra-period volatility.`,
      C: `Significant volume deviations from the monthly average may indicate seasonal patterns, business events, or anomalous activity requiring investigation.`,
      D: `Monthly transaction volumes inform liquidity forecasting, correspondent banking capacity planning, and threshold-based monitoring calibration.`,
      E: ci || 'No material directional shift detected across the reporting period.',
      F: qn || 'All monthly records processed. No quality issues identified.'
    },
    Top10_Counterparty_Exposures: {
      A: `The largest counterparty exposure is with ${m['Largest Counterparty']} at ${m['Largest Exposure']}, representing ${m['Concentration Share']} of the top-10 aggregate.`,
      B: `Total top-10 counterparty exposure stands at ${m['Total Top-10 Exposure']}. The leading counterparty's disproportionate share warrants limit review.`,
      C: `Concentration in single counterparties amplifies credit, settlement, and systemic risk. Regulatory large exposure limits must be verified against current thresholds.`,
      D: `Counterparty exposure concentration is a key driver of credit risk appetite utilisation, PFE calculations, and stress testing outcomes.`,
      E: ci || 'Exposure is distributed across the top 10 counterparties without extreme single-name concentration.',
      F: qn || 'All counterparty records processed. No quality issues identified.'
    }
  };

  const s = sections[ds];
  if (!s) return `Analysis complete for ${ds}. ${Object.entries(m).map(([k, v]) => `${k}: ${v}`).join('. ')}.`;

  return [
    `A. Situation Summary: ${s.A}`,
    `B. Key Findings: ${s.B}`,
    `C. Risk Interpretation: ${s.C}`,
    `D. Material Drivers: ${s.D}`,
    `E. Concentration / Trend: ${s.E}`,
    `F. Data Quality Note: ${s.F}`
  ].join('\n');
}

async function groqSummary(question, result) {
  const key = process.env.GROQ_API_KEY;
  if (!key) return templateSummary(result);

  const dataRows = (result.chart_data?.labels || [])
    .map((lbl, i) => `  ${lbl}: ${result.chart_data.values[i]}`)
    .join('\n');

  const systemPrompt =
    `You are a senior Risk & Compliance intelligence analyst briefing the Chief Risk & Compliance Officer (CRCO). ` +
    `You produce structured, data-driven executive intelligence. You never hallucinate, invent names, or extrapolate beyond the data provided. ` +
    `Your language is institutional, precise, and board-ready.`;

  const userPrompt =
    `QUESTION: ${question}\n` +
    `DATASET: ${result.dataset_used.replace(/_/g, ' ')}\n` +
    `KEY METRICS:\n${Object.entries(result.key_metrics).map(([k, v]) => `  ${k}: ${v}`).join('\n')}\n` +
    (dataRows ? `DATA BREAKDOWN:\n${dataRows}\n` : '') +
    (result.concentration_insight ? `CONCENTRATION SIGNAL: ${result.concentration_insight}\n` : '') +
    (result.quality_note ? `DATA QUALITY: ${result.quality_note}\n` : '') +
    `\nUsing ONLY the data above (no external knowledge), write a structured executive intelligence summary with EXACTLY these 6 labelled sections:\n\n` +
    `A. Situation Summary: [One sentence — the direct answer using actual names and numbers from the data]\n` +
    `B. Key Findings: [Two to three specific observations drawn from the data — use exact values]\n` +
    `C. Risk Interpretation: [One sentence on the risk implication for the CRO]\n` +
    `D. Material Drivers: [One sentence on what is driving the pattern observed]\n` +
    `E. Concentration / Trend: [One sentence on concentration risk or trend direction — use the concentration signal if provided]\n` +
    `F. Data Quality Note: [One sentence on data completeness or exclusions — use the quality note if provided, otherwise state "All records processed."]\n\n` +
    `Use exact names and values from the data. Do not invent any information.`;

  try {
    const res = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user',   content: userPrompt }
        ],
        max_tokens: 600,
        temperature: 0.15
      },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    return res.data.choices?.[0]?.message?.content || templateSummary(result);
  } catch (err) {
    console.error('Groq error:', err.message);
    return templateSummary(result);
  }
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/api/catalogue', (req, res) => {
  const datasets = loadDatasets();
  res.json({
    datasets: Object.entries(datasets).map(([name, rows]) => ({
      name,
      rowCount: rows.length,
      columns: rows[0] ? Object.keys(rows[0]) : []
    })),
    hasData: Object.keys(datasets).length > 0
  });
});

app.post('/api/insights', async (req, res) => {
  const { question } = req.body || {};
  if (!question?.trim()) return res.status(400).json({ error: 'Question is required' });

  const datasets = loadDatasets();
  if (!Object.keys(datasets).length) {
    return res.json({
      answer_available: false,
      message: 'No datasets found. Please place Dataset.xlsx inside the data/insights/ folder and restart the server.'
    });
  }

  const target        = detectDataset(question);
  const preferredChart = detectChartType(question);
  const rankByRisk    = isRiskRankQuestion(question);
  const findLowest    = isLowestQuestion(question);

  if (!target || !datasets[target]) {
    return res.json({
      answer_available: false,
      message: `I cannot answer this from the currently available dataset. Please upload the required dataset or select a question related to the available data.\n\nAvailable datasets: ${Object.keys(datasets).join(', ')}`
    });
  }

  let result;
  try {
    switch (target) {
      case 'Parties':
        result = processParties(datasets[target], { chartType: preferredChart, rankByRisk, findLowest, question });
        break;
      case 'Risk_Ratings':
        result = processRiskRatings(datasets[target], { question });
        break;
      case 'Top10_Accounts_Volume':
        result = processTopAccounts(datasets[target], { question });
        break;
      case 'Monthly_TXN_Trends':
        result = processMonthlyTrends(datasets[target]);
        break;
      case 'Top10_Counterparty_Exposures':
        result = processCounterparties(datasets[target], { question });
        break;
      default: throw new Error('No processor available');
    }
  } catch (err) {
    return res.status(500).json({ error: `Data processing failed: ${err.message}` });
  }

  result.summary = await groqSummary(question, result);
  res.json(result);
});

// ─── Analysis Module ──────────────────────────────────────────────────────────

const ANALYSIS_DIR = path.join(__dirname, 'data', 'analysis');

const ANALYSIS_EXPECTED = {
  Feature_Catalogue: ['Category', 'Feature Name', 'Description', 'Tags', 'Risk Relevance'],
  Scatter_Plot_Data: ['Party ID', 'Cluster', 'Cluster Label', 'Risk Rating', 'avg_txn_amount_gbp', 'std_txn_amount_gbp'],
  Anomaly_Detection: ['Rank', 'Party ID', 'Cluster', 'Cluster Label', 'Distance from Centroid', 'Anomaly Type', 'Investigation Status']
};

function cleanAnalysisCol(k) {
  return String(k).replace(/\r\n/g, ' ').replace(/\n/g, ' ').replace(/\r/g, ' ').replace(/\s+/g, ' ').trim();
}

function getNumericCols(rows) {
  if (!rows || !rows.length) return [];
  const sample = rows.slice(0, Math.min(20, rows.length));
  return Object.keys(rows[0]).filter(col => {
    if (/^(party.?id|cluster.?label|cluster$|risk.?rating|flag|status|type|label|note|narrative|rank|invest|anomaly)/i.test(col)) return false;
    const vals = sample.map(r => r[col]).filter(v => v !== null && v !== undefined && v !== '');
    if (!vals.length) return false;
    const numCount = vals.filter(v => typeof v === 'number' || (!isNaN(parseFloat(v)) && String(v).trim() !== '')).length;
    return numCount / vals.length >= 0.7;
  });
}

function detectAnalysisHeaderRow(ws, expected) {
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  // First pass: find a row with multiple non-empty cells that match expected column names.
  // Rows with only 1 non-empty cell are title rows (all content packed into column A).
  for (let i = 0; i < Math.min(10, raw.length); i++) {
    const cells = raw[i].map(c => cleanAnalysisCol(String(c)).toLowerCase()).filter(c => c !== '');
    if (cells.length <= 1) continue; // title rows have only one populated cell
    const hits = expected.filter(e =>
      cells.some(cell => cell === e.toLowerCase() || cell.includes(e.toLowerCase()) || e.toLowerCase().includes(cell))
    );
    if (hits.length >= Math.max(2, Math.floor(expected.length / 3))) return i;
  }

  // Fallback: first row that has more than 3 non-empty cells (likely a real header)
  for (let i = 0; i < Math.min(10, raw.length); i++) {
    if (raw[i].filter(c => c !== '').length > 3) return i;
  }
  return 2;
}

function loadAnalysisData() {
  if (!fs.existsSync(ANALYSIS_DIR)) return null;
  const files = fs.readdirSync(ANALYSIS_DIR).filter(f => /\.xlsx?$/i.test(f));
  if (!files.length) return null;

  const result = {};
  for (const file of files) {
    try {
      const wb = XLSX.readFile(path.join(ANALYSIS_DIR, file));
      for (const sheet of wb.SheetNames) {
        const ws       = wb.Sheets[sheet];
        const expected = ANALYSIS_EXPECTED[sheet] || [];
        const hRow     = expected.length ? detectAnalysisHeaderRow(ws, expected) : 1;
        const rows     = XLSX.utils.sheet_to_json(ws, { defval: null, range: hRow });
        let mapped = rows
          .map(row => {
            const r = {};
            for (const [k, v] of Object.entries(row)) r[cleanAnalysisCol(k)] = v;
            return r;
          })
          .filter(r => Object.values(r).some(v => v !== null && v !== '' && v !== undefined));

        // Forward-fill sparse columns (Category, Cluster Label, etc.)
        // where Excel writes the value once and leaves subsequent rows blank
        const fillCols = ['Category', 'Cluster Label', 'Cluster'];
        const lastVal  = {};
        mapped = mapped.map(r => {
          for (const col of fillCols) {
            if (col in r) {
              if (r[col] !== null && r[col] !== '' && r[col] !== undefined) {
                lastVal[col] = r[col];
              } else if (lastVal[col] != null) {
                r[col] = lastVal[col];
              }
            }
          }
          return r;
        });

        result[sheet] = mapped;
      }
    } catch (err) {
      console.error(`Analysis load error (${file}): ${err.message}`);
    }
  }
  return result;
}

function computeClusterStats(rows, selectedCols) {
  const clusters = {};
  rows.forEach(r => {
    const label = String(r['Cluster Label'] || r['Cluster'] || 'Unknown');
    if (!clusters[label]) clusters[label] = [];
    clusters[label].push(r);
  });

  const mean = arr => arr.reduce((s, v) => s + (parseFloat(v) || 0), 0) / (arr.length || 1);
  const colsToUse = (selectedCols && selectedCols.length > 0) ? selectedCols : getNumericCols(rows);

  return Object.entries(clusters).map(([label, items]) => {
    const riskCounts = {};
    items.forEach(r => { const rr = r['Risk Rating'] || 'Unknown'; riskCounts[rr] = (riskCounts[rr] || 0) + 1; });

    const colStats = {};
    colsToUse.forEach(col => {
      const vals = items.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
      if (vals.length > 0) colStats[col] = parseFloat(mean(vals).toFixed(2));
    });

    return { label, parties: items.length, column_averages: colStats, risk_distribution: riskCounts };
  });
}

function templateClusterDesc(s) {
  const topRisk    = Object.entries(s.risk_distribution).sort((a, b) => b[1] - a[1])[0];
  const colEntries = Object.entries(s.column_averages || {});
  const parties    = s.parties;
  const dominantRisk = topRisk?.[0] || 'Unknown';

  const avgTxnEntry = colEntries.find(([k]) => /avg.?txn|avg.?amount|avg.?value|avg_txn/i.test(k));
  const avgTxnVal   = avgTxnEntry ? Number(avgTxnEntry[1]) : null;

  let description;
  if (avgTxnVal !== null && avgTxnVal >= 500000) {
    description =
      `Cluster contains ${parties} parties exhibiting materially higher-value transaction profiles consistent with corporate or private banking relationships. ` +
      `Elevated average transaction values and concentrated payment flows indicate enhanced monitoring relevance. ` +
      `Behaviour patterns suggest cross-border complexity and large-value flow indicators warranting enhanced review. ` +
      `Predominant risk rating within this segment is ${dominantRisk}.`;
  } else if (avgTxnVal !== null && avgTxnVal >= 50000) {
    description =
      `Cluster contains ${parties} parties demonstrating elevated commercial transaction throughput with moderate behavioural variability. ` +
      `Increased transaction dispersion and cross-border activity indicate the need for continued monitoring of payment purpose, counterparty profile, and volume consistency. ` +
      `Cluster behaviour is consistent with mid-tier commercial risk typology indicators. ` +
      `Predominant risk rating is ${dominantRisk}.`;
  } else if (avgTxnVal !== null) {
    description =
      `Cluster contains ${parties} parties exhibiting lower-value transactional behaviour with relatively stable transaction volatility and limited cross-border complexity. ` +
      `Behaviour patterns are broadly aligned to lower inherent AML exposure profiles, subject to normal monitoring thresholds. ` +
      `Predominant risk rating is ${dominantRisk}.`;
  } else {
    const colSummary = colEntries.slice(0, 4).map(([k, v]) => {
      const isGBP = /gbp|amount/i.test(k);
      return `${k}: ${isGBP ? '£' + Number(v).toLocaleString() : v}`;
    }).join('; ');
    description =
      `Cluster contains ${parties} parties. ` +
      (colSummary ? `Feature averages — ${colSummary}. ` : '') +
      `Behaviour patterns require review against customer profile and expected activity thresholds. ` +
      `Predominant risk rating is ${dominantRisk}.`;
  }

  return { label: s.label, description };
}

function templateInsightsSummary(clusterDescriptions, anomalySummary) {
  const { total = 0, highRisk = 0, escalated = 0 } = anomalySummary || {};
  const clusterCount = clusterDescriptions?.length || 0;
  const investigating = total - highRisk;

  return [
    `1. Behavioural Segmentation Summary: Behavioural clustering has identified ${clusterCount} distinct customer segment${clusterCount !== 1 ? 's' : ''} based on transaction value, volatility, and cross-border activity patterns. Each cluster exhibits differentiated AML exposure characteristics requiring proportionate monitoring calibration.`,
    `2. Key Risk Indicators: The highest-exposure segments demonstrate elevated average transaction values and concentrated payment flows, consistent with cross-border complexity and large-value flow typology indicators. Lower-exposure segments exhibit stable, lower-value transactional behaviour broadly aligned to standard monitoring thresholds.`,
    `3. Anomaly Concentration: Anomaly detection identified ${total} record${total !== 1 ? 's' : ''} requiring monitoring attention, of which ${highRisk} carr${highRisk !== 1 ? 'y' : 'ies'} a High risk rating. ${escalated} record${escalated !== 1 ? 's are' : ' is'} currently escalated; remaining records are under investigation, enhanced review, or monitoring disposition.`,
    `4. Monitoring Implication: The observed behavioural patterns indicate that existing monitoring thresholds may require recalibration across high-exposure segments. Enhanced due diligence review is indicated for all High-risk and escalated records prior to any disposition decision.`,
    `5. Suggested Next Action: Priority review is indicated for all escalated and High-risk records. Validate supporting evidence, transaction purpose, counterparty relationships, and source of funds documentation before determining further action or case closure.`
  ].join('\n\n');
}

async function generateClusterDescriptions(stats) {
  const key = process.env.GROQ_API_KEY;
  if (!key) return stats.map(templateClusterDesc);

  const systemPrompt =
    `You are a senior AML analytics platform generating cluster behavioural descriptions for Chief Risk Officer decision-support. ` +
    `Strict tone rules: Do not use first person. Do not say "our analysis reveals" or "I recommend". ` +
    `Do not make accusatory statements. Do not say customers are involved in money laundering or terrorist financing. ` +
    `Never use: suspicious, fraudulent, illegal, criminal, money laundering, terrorist. ` +
    `Use institutional risk-indicator language only: ` +
    `"indicates elevated monitoring attention", "suggests behavioural deviation", "requires enhanced review", ` +
    `"may warrant further investigation", "is consistent with risk typology indicators", ` +
    `"potential financial crime risk indicators", "AML monitoring indicators", "behavioural outlier", ` +
    `"transaction concentration", "cross-border complexity", "structuring indicator". ` +
    `All output must be derived solely from the cluster data provided. Do not invent any figures or names.`;

  const userPrompt =
    `Cluster data:\n${JSON.stringify(stats, null, 2)}\n\n` +
    `For each cluster, write a description matching the most appropriate tier below based on avg transaction values:\n\n` +
    `LOWER-VALUE / RETAIL:\n"Cluster contains [N] parties exhibiting lower-value transactional behaviour with relatively stable transaction volatility and limited cross-border complexity. Behaviour patterns are broadly aligned to lower inherent AML exposure profiles, subject to normal monitoring thresholds. Predominant risk rating is [X]."\n\n` +
    `MID-TIER COMMERCIAL:\n"Cluster contains [N] parties demonstrating elevated commercial transaction throughput with moderate behavioural variability. Increased cross-border activity and transaction dispersion indicate the need for continued monitoring of payment purpose, counterparty profile, and volume consistency. Cluster behaviour is consistent with mid-tier commercial risk typology indicators."\n\n` +
    `HIGH-VALUE CORPORATE / PB:\n"Cluster contains [N] parties exhibiting materially higher-value transaction profiles consistent with corporate or private banking relationships. Elevated average transaction values and concentrated payment flows indicate enhanced monitoring relevance. Behaviour patterns suggest cross-border complexity and large-value flow indicators warranting enhanced review."\n\n` +
    `Keep each description under 80 words. Use exact party counts from the data.\n\n` +
    `Return a JSON array ONLY: [{"label": "...", "description": "..."}]`;

  try {
    const res = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user',   content: userPrompt }
        ],
        max_tokens: 800,
        temperature: 0.15
      },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    const text = res.data.choices?.[0]?.message?.content || '';
    const match = text.match(/\[[\s\S]*\]/);
    if (match) return JSON.parse(match[0]);
    return stats.map(templateClusterDesc);
  } catch (err) {
    console.error('Groq cluster error:', err.message);
    return stats.map(templateClusterDesc);
  }
}

// ─── Analysis Routes ──────────────────────────────────────────────────────────

app.get('/api/analysis/data', (req, res) => {
  const data = loadAnalysisData();
  if (!data) {
    return res.json({ available: false, message: 'Analysis dataset not found. Please place the Excel file in data/analysis/.' });
  }
  const missing = ['Feature_Catalogue', 'Scatter_Plot_Data', 'Anomaly_Detection'].filter(s => !data[s]);
  if (missing.length) {
    return res.json({ available: false, message: `Required sheets missing: ${missing.join(', ')}` });
  }
  const scatterNumericCols = getNumericCols(data['Scatter_Plot_Data']);
  res.json({
    available: true,
    feature_catalogue:    data['Feature_Catalogue'],
    scatter_data:         data['Scatter_Plot_Data'],
    anomaly_data:         data['Anomaly_Detection'],
    scatter_numeric_cols: scatterNumericCols,
    loaded_at:            new Date().toISOString()
  });
});

app.post('/api/analysis/cluster-descriptions', async (req, res) => {
  const data = loadAnalysisData();
  if (!data?.['Scatter_Plot_Data']) return res.status(404).json({ error: 'Scatter_Plot_Data sheet not found' });
  const { selectedCols } = req.body || {};
  const stats = computeClusterStats(data['Scatter_Plot_Data'], selectedCols);
  const descriptions = await generateClusterDescriptions(stats);
  res.json({ descriptions, stats });
});

app.post('/api/analysis/insights-summary', async (req, res) => {
  const { clusterDescriptions, anomalySummary } = req.body || {};
  const key = process.env.GROQ_API_KEY;
  if (!key) return res.json({ summary: templateInsightsSummary(clusterDescriptions, anomalySummary) });

  const systemPrompt =
    `You are a senior AML Risk & Compliance analytics platform generating executive intelligence summaries for Chief Risk Officer decision-support. ` +
    `Strict tone rules: Do not use first person. Do not say "our analysis reveals" or "I recommend". ` +
    `Do not make accusatory statements. Do not say customers are involved in money laundering or terrorist financing. ` +
    `Never use: suspicious, fraudulent, illegal, criminal, money laundering, terrorist financing. ` +
    `Use institutional risk-indicator language: ` +
    `"indicates elevated monitoring attention", "suggests behavioural deviation", "requires enhanced review", ` +
    `"potential financial crime risk indicators", "AML monitoring indicators", "behavioural outlier", ` +
    `"transaction concentration", "structuring indicator", "sanctions corridor exposure pattern". ` +
    `Use ONLY the data provided. If data is not available, state: "Not available in source dataset."`;

  const userPrompt =
    `CLUSTER DESCRIPTIONS:\n${JSON.stringify(clusterDescriptions, null, 2)}\n\n` +
    `ANOMALY SUMMARY:\n${JSON.stringify(anomalySummary, null, 2)}\n\n` +
    `Generate an Analytics Intelligence Summary with EXACTLY these 5 numbered sections:\n\n` +
    `1. Behavioural Segmentation Summary: [Describe customer cluster profile distribution using risk-indicator language. Reference cluster count and segment characteristics.]\n` +
    `2. Key Risk Indicators: [Identify the material risk signals from the cluster data — transaction values, volatility, cross-border complexity, concentration.]\n` +
    `3. Anomaly Concentration: [Quantify anomaly distribution — state total anomalies, high-risk count, escalated count, and investigation dispositions from the data.]\n` +
    `4. Monitoring Implication: [State what the observed behavioural patterns mean for AML monitoring threshold calibration and oversight requirements.]\n` +
    `5. Suggested Next Action: [One concise institutional action statement — not "I recommend". Use: "Priority review is indicated for..." or "Enhanced due diligence review is required..."]\n\n` +
    `Use exact numbers from the data. Keep each section to 2-3 sentences.`;

  try {
    const response = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user',   content: userPrompt }
        ],
        max_tokens: 600,
        temperature: 0.15
      },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    res.json({ summary: response.data.choices?.[0]?.message?.content || templateInsightsSummary(clusterDescriptions, anomalySummary) });
  } catch (err) {
    console.error('Insights summary error:', err.message);
    res.json({ summary: templateInsightsSummary(clusterDescriptions, anomalySummary) });
  }
});

// ─── Vault Module ─────────────────────────────────────────────────────────────

let mammoth;
try { mammoth = require('mammoth'); } catch {}

const VAULT_DIR = path.join(__dirname, 'data', 'vault');

function extractVaultMeta(rawLines) {
  const meta = {};
  const labelMap = {
    'date': 'date', 'author': 'author', 'from': 'author',
    'classification': 'classification', 'status': 'status',
    'to': 'to', 'subject': 'subject', 'cc': 'cc',
    'deadline': 'deadline', 'source': 'source',
    'owner': 'owner', 'review cycle': 'reviewCycle',
    'prepared by': 'author', 'issued by': 'author',
  };

  for (let i = 0; i < rawLines.length; i++) {
    const t = rawLines[i].trim();
    if (!t) continue;

    // "Key: value" on same line
    const m1 = t.match(/^([A-Za-z][A-Za-z\s\/]{0,25}):\s+(.+)$/);
    if (m1) {
      const k = labelMap[m1[1].toLowerCase().trim()];
      if (k && !meta[k]) meta[k] = m1[2].trim();
      continue;
    }

    // "Key:" alone on a line → value is the next non-empty line
    const m2 = t.match(/^([A-Za-z][A-Za-z\s\/]{0,25}):$/);
    if (m2) {
      const k = labelMap[m2[1].toLowerCase().trim()];
      if (k) {
        for (let j = i + 1; j < Math.min(i + 6, rawLines.length); j++) {
          const val = rawLines[j].trim();
          if (val) { if (!meta[k]) meta[k] = val; break; }
        }
      }
    }
  }
  return meta;
}

function toTitleCase(str) {
  if (!str) return str;
  const small = new Set(['a','an','the','and','but','or','for','nor','in','on','at','to','by','up','as','of','with']);
  return str.toLowerCase().split(' ')
    .map((w, i) => (i === 0 || !small.has(w)) ? w.charAt(0).toUpperCase() + w.slice(1) : w)
    .join(' ');
}

function buildVaultDoc(rawLines, section, idx) {
  const trimmed = rawLines.map(l => l.trim()).filter(Boolean);
  if (trimmed.length === 0) return null;

  const meta = extractVaultMeta(rawLines.slice(0, 80));

  // Determine title
  let title;
  if (section === 'emails') {
    title = meta.subject || toTitleCase(trimmed[0]);
  } else {
    // First non-empty line is title (convert from ALL CAPS if needed)
    const raw = trimmed[0];
    title = raw === raw.toUpperCase() && raw.length > 4 ? toTitleCase(raw) : raw;
  }

  // Find body start — skip past email headers or metadata block
  const bodyMarkers = ['Executive Summary', '⚠', 'Overview', 'Background', 'Key Findings'];
  let bodyStart = 0;
  if (section === 'emails') {
    // Walk past header pairs (label on one line, value on next): From/To/Cc/Date/Subject
    const emailHeaderLabels = new Set(['from:', 'to:', 'cc:', 'date:', 'subject:']);
    let hi = 0;
    while (hi < trimmed.length) {
      const lc = trimmed[hi].toLowerCase();
      if (emailHeaderLabels.has(lc)) {
        const isSubject = lc === 'subject:';
        hi++; // skip label
        if (hi < trimmed.length) hi++; // skip value
        if (isSubject) break; // body starts after Subject value
      } else {
        break;
      }
    }
    bodyStart = hi < trimmed.length ? hi : Math.min(12, trimmed.length - 1);
  } else {
    for (let i = 0; i < trimmed.length; i++) {
      if (bodyMarkers.some(m => trimmed[i].startsWith(m))) { bodyStart = i; break; }
    }
    if (!bodyStart) bodyStart = Math.min(12, trimmed.length - 1);
  }

  const preview = trimmed.slice(bodyStart, bodyStart + 6).join(' ').substring(0, 280);

  return {
    index: idx, section, title, meta,
    preview: preview + (preview.length >= 280 ? '…' : ''),
    isEmail: section === 'emails',
    content: rawLines.join('\n').trim()
  };
}

function parseVaultText(text) {
  const result = { risk: [], regulatory: [], emails: [] };
  const lines   = text.replace(/\r\n/g, '\n').split('\n');

  let currentSection = null;
  let currentDocLines = [];
  let inDoc  = false;
  const docIdx = { risk: 0, regulatory: 0, emails: 0 };

  const pushDoc = () => {
    if (!currentSection || currentDocLines.length === 0) return;
    const doc = buildVaultDoc(currentDocLines, currentSection, docIdx[currentSection]);
    if (doc) { result[currentSection].push(doc); docIdx[currentSection]++; }
    currentDocLines = [];
    inDoc = false;
  };

  for (const rawLine of lines) {
    const line = rawLine.trim();

    // Section headers repeat before every document — switch section only when it changes
    if (/^SECTION\s+0?1$/i.test(line)) {
      if (currentSection !== 'risk') { pushDoc(); currentSection = 'risk'; inDoc = false; }
      continue;
    }
    if (/^SECTION\s+0?2$/i.test(line)) {
      if (currentSection !== 'regulatory') { pushDoc(); currentSection = 'regulatory'; inDoc = false; }
      continue;
    }
    if (/^SECTION\s+0?3$/i.test(line)) {
      if (currentSection !== 'emails') { pushDoc(); currentSection = 'emails'; inDoc = false; }
      continue;
    }
    // Sub-heading lines after section markers — skip (only when not inside a document)
    if (!inDoc && /^RISK ANALYSIS|^REGULATORY$|^EMAIL INTELLIGENCE$/i.test(line)) continue;

    if (currentSection && /^document\s+\d+\s+of\s+\d+/i.test(line)) {
      pushDoc(); inDoc = true;
      const after = line.replace(/^document\s+\d+\s+of\s+\d+:?\s*/i, '').trim();
      if (after) currentDocLines.push(after);
      continue;
    }

    if (inDoc) currentDocLines.push(rawLine);
  }
  pushDoc();
  return result;
}

async function loadVaultData() {
  if (!mammoth) return { error: 'mammoth not installed — run: npm install mammoth' };
  if (!fs.existsSync(VAULT_DIR)) return { error: 'data/vault/ folder not found.' };

  const files = fs.readdirSync(VAULT_DIR).filter(f => /\.docx?$/i.test(f));
  if (!files.length) return { error: 'No Word document found in data/vault/. Place Vault+Report_Page_Data.docx there.' };

  try {
    const fp = path.join(VAULT_DIR, files[0]);
    const { value } = await mammoth.extractRawText({ path: fp });
    const docs = parseVaultText(value);
    return { ok: true, docs, filename: files[0] };
  } catch (err) {
    return { error: `Failed to parse vault document: ${err.message}` };
  }
}

function vaultTemplateSummary(doc) {
  const na = 'Not available in source document.';
  const overview = [doc.title, doc.meta.date && `Date: ${doc.meta.date}`, doc.meta.author && `Author: ${doc.meta.author}`, doc.meta.classification && `Classification: ${doc.meta.classification}`].filter(Boolean).join(' | ');

  if (doc.isEmail || doc.section === 'emails') {
    return [
      `**Document Overview [EXTRACTED]:** ${overview}`,
      `**Main Message [EXTRACTED]:** ${doc.preview || na}`,
      `**Required Actions [EXTRACTED]:** ${na} Review source document for explicit action items.`,
      `**Deadlines [EXTRACTED]:** ${doc.meta.deadline || na}`,
      `**Impacted Customers / Jurisdictions [EXTRACTED]:** ${na}`,
      `**Policy / Regulatory Changes [EXTRACTED]:** ${na}`,
      `**Escalation Requirements [EXTRACTED]:** ${na}`,
      `**Source Traceability [EXTRACTED]:** Source: ${doc.title}. Full document available in Knowledge Vault.`,
    ].join('\n');
  }

  if (doc.section === 'regulatory') {
    return [
      `**Document Overview [EXTRACTED]:** ${overview}`,
      `**Regulation / Change Description [EXTRACTED]:** ${doc.preview || na}`,
      `**Effective Date [EXTRACTED]:** ${na}`,
      `**Impacted Business Areas [EXTRACTED]:** ${na}`,
      `**Capital / Liquidity / Control Impact [INFERRED]:** ${na}`,
      `**Required Implementation Actions [EXTRACTED]:** ${na} Review source document for explicit action items.`,
      `**Deadlines [EXTRACTED]:** ${doc.meta.deadline || na}`,
      `**Regulatory Authority [EXTRACTED]:** ${doc.meta.author || na}`,
      `**Source Traceability [EXTRACTED]:** Source: ${doc.title}. Full document available in Knowledge Vault.`,
    ].join('\n');
  }

  return [
    `**Document Overview [EXTRACTED]:** ${overview}`,
    `**Key Risk Indicators [EXTRACTED]:** ${doc.preview || na}`,
    `**Quantified Exposure [EXTRACTED]:** ${na}`,
    `**High-Risk Jurisdictions [EXTRACTED]:** ${na}`,
    `**High-Risk Entities [EXTRACTED]:** ${na}`,
    `**Typologies Identified [EXTRACTED]:** ${na}`,
    `**Required Actions [EXTRACTED]:** ${na} Review source document for explicit action items.`,
    `**Escalations & Deadlines [EXTRACTED]:** ${doc.meta.deadline || na}`,
    `**Source Traceability [EXTRACTED]:** Source: ${doc.title}. Full document available in Knowledge Vault.`,
  ].join('\n');
}

async function generateVaultSummary(doc) {
  const key = process.env.GROQ_API_KEY;
  if (!key) return vaultTemplateSummary(doc);

  const truncated = doc.content.length > 4000 ? doc.content.substring(0, 4000) + '\n...[content truncated to fit context window]' : doc.content;

  const systemPrompt =
    `You are an enterprise intelligence extraction platform producing governance-safe structured summaries for regulatory and CRO review. ` +
    `Extraction rules: ` +
    `(1) Extract ONLY what is explicitly stated in the source document. Do not infer, expand, or speculate beyond the content. ` +
    `(2) Preserve all numeric values, dates, entity names, and thresholds exactly as they appear. ` +
    `(3) Tag each section with [EXTRACTED] if content is directly stated in the source, or [INFERRED] if synthesised from multiple parts. ` +
    `(4) If information is absent, state exactly: "Not available in source document." ` +
    `(5) Reference source section headings in parentheses where identifiable, e.g., (Source: Section 2.1). ` +
    `Tone rules: Do not say "our analysis", "we recommend", or "leadership should". ` +
    `Do not generate speculative conclusions or unsupported attributions. ` +
    `Do not rewrite paragraphs — extract and structure factual content only. ` +
    `Use quantified intelligence language: "47 transactions exceeded threshold", "14 entities identified", not "many" or "significant". ` +
    `Forbidden: "our analysis", "we recommend", "leadership should", "this demonstrates significant", "criminal exposure", "major risk", "sanctions evasion confirmed".`;

  const isEmail      = doc.isEmail || doc.section === 'emails';
  const isRegulatory = doc.section === 'regulatory';

  const emailPrompt =
    `SOURCE TYPE: Internal Email\n\n` +
    `DOCUMENT CONTENT:\n${truncated}\n\n` +
    `Return ONLY a structured intelligence extraction using EXACTLY these bold labels with confidence tags (one per line):\n\n` +
    `**Document Overview [EXTRACTED]:** [Title | Date | From | To — extract from headers only]\n` +
    `**Main Message [EXTRACTED]:** [Direct extract of the core message — do not paraphrase]\n` +
    `**Required Actions [EXTRACTED]:** [List each explicit action item numbered. If none: "Not available in source document."]\n` +
    `**Deadlines [EXTRACTED]:** [Extract all explicit dates and timelines. If absent: "Not available in source document."]\n` +
    `**Impacted Customers / Jurisdictions [EXTRACTED]:** [Named customers, jurisdictions, countries from document. If absent: "Not available in source document."]\n` +
    `**Policy / Regulatory Changes [EXTRACTED]:** [Explicit policy or regulatory references stated. If absent: "Not available in source document."]\n` +
    `**Escalation Requirements [EXTRACTED]:** [Explicit escalation instructions from document. If absent: "Not available in source document."]\n` +
    `**Source Traceability [EXTRACTED]:** Source: ${doc.title}. Reference section headings where visible in the source content.`;

  const regulatoryPrompt =
    `SOURCE TYPE: Regulatory Document\n\n` +
    `DOCUMENT CONTENT:\n${truncated}\n\n` +
    `Return ONLY a structured intelligence extraction using EXACTLY these bold labels with confidence tags (one per line):\n\n` +
    `**Document Overview [EXTRACTED]:** [Title | Date | Author | Classification — from document headers]\n` +
    `**Regulation / Change Description [EXTRACTED]:** [Name and description of regulation or policy change as stated. No paraphrase.]\n` +
    `**Effective Date [EXTRACTED]:** [Exact date stated. If absent: "Not available in source document."]\n` +
    `**Impacted Business Areas [EXTRACTED]:** [Named business areas, desks, or functions from document. If absent: "Not available in source document."]\n` +
    `**Capital / Liquidity / Control Impact [INFERRED]:** [Quantified impact values if stated; otherwise synthesise from explicit figures. If none: "Not available in source document."]\n` +
    `**Required Implementation Actions [EXTRACTED]:** [Numbered list of explicit required actions. If absent: "Not available in source document."]\n` +
    `**Deadlines [EXTRACTED]:** [Extract all compliance deadlines and dates. If absent: "Not available in source document."]\n` +
    `**Regulatory Authority [EXTRACTED]:** [Named regulatory body, regulator, or issuing authority. If absent: "Not available in source document."]\n` +
    `**Source Traceability [EXTRACTED]:** Source: ${doc.title}. Reference section headings where visible.`;

  const riskPrompt =
    `SOURCE TYPE: Risk Analysis Document\n\n` +
    `DOCUMENT CONTENT:\n${truncated}\n\n` +
    `Return ONLY a structured intelligence extraction using EXACTLY these bold labels with confidence tags (one per line):\n\n` +
    `**Document Overview [EXTRACTED]:** [Title | Date | Author | Classification — from document headers]\n` +
    `**Key Risk Indicators [EXTRACTED]:** [List stated risk indicators using exact language from source. Preserve quantities: "47 transactions", "14 chains", "87 counterparties".]\n` +
    `**Quantified Exposure [EXTRACTED]:** [Extract all numeric exposure values, transaction counts, thresholds exactly as stated. If absent: "Not available in source document."]\n` +
    `**High-Risk Jurisdictions [EXTRACTED]:** [Named jurisdictions, countries, or corridors from source. If absent: "Not available in source document."]\n` +
    `**High-Risk Entities [EXTRACTED]:** [Named counterparties, institutions, or entities from source. Preserve exact names. If absent: "Not available in source document."]\n` +
    `**Typologies Identified [EXTRACTED]:** [Named risk typologies or patterns from source. If absent: "Not available in source document."]\n` +
    `**Required Actions [EXTRACTED]:** [Numbered list of explicit required actions from document. If absent: "Not available in source document."]\n` +
    `**Escalations & Deadlines [EXTRACTED]:** [Explicit escalation requirements and dates. If absent: "Not available in source document."]\n` +
    `**Source Traceability [EXTRACTED]:** Source: ${doc.title}. Reference section headings where visible in the source content.`;

  const userPrompt = isEmail ? emailPrompt : isRegulatory ? regulatoryPrompt : riskPrompt;

  try {
    const res = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user',   content: userPrompt }
        ],
        max_tokens: 1600,
        temperature: 0.1
      },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    return res.data.choices?.[0]?.message?.content || vaultTemplateSummary(doc);
  } catch (err) {
    console.error('Groq vault error:', err.message);
    return vaultTemplateSummary(doc);
  }
}

app.get('/api/vault/documents', async (req, res) => {
  const data = await loadVaultData();
  if (!data.ok) return res.json({ available: false, message: data.error });
  const docs = {};
  for (const sec of ['risk', 'regulatory', 'emails']) {
    docs[sec] = (data.docs[sec] || []).map(({ content, ...rest }) => rest);
  }
  res.json({ available: true, documents: docs, filename: data.filename });
});

app.post('/api/vault/summary', async (req, res) => {
  const { section, index } = req.body || {};
  const data = await loadVaultData();
  if (!data.ok) return res.status(503).json({ error: data.error });
  const doc = data.docs[section]?.[index];
  if (!doc) return res.status(404).json({ error: 'Document not found.' });
  const summary = await generateVaultSummary(doc);
  res.json({ summary, source: data.filename });
});

// ─── Reports Module ───────────────────────────────────────────────────────────

function parseDateStr(s) {
  if (!s) return null;
  const cleaned = s.replace(/,.*$/, '').trim();
  const d1 = new Date(cleaned);
  if (!isNaN(d1.getTime())) return d1;
  const m = cleaned.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/);
  if (m) { const d2 = new Date(`${m[2]} ${m[1]}, ${m[3]}`); if (!isNaN(d2.getTime())) return d2; }
  return null;
}

function filterDocsByDate(docs, fromDate, toDate, includeUndated) {
  // Parse ISO dates (YYYY-MM-DD) as local midnight to match how doc dates are parsed
  const parseISOLocal = s => {
    if (!s) return null;
    const [y, m, d] = s.split('-').map(Number);
    return new Date(y, m - 1, d); // local midnight — avoids UTC/local timezone mismatch
  };
  const dateOnly = d => new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const from = parseISOLocal(fromDate);
  const to   = parseISOLocal(toDate);
  return docs.filter(doc => {
    const d = parseDateStr(doc.meta?.date);
    if (!d) return includeUndated !== false;
    const day = dateOnly(d);
    if (from && day < from) return false;
    if (to   && day > to)   return false;
    return true;
  });
}

function docToText(doc, maxContent = 800) {
  const lines = [];
  lines.push(`TITLE: ${doc.title}`);
  if (doc.meta?.date)           lines.push(`DATE: ${doc.meta.date}`);
  if (doc.meta?.author)         lines.push(`AUTHOR: ${doc.meta.author}`);
  if (doc.meta?.classification) lines.push(`CLASSIFICATION: ${doc.meta.classification}`);
  if (doc.meta?.status)         lines.push(`STATUS: ${doc.meta.status}`);
  lines.push('');
  const content = (doc.content || doc.preview || '').trim();
  lines.push(content.length > maxContent ? content.substring(0, maxContent) + '\n...[truncated]' : content);
  return lines.join('\n');
}

function parseMonthLabel(label) {
  // Handles "Jan-2024", "Feb 2024", "January 2024", etc.
  const s = String(label).trim();
  const m = s.match(/^([A-Za-z]{3,9})[- ](\d{4})$/);
  if (m) { const d = new Date(`${m[1]} 1, ${m[2]}`); if (!isNaN(d.getTime())) return d; }
  return null;
}

function filterMonthlyTrends(data, fromDate, toDate) {
  if ((!fromDate && !toDate) || !data?.chart_data) return data;
  const parseISOLocal = s => { if (!s) return null; const [y,m,d] = s.split('-').map(Number); return new Date(y, m-1, d); };
  const from = parseISOLocal(fromDate);
  const to   = parseISOLocal(toDate);

  const { labels, values } = data.chart_data;
  const kept = labels.reduce((acc, lbl, i) => {
    const d = parseMonthLabel(lbl);
    if (!d) return acc; // summary rows (e.g. "TOTAL / AVG") have no date — exclude when filtering
    const monthStart = new Date(d.getFullYear(), d.getMonth(), 1);
    const monthEnd   = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    if (from && monthEnd   < from) return acc;
    if (to   && monthStart > to)   return acc;
    acc.push(i);
    return acc;
  }, []);

  if (!kept.length) return { ...data, chart_data: { labels: [], values: [] }, dateFiltered: true };
  return {
    ...data,
    chart_data: { labels: kept.map(i => labels[i]), values: kept.map(i => values[i]) },
    dateFiltered: kept.length < labels.length,
  };
}

function loadReportInsightsData(fromDate, toDate) {
  const raw = loadDatasets();
  const result = {};
  try { if (raw['Parties'])                      result.parties          = processParties(raw['Parties']); }          catch {}
  try { if (raw['Risk_Ratings'])                 result.riskRatings      = processRiskRatings(raw['Risk_Ratings']); } catch {}
  try { if (raw['Monthly_TXN_Trends'])           result.monthlyTrends    = filterMonthlyTrends(processMonthlyTrends(raw['Monthly_TXN_Trends']), fromDate, toDate); } catch {}
  try { if (raw['Top10_Accounts_Volume'])        result.topAccounts      = processTopAccounts(raw['Top10_Accounts_Volume']); }        catch {}
  try { if (raw['Top10_Counterparty_Exposures']) result.topCounterparties = processCounterparties(raw['Top10_Counterparty_Exposures']); } catch {}
  return result;
}

function insightsToText(insightsData) {
  if (!insightsData || !Object.keys(insightsData).length) return '';
  const parts = ['INSIGHTS DATA SUMMARY (from AML database):'];
  for (const [, d] of Object.entries(insightsData)) {
    if (!d) continue;
    parts.push(`\n[${d.title}]`);
    for (const [k, v] of Object.entries(d.key_metrics)) parts.push(`  ${k}: ${v}`);
  }
  return parts.join('\n');
}

function parseReportSections(text) {
  const sections = [];
  const parts = text.split(/(?:^|\n)##\s+/);
  for (const part of parts) {
    const trimmed = part.trim();
    if (!trimmed) continue;
    const nl = trimmed.indexOf('\n');
    const title   = nl >= 0 ? trimmed.substring(0, nl).trim() : trimmed;
    const content = nl >= 0 ? trimmed.substring(nl + 1).trim() : '';
    if (title) sections.push({ id: title.toLowerCase().replace(/\W+/g, '_'), title, content });
  }
  return sections;
}

app.get('/api/reports/sources', async (req, res) => {
  const vaultData      = await loadVaultData();
  const insightsRaw    = loadDatasets();
  const insightSheets  = ['Parties','Risk_Ratings','Monthly_TXN_Trends','Top10_Accounts_Volume','Top10_Counterparty_Exposures'].filter(s => insightsRaw[s]);
  res.json({
    risk:       { count: vaultData.ok ? vaultData.docs.risk.length : 0,        available: !!vaultData.ok },
    regulatory: { count: vaultData.ok ? vaultData.docs.regulatory.length : 0,  available: !!vaultData.ok },
    emails:     { count: vaultData.ok ? vaultData.docs.emails.length : 0,      available: !!vaultData.ok },
    insights:   { count: insightSheets.length, available: insightSheets.length > 0 },
  });
});

app.post('/api/reports/generate', async (req, res) => {
  const { selectedSources = {}, fromDate, toDate, includeUndated = true } = req.body || {};
  const contextParts = [];
  let docsIncluded = 0;
  let insightsPayload = null;

  // Vault sources
  if (selectedSources.risk || selectedSources.regulatory || selectedSources.emails) {
    const vaultData = await loadVaultData();
    if (vaultData.ok) {
      const secMap = { risk: 'RISK ANALYSIS DOCUMENTS', regulatory: 'REGULATORY DOCUMENTS', emails: 'EMAIL INTELLIGENCE' };
      for (const sec of ['risk', 'regulatory', 'emails']) {
        if (!selectedSources[sec]) continue;
        const filtered = filterDocsByDate(vaultData.docs[sec], fromDate, toDate, includeUndated);
        if (!filtered.length) continue;
        contextParts.push(`=== ${secMap[sec]} ===`);
        filtered.forEach(d => { contextParts.push(docToText(d)); contextParts.push('---'); });
        docsIncluded += filtered.length;
      }
    }
  }

  // Insights source
  if (selectedSources.insights) {
    const insights = loadReportInsightsData(fromDate, toDate);
    if (Object.keys(insights).length) {
      insightsPayload = insights;
      contextParts.push(insightsToText(insights));
      docsIncluded += Object.keys(insights).length;
    }
  }

  const SECTIONS = ['Executive Summary','Key Risk Themes','Regulatory Developments',
    'Customer & Transaction Insights','Email Intelligence','Required Actions',
    'Deadlines','Ownership & Teams Mentioned','Leadership Attention Required','Supporting Evidence'];

  if (!contextParts.length) {
    return res.json({
      report: { sections: [{ id: 'no_data', title: 'No Data Available', content: 'No source data matched the selected filters and date range. Adjust your selections or date range and try again.' }], metadata: { docsIncluded: 0, generatedAt: new Date().toISOString(), fromDate: fromDate || null, toDate: toDate || null } },
      insightsData: null
    });
  }

  const sourceContext = contextParts.join('\n\n');
  const truncated = sourceContext.length > 6000 ? sourceContext.substring(0, 6000) + '\n\n[Content truncated to fit context window]' : sourceContext;
  const dateRangeText = fromDate || toDate ? `Date Range Covered: ${fromDate || 'All dates'} to ${toDate || 'Present'}` : 'Date Range Covered: All available dates';

  const key = process.env.GROQ_API_KEY;
  if (!key) {
    return res.json({
      report: { sections: SECTIONS.map(s => ({ id: s.toLowerCase().replace(/\W+/g,'_'), title: s, content: 'Groq API key not configured. Add GROQ_API_KEY to your .env file to enable AI report generation.' })), metadata: { generatedAt: new Date().toISOString(), docsIncluded } },
      insightsData: insightsPayload
    });
  }

  const prompt = `You are an expert risk analyst generating an executive intelligence report for senior leadership at Atlas International Bank.

RULES:
1. Use ONLY the source documents provided below. Do not use outside knowledge or invent facts.
2. Do not invent figures, regulation names, party names, or action items not present in the source data.
3. If information is not available in the source data, write exactly: "Not available in selected source data."
4. Be precise, professional, and leadership-ready. Use bullet points starting with • for lists.
5. Each section should be substantive — avoid one-line answers.

${dateRangeText}

SOURCE DATA:
${truncated}

Generate a structured intelligence report with EXACTLY these section headers (use ## prefix):
${SECTIONS.map(s => `## ${s}`).join('\n')}

Begin the report now:`;

  try {
    const response = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      { model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant', messages: [{ role: 'user', content: prompt }], max_tokens: 4096, temperature: 0.25 },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    const text = response.data.choices?.[0]?.message?.content || '';
    const sections = parseReportSections(text);
    res.json({
      report: { sections: sections.length ? sections : [{ id: 'report', title: 'Intelligence Report', content: text }], metadata: { generatedAt: new Date().toISOString(), docsIncluded, fromDate: fromDate || null, toDate: toDate || null } },
      insightsData: insightsPayload
    });
  } catch (err) {
    const detail = err.response?.data ? JSON.stringify(err.response.data) : err.message;
    console.error('Groq report error:', detail);
    res.json({
      report: { sections: [{ id: 'error', title: 'Generation Error', content: `Failed to generate report: ${detail}` }], metadata: { generatedAt: new Date().toISOString(), docsIncluded } },
      insightsData: insightsPayload
    });
  }
});

app.post('/api/reports/email', async (req, res) => {
  const { selectedSources = {}, fromDate, toDate, includeUndated = true } = req.body || {};
  const contextParts = [];

  if (selectedSources.risk || selectedSources.regulatory || selectedSources.emails) {
    const vaultData = await loadVaultData();
    if (vaultData.ok) {
      const labelMap = { risk: 'RISK ANALYSIS', regulatory: 'REGULATORY', emails: 'EMAIL INTELLIGENCE' };
      for (const sec of ['risk', 'regulatory', 'emails']) {
        if (!selectedSources[sec]) continue;
        const filtered = filterDocsByDate(vaultData.docs[sec], fromDate, toDate, includeUndated);
        if (!filtered.length) continue;
        contextParts.push(`=== ${labelMap[sec]} ===`);
        filtered.forEach(d => contextParts.push(`• ${d.title} (${d.meta?.date || 'undated'}): ${d.preview.substring(0, 250)}`));
      }
    }
  }

  if (selectedSources.insights) {
    const insights = loadReportInsightsData();
    if (Object.keys(insights).length) contextParts.push(insightsToText(insights));
  }

  if (!contextParts.length) return res.json({ email: 'No source data matched the selected filters.' });

  const truncated = contextParts.join('\n\n').substring(0, 8000);
  const key = process.env.GROQ_API_KEY;
  if (!key) return res.json({ email: '**SUBJECT:** Risk Intelligence Summary: Key AML, Regulatory and Customer Risk Updates\n\n**Dear Senior Leadership Team,**\n\nGroq API key not configured. Add GROQ_API_KEY to .env to enable AI email generation.\n\nKind regards,\nRisk Intelligence Platform' });

  const prompt = `You are generating a concise, leadership-ready risk intelligence email for senior management at Atlas International Bank.

RULES:
1. Use ONLY the source data provided. Do not invent any facts, figures, names, or deadlines.
2. Keep the tone executive-friendly: clear, direct, action-oriented.
3. If information is missing, omit that bullet rather than fabricating it.

SOURCE DATA:
${truncated}

Generate the email in EXACTLY this format:
**SUBJECT:** [subject line]

**Dear Senior Leadership Team,**

[One sentence context sentence]

**Key Intelligence Updates:**
• [update 1]
• [update 2]
• [update 3]
• [update 4]
• [update 5]

**Required Actions:**
• [action 1]
• [action 2]
• [action 3]

**Upcoming Deadlines:**
• [deadline 1]
• [deadline 2]

**Next Steps:**
[1-2 sentence closing with escalation ask]

Kind regards,
Risk Intelligence Platform`;

  try {
    const response = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      { model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant', messages: [{ role: 'user', content: prompt }], max_tokens: 1200, temperature: 0.25 },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    res.json({ email: response.data.choices?.[0]?.message?.content || '' });
  } catch (err) {
    console.error('Groq email error:', err.message);
    res.json({ email: `Failed to generate email: ${err.message}` });
  }
});

// ─── Scenario Intelligence Module ────────────────────────────────────────────

const SCENARIO_DIR = path.join(__dirname, 'data', 'scenario');

function loadScenarioChartData() {
  if (!fs.existsSync(SCENARIO_DIR)) return {};
  const files = fs.readdirSync(SCENARIO_DIR).filter(f => /\.xlsx?$/i.test(f));
  if (!files.length) return {};
  const result = {};
  try {
    const wb = XLSX.readFile(path.join(SCENARIO_DIR, files[0]));
    console.log('Scenario Excel sheets:', wb.SheetNames);
    for (const sheet of wb.SheetNames) {
      const ws  = wb.Sheets[sheet];
      const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      // Find first non-empty row
      const firstNonEmpty = raw.find(r => r.some(c => c !== '')) || [];

      // Decide if first non-empty row is a header row (contains at least one non-numeric string cell)
      // or a pure data row (all cells are numbers or empty)
      const nonEmptyCells = firstNonEmpty.filter(c => c !== '');
      const stringCount   = nonEmptyCells.filter(c => typeof c === 'string' && isNaN(Number(c))).length;
      const hasHeaders    = stringCount >= 1;

      let rows;
      if (hasHeaders) {
        // Skip to that row as the header
        const headerIdx = raw.indexOf(firstNonEmpty);
        rows = XLSX.utils.sheet_to_json(ws, { defval: null, range: headerIdx });
        rows = rows.map(row => { const r = {}; for (const [k, v] of Object.entries(row)) r[String(k).trim()] = v; return r; });
      } else {
        // No headers — read raw arrays and assign generic column names
        const dataRows = raw.filter(r => r.some(c => c !== ''));
        const maxCols  = Math.max(...dataRows.map(r => r.length));
        const colNames = Array.from({ length: maxCols }, (_, i) => i === 0 ? 'Label' : `Value${i}`);
        rows = dataRows.map(r => {
          const obj = {};
          colNames.forEach((name, i) => { obj[name] = r[i] !== undefined ? r[i] : null; });
          return obj;
        });
      }

      result[sheet] = rows.filter(r => Object.values(r).some(v => v !== null && v !== '' && v !== undefined));
    }
  } catch (err) {
    console.error('Scenario Excel error:', err.message);
  }
  return result;
}

async function loadScenarioDocument() {
  if (!mammoth) return null;
  if (!fs.existsSync(SCENARIO_DIR)) return null;
  const files = fs.readdirSync(SCENARIO_DIR).filter(f => /\.docx?$/i.test(f));
  if (!files.length) return null;
  try {
    const { value } = await mammoth.extractRawText({ path: path.join(SCENARIO_DIR, files[0]) });
    return value;
  } catch (err) {
    console.error('Scenario doc error:', err.message);
    return null;
  }
}

function extractScenarioExecutiveSummary(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);

  // Pass 1: look for explicit "Executive Summary" heading
  let collecting = false;
  const collected = [];
  for (const line of lines) {
    if (/executive\s+summary/i.test(line)) { collecting = true; continue; }
    if (collecting) {
      if (/^(first.?order|second.?order|scope|model.?limit|regulatory|action|section\s+\d|table\s+\d|\d+\.\s+[A-Z])/i.test(line) && collected.length > 2) break;
      if (line.length > 5) collected.push(line);
      if (collected.length >= 20) break;
    }
  }
  if (collected.length) return collected.join(' ').trim();

  // Pass 2: return the first few substantial paragraphs (>50 chars) as the summary
  const paras = lines.filter(l => l.length > 50).slice(0, 8);
  return paras.join(' ').trim();
}

function extractQuantificationValues(text) {
  const result = {};
  const METRICS = {
    'Trading VaR Breach':   /(?:trading\s+)?var\s+breach[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
    'PVBP Spike':           /pvbp\s+(?:spike|increase|change)[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
    'CS01 Widening':        /cs01\s+(?:widening|spike|change)[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
    'FX Net Short':         /fx\s+net\s+short[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
    'Margin Call Volumes':  /margin\s+call\s+(?:volumes?|amount)[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
    'Asset Liquidity':      /asset\s+liquidity[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
    'Survival Horizon':     /survival\s+horizon[:\s]+([£$€\d.,\s%MBKmbn]+)/i,
  };
  for (const [key, regex] of Object.entries(METRICS)) {
    const m = text.match(regex);
    const val = m ? m[1].trim().replace(/\s+/g, ' ') : null;
    // Only include if it doesn't look like a placeholder (xx, yy, $z, n/a)
    result[key] = (val && !/^(xx|yy|zz|\$[xyz]|n\/a|tbc|tbd|\?+)$/i.test(val)) ? val : null;
  }
  return result;
}

app.get('/api/scenario/data', async (req, res) => {
  const docFiles   = fs.existsSync(SCENARIO_DIR) ? fs.readdirSync(SCENARIO_DIR).filter(f => /\.docx?$/i.test(f)) : [];
  const excelFiles = fs.existsSync(SCENARIO_DIR) ? fs.readdirSync(SCENARIO_DIR).filter(f => /\.xlsx?$/i.test(f)) : [];

  if (!docFiles.length && !excelFiles.length) {
    return res.json({ available: false, message: 'Scenario files not found. Please place Scenario_Analysis_Module.docx and Scenario_Analysis_ChartData.xlsx in /data/scenario/.' });
  }

  const text = await loadScenarioDocument();
  const executiveSummary   = text ? extractScenarioExecutiveSummary(text) : '';
  const quantification     = text ? extractQuantificationValues(text)     : {};
  res.json({ available: true, executiveSummary, quantification, hasDoc: !!docFiles.length, hasExcel: !!excelFiles.length });
});

app.get('/api/scenario/chart-data', (req, res) => {
  const excelFiles = fs.existsSync(SCENARIO_DIR) ? fs.readdirSync(SCENARIO_DIR).filter(f => /\.xlsx?$/i.test(f)) : [];
  if (!excelFiles.length) {
    return res.json({ available: false, message: 'Scenario chart data not found. Please place Scenario_Analysis_ChartData.xlsx in /data/scenario/.' });
  }

  const raw = loadScenarioChartData();
  if (!Object.keys(raw).length) {
    return res.json({ available: false, message: 'Chart data not available in uploaded Excel file.' });
  }

  const SHEET_CFG = {
    '1_VaR_Breach_Timeline':  { type: 'line',          title: 'Trading VaR Escalation' },
    '2_Unhedged_Sensitivity': { type: 'bar',            title: 'Unhedged Sensitivity Build-Up' },
    '3_Liquidity_Impact':     { type: 'bar',            title: 'Liquidity / LCR / Survival Horizon Impact' },
    '4_PnL_Impact':           { type: 'bar',            title: 'Daily P&L Impact by Desk' },
    '5_RWA_Capital_Impact':   { type: 'bar',            title: 'RWA and Capital Impact' },
    '6_Margin_Collateral':    { type: 'horizontalBar',  title: 'Trapped Margin and Collateral Impact' },
    '7_Risk_Stripe_Heatmap':  { type: 'heatmap',        title: 'Risk Stripe Severity Heatmap' },
  };

  const charts = {};
  for (const [sheet, rows] of Object.entries(raw)) {
    if (sheet === '0_Chart_Guide') continue;
    const cfg  = SHEET_CFG[sheet];
    if (!cfg) continue;
    if (!rows.length) { charts[sheet] = { ...cfg, available: false }; continue; }

    const cols = Object.keys(rows[0]);

    if (cfg.type === 'heatmap') {
      charts[sheet] = { ...cfg, available: true, rows, columns: cols };
      continue;
    }

    const labelCol  = cols[0];
    const valueCols = cols.slice(1).filter(c =>
      rows.some(r => r[c] !== null && r[c] !== undefined && !isNaN(parseFloat(r[c])))
    );

    if (!valueCols.length) { charts[sheet] = { ...cfg, available: false }; continue; }

    if (valueCols.length === 1) {
      charts[sheet] = {
        ...cfg, available: true,
        chart_data: {
          labels: rows.map(r => String(r[labelCol] ?? '')),
          values: rows.map(r => parseFloat(r[valueCols[0]]) || 0)
        }
      };
    } else {
      charts[sheet] = {
        ...cfg, available: true, multi: true,
        chart_data: {
          labels:   rows.map(r => String(r[labelCol] ?? '')),
          datasets: valueCols.map(col => ({
            label:  col,
            values: rows.map(r => parseFloat(r[col]) || 0)
          }))
        }
      };
    }
  }

  res.json({ available: true, charts });
});

app.post('/api/scenario/briefing', async (req, res) => {
  const { type = 'cro' } = req.body || {};
  const key = process.env.GROQ_API_KEY;
  if (!key) return res.json({ briefing: 'Groq API key is not configured. Please add GROQ_API_KEY to environment variables.' });

  const text = await loadScenarioDocument();
  if (!text) return res.json({ briefing: 'Scenario document not found. Please place Scenario_Analysis_Module.docx in /data/scenario/.' });

  const truncated = text.length > 5000 ? text.substring(0, 5000) + '\n[Content truncated]' : text;

  const FORMATS = {
    cro: `Generate a CRO Briefing with these numbered sections:\n1. Situation Overview\n2. Why This Matters\n3. Risk Stripes Activated\n4. Financial / Capital Impact\n5. Liquidity / Survival Horizon Considerations\n6. Regulatory Notifications\n7. Immediate Decisions Required\n8. Recommended Management Actions\n9. Open Data Gaps`,
    board: `Generate a Board Update with these numbered sections:\n1. Event Summary\n2. Key Risk Concerns\n3. Financial Impact\n4. Required Board Awareness\n5. Decisions / Approvals Required`,
    regulator: `Generate a Regulator Update with these numbered sections:\n1. Incident Description\n2. Timeline\n3. Bank Impact\n4. Client / Market Impact\n5. Controls Activated\n6. Further Updates`,
    actions: `Generate a concise executive Action Playbook Summary covering:\n1. Critical Immediate Priorities (0–4 Hours)\n2. Key Short-Term Actions (4–24 Hours)\n3. Ongoing Management Requirements\n4. Ownership and Governance\n5. Escalation Triggers\nFocus on the most critical decisions required from CRO leadership. Be direct and action-oriented.`
  };

  const prompt =
    `You are an executive risk intelligence assistant for a CRO dashboard.\n` +
    `Use ONLY the source data provided. Do not use external knowledge or invent facts, figures, legal obligations, or exposures.\n` +
    `If information is missing, state: "Not available in source document."\n` +
    `Write in concise senior leadership language. Keep the tone board-ready, factual, and risk-focused.\n\n` +
    `SCENARIO: CME Exchange Failure — Closure of an Exchange due to Cyber Attack\n\n` +
    `SOURCE DOCUMENT:\n${truncated}\n\n` +
    (FORMATS[type] || FORMATS.cro);

  try {
    const response = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      { model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant', messages: [{ role: 'user', content: prompt }], max_tokens: 1500, temperature: 0.2 },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    res.json({ briefing: response.data.choices?.[0]?.message?.content || 'Unable to generate briefing.' });
  } catch (err) {
    console.error('Scenario briefing error:', err.message);
    res.json({ briefing: `Failed to generate briefing: ${err.message}` });
  }
});

// Serve React app for all non-API routes
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─── Start ────────────────────────────────────────────────────────────────────

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log('\n  ╔══════════════════════════════════════════╗');
  console.log('  ║   Risk Intelligence Platform             ║');
  console.log('  ╚══════════════════════════════════════════╝');
  console.log(`\n  Running at : http://localhost:${PORT}`);
  console.log(`  Data folder: ${DATA_DIR}`);
  console.log(`  Groq API   : ${process.env.GROQ_API_KEY ? '✓ Configured' : '✗ Not set (using template summaries)'}`);
  console.log('\n  Place Dataset.xlsx inside data/insights/ to begin.\n');
});
