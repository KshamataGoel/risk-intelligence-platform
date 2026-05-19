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
        const headerRow = (raw.length >= 2 && cnt1 > cnt0) ? 1 : 0;

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

  // Strip summary/total rows
  const data = rows.filter(r => {
    const t = String(r[tCol] ?? '').toLowerCase().trim();
    return t && t !== 'total' && t !== 'grand total' && t !== 'subtotal';
  });

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
        'Highest Risk Type':  ranked[0]?.type  ?? 'N/A',
        'Top Risk Score':     ranked[0]?.score ?? 'N/A',
        'Party Types Ranked': ranked.length
      },
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

  const metricLabel = hlType === 'min' ? 'Smallest Segment'
                    : hlType === 'mention' ? 'Highlighted Segment'
                    : 'Largest Segment';

  return {
    answer_available: true,
    dataset_used: 'Parties',
    chart_type: opts.chartType || 'bar',
    title: opts.chartType === 'doughnut'
      ? 'Party Type Distribution'
      : 'Party Count by Type',
    chart_data:      { labels, values },
    highlight_index: hlIdx,
    highlight_type:  hlType,
    key_metrics: {
      'Total Parties':              total,
      'Party Types':                labels.length,
      [metricLabel]:                labels[hlIdx],
      [`${metricLabel} Count`]:     values[hlIdx]
    },
    columns: cols,
    preview: data.slice(0, 10)
  };
}

function processRiskRatings(rows, opts = {}) {
  if (!rows.length) throw new Error('Risk_Ratings sheet is empty');
  const cols  = Object.keys(rows[0]);
  const rCol  = cols.find(c => /risk.*rating|rating/i.test(c)) || cols[0];
  const cCol  = cols.find(c => /count|customer/i.test(c));

  // Strip summary/total rows
  const data = rows.filter(r => {
    const t = String(r[rCol] ?? '').toLowerCase().trim();
    return t && t !== 'total' && t !== 'grand total' && t !== 'subtotal';
  });

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

  const tierLabel = hlType === 'min' ? 'Least Common Tier'
                  : hlType === 'mention' ? 'Highlighted Tier'
                  : 'Most Common Tier';

  return {
    answer_available: true,
    dataset_used: 'Risk_Ratings',
    chart_type: 'doughnut',
    title: 'Risk Rating Distribution',
    chart_data:      { labels, values },
    highlight_index: hlIdx,
    highlight_type:  hlType,
    key_metrics: {
      'Total Customers':      total,
      'Risk Categories':      labels.length,
      [tierLabel]:            labels[hlIdx],
      [`${tierLabel} Count`]: values[hlIdx]
    },
    columns: cols,
    preview: data.slice(0, 10)
  };
}

function processTopAccounts(rows, opts = {}) {
  if (!rows.length) throw new Error('Top10_Accounts_Volume sheet is empty');
  const cols   = Object.keys(rows[0]);
  // Prefer "Party Name" / "Account Name" over plain "Party ID" or "Party"
  const nCol   = cols.find(c => /party.?name|account.?name|\bname\b/i.test(c))
              || cols.find(c => /name|party|account/i.test(c))
              || cols[0];
  const vCol   = cols.find(c => /volume|total|amount|value|txn/i.test(c)) || cols[cols.length - 1];

  const data = rows
    .map(r => ({ name: String(r[nCol] ?? 'Unknown'), volume: parseGBP(r[vCol]) ?? 0 }))
    .sort((a, b) => b.volume - a.volume)
    .slice(0, 10);

  const labels = data.map(d => d.name);
  const values = data.map(d => d.volume);
  const { idx: hlIdx, type: hlType } = resolveHighlight(opts.question || '', labels, values);

  return {
    answer_available: true,
    dataset_used: 'Top10_Accounts_Volume',
    chart_type: 'horizontalBar',
    title: 'Top 10 Accounts by Transaction Volume',
    chart_data:      { labels, values },
    highlight_index: hlIdx,
    highlight_type:  hlType,
    key_metrics: {
      'Highlighted Account': labels[hlIdx] ?? 'N/A',
      'Volume':              `£${(values[hlIdx] || 0).toLocaleString()}`,
      'Top 10 Total':        `£${values.reduce((s, v) => s + v, 0).toLocaleString()}`
    },
    columns: cols,
    preview: rows.slice(0, 10)
  };
}

function processMonthlyTrends(rows) {
  if (!rows.length) throw new Error('Monthly_TXN_Trends sheet is empty');
  const cols  = Object.keys(rows[0]);
  const mCol  = cols.find(c => /month|date|period/i.test(c)) || cols[0];
  const vCol  = cols.find(c => /volume|total|amount|value|txn/i.test(c)) || cols[cols.length - 1];

  const data  = rows.map(r => ({ month: String(r[mCol] ?? ''), volume: parseGBP(r[vCol]) ?? 0 }));
  const total = data.reduce((s, d) => s + d.volume, 0);
  const avg   = total / (data.length || 1);
  const max   = Math.max(...data.map(d => d.volume));
  const peak  = data.find(d => d.volume === max);

  return {
    answer_available: true,
    dataset_used: 'Monthly_TXN_Trends',
    chart_type: 'line',
    title: 'Monthly Transaction Volume Trends',
    chart_data: { labels: data.map(d => d.month), values: data.map(d => d.volume) },
    key_metrics: {
      'Total Volume':    `£${total.toLocaleString()}`,
      'Monthly Average': `£${Math.round(avg).toLocaleString()}`,
      'Peak Month':      peak?.month ?? 'N/A'
    },
    columns: cols,
    preview: rows.slice(0, 10)
  };
}

function processCounterparties(rows, opts = {}) {
  if (!rows.length) throw new Error('Top10_Counterparty_Exposures sheet is empty');
  const cols  = Object.keys(rows[0]);
  const nCol  = cols.find(c => /counterpart|name|party/i.test(c)) || cols[0];
  const eCol  = cols.find(c => /exposure|gross|amount|value/i.test(c)) || cols[cols.length - 1];

  const data = rows
    .map(r => ({ name: String(r[nCol] ?? 'Unknown'), exposure: parseGBP(r[eCol]) ?? 0 }))
    .sort((a, b) => b.exposure - a.exposure)
    .slice(0, 10);

  const labels = data.map(d => d.name);
  const values = data.map(d => d.exposure);
  const { idx: hlIdx, type: hlType } = resolveHighlight(opts.question || '', labels, values);

  return {
    answer_available: true,
    dataset_used: 'Top10_Counterparty_Exposures',
    chart_type: 'horizontalBar',
    title: 'Top 10 Counterparty Exposures',
    chart_data:      { labels, values },
    highlight_index: hlIdx,
    highlight_type:  hlType,
    key_metrics: {
      'Highlighted Counterparty': labels[hlIdx] ?? 'N/A',
      'Exposure':                 `£${(values[hlIdx] || 0).toLocaleString()}`,
      'Total Exposure':           `£${values.reduce((s, v) => s + v, 0).toLocaleString()}`
    },
    columns: cols,
    preview: rows.slice(0, 10)
  };
}

// ─── Summaries ────────────────────────────────────────────────────────────────

function templateSummary(result) {
  const { dataset_used: ds, key_metrics: m } = result;
  const T = {
    Parties:
      `The database contains ${m['Total Parties']} registered parties across ${m['Party Types']} distinct types. ` +
      `The ${m['Largest Segment']} segment represents the largest concentration. ` +
      `Leadership should review segment distribution to assess portfolio diversification and onboarding patterns.`,
    Risk_Ratings:
      `The risk profile covers ${m['Total Customers']} customers distributed across ${m['Risk Categories']} rating categories. ` +
      `The most prevalent risk tier is ${m['Most Common Tier']}. ` +
      `Senior leadership should validate elevated-risk concentrations ahead of the next risk committee briefing.`,
    Top10_Accounts_Volume:
      `The leading account by transaction volume is ${m['Top Account']} at ${m['Highest Volume']}. ` +
      `Combined top-10 volume stands at ${m['Top 10 Total']}, representing a material concentration. ` +
      `Concentration risk thresholds for the highest-volume accounts warrant leadership review.`,
    Monthly_TXN_Trends:
      `Total transaction volume over the reporting period reached ${m['Total Volume']}, ` +
      `with a monthly average of ${m['Monthly Average']}. Volume peaked in ${m['Peak Month']}. ` +
      `These patterns should inform liquidity planning and capacity forecasting for the next quarter.`,
    Top10_Counterparty_Exposures:
      `The largest counterparty exposure is with ${m['Largest Counterparty']} at ${m['Highest Exposure']}. ` +
      `Aggregate top-10 exposure totals ${m['Total Exposure']}. ` +
      `A review of counterparty limits and concentration thresholds is recommended for the next governance cycle.`
  };
  return T[ds] || `Analysis complete for ${ds}. ${Object.entries(m).map(([k, v]) => `${k}: ${v}`).join('. ')}.`;
}

async function groqSummary(question, result) {
  const key = process.env.GROQ_API_KEY;
  if (!key) return templateSummary(result);

  // Build a data table string so the AI uses real names/values, not hallucinated ones
  const dataRows = (result.chart_data?.labels || [])
    .map((lbl, i) => `  ${lbl}: ${result.chart_data.values[i]}`)
    .join('\n');

  const prompt =
    `You are a senior Risk & Compliance analytics assistant briefing CRO leadership.\n` +
    `CRITICAL: Use ONLY the exact names and numbers from the data below. Do NOT invent names like "Market A" or "Account X".\n` +
    `Use precise, executive-level language. Maximum 4 sentences.\n\n` +
    `User question: ${question}\n` +
    `Dataset: ${result.dataset_used}\n` +
    `Key metrics: ${JSON.stringify(result.key_metrics)}\n` +
    (dataRows ? `Data breakdown:\n${dataRows}\n` : '') +
    `\nStructure:\n1. Direct answer using actual names from the data\n2. Two key observations\n3. Recommended leadership action`;

  try {
    const res = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 320,
        temperature: 0.2
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
  const topRisk = Object.entries(s.risk_distribution).sort((a, b) => b[1] - a[1])[0];
  const colEntries = Object.entries(s.column_averages || {}).slice(0, 5);
  const colSummary = colEntries.map(([k, v]) => {
    const isGBP = k.toLowerCase().includes('gbp') || k.toLowerCase().includes('amount');
    return `${k}: ${isGBP ? '£' + Number(v).toLocaleString() : v}`;
  }).join(', ');
  return {
    label: s.label,
    description: `This cluster contains ${s.parties} parties. ` +
      (colSummary ? `Feature averages — ${colSummary}. ` : '') +
      `The predominant risk rating is ${topRisk?.[0] || 'Unknown'}.`
  };
}

async function generateClusterDescriptions(stats) {
  const key = process.env.GROQ_API_KEY;
  if (!key) return stats.map(templateClusterDesc);

  const prompt =
    `You are a senior AML analytics assistant preparing cluster summaries for CRO-level leadership.\n` +
    `Use ONLY the provided cluster summary data. Do not assume anything outside the data.\n` +
    `Do not invent regulatory facts, customer names, or external context.\n\n` +
    `For each cluster provide:\n1. A business-friendly description (what the cluster represents)\n` +
    `2. Key risk indicators visible in the data\nKeep each summary under 80 words.\n\n` +
    `Cluster summary:\n${JSON.stringify(stats, null, 2)}\n\n` +
    `Return a JSON array ONLY: [{"label": "...", "description": "..."}]`;

  try {
    const res = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      { model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant', messages: [{ role: 'user', content: prompt }], max_tokens: 700, temperature: 0.2 },
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
  if (!key) return res.json({ summary: 'Groq API key not configured.' });

  const prompt =
    `You are a senior AML Risk & Compliance analyst preparing a concise executive insight summary for CRO leadership.\n` +
    `Use ONLY the data provided below. Do not invent facts or figures.\n` +
    `Write 3-4 sentences covering: key cluster patterns, notable anomaly risk, and one recommended leadership action.\n\n` +
    `Cluster Descriptions:\n${JSON.stringify(clusterDescriptions, null, 2)}\n\n` +
    `Anomaly Summary:\n${JSON.stringify(anomalySummary, null, 2)}`;

  try {
    const response = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 280,
        temperature: 0.2
      },
      { headers: { Authorization: `Bearer ${key}`, 'Content-Type': 'application/json' } }
    );
    res.json({ summary: response.data.choices?.[0]?.message?.content || 'Summary not available.' });
  } catch (err) {
    console.error('Insights summary error:', err.message);
    res.json({ summary: 'Summary generation failed.' });
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
  if (doc.isEmail) {
    return [
      `**Sender:** ${doc.meta.author || 'Not available in source document.'}`,
      `**Recipients:** ${doc.meta.to || 'Not available in source document.'}`,
      `**Subject:** ${doc.meta.subject || doc.title}`,
      `**Date:** ${doc.meta.date || 'Not available in source document.'}`,
      `**Main Message:** ${doc.preview || 'Not available in source document.'}`,
      `**Required Actions:** Review document for specific action items.`,
      `**Deadline:** ${doc.meta.deadline || 'Not available in source document.'}`,
      `**Impacted Customers / Jurisdictions:** Not available in source document.`,
      `**Escalation Required:** Not available in source document.`,
    ].join('\n');
  }
  return [
    `**Executive Summary:** ${doc.preview || 'Not available in source document.'}`,
    `**Key Findings:** Review document for detailed findings.`,
    `**Risk / Regulatory Impact:** Not available in source document.`,
    `**Required Actions:** Review document for action items.`,
    `**Deadlines:** ${doc.meta.deadline || 'Not available in source document.'}`,
    `**Owners / Teams Mentioned:** ${doc.meta.author || doc.meta.owner || 'Not available in source document.'}`,
    `**Leadership Attention Required:** Please review the full document for leadership actions.`,
  ].join('\n');
}

async function generateVaultSummary(doc) {
  const key = process.env.GROQ_API_KEY;
  if (!key) return vaultTemplateSummary(doc);

  const emailPrompt =
    `You are a senior Compliance analyst summarizing an internal email for CRO leadership.\n` +
    `Use ONLY the content below. Do not invent any information.\n` +
    `If a field is absent, say "Not available in source document."\n\n` +
    `EMAIL CONTENT:\n${doc.content}\n\n` +
    `Return ONLY a structured summary using exactly these bold labels (one per line):\n` +
    `**Sender:**\n**Recipients:**\n**Subject:**\n**Date:**\n**Main Message:**\n` +
    `**Required Actions:**\n**Deadline:**\n**Impacted Customers / Jurisdictions:**\n**Escalation Required:**`;

  const docPrompt =
    `You are a senior Risk & Compliance analyst preparing an executive brief for CRO leadership.\n` +
    `Use ONLY the document content below. Do not invent any information.\n` +
    `If a field is absent, say "Not available in source document."\n\n` +
    `DOCUMENT CONTENT:\n${doc.content}\n\n` +
    `Return ONLY a structured summary using exactly these bold labels (one per line):\n` +
    `**Executive Summary:**\n**Key Findings:**\n**Risk / Regulatory Impact:**\n` +
    `**Required Actions:**\n**Deadlines:**\n**Owners / Teams Mentioned:**\n**Leadership Attention Required:**`;

  try {
    const res = await axios.post(
      'https://api.groq.com/openai/v1/chat/completions',
      {
        model: process.env.GROQ_MODEL || 'llama-3.1-8b-instant',
        messages: [{ role: 'user', content: doc.isEmail ? emailPrompt : docPrompt }],
        max_tokens: 1400,
        temperature: 0.15
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
