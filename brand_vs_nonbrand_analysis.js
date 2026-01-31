/**
 * Brand vs Non-Brand Google Ads Analysis Script
 *
 * Analyzes branded vs non-branded search terms across Search, Performance Max,
 * and Shopping campaigns. Outputs a Google Sheet with raw data tabs and six
 * metric charts (impressions, clicks, cost, conversions, conversion value,
 * cost per conversion) over time, with configurable date range and time
 * granularity (month or week).
 *
 * Pmax: Rows with status EXCLUDED/ADDED_EXCLUDED are skipped. The API often still returns
 * branded terms as ADDED/NONE when excluded in the UI; use PMAX_TREAT_ALL_AS_NON_BRANDED
 * to report all Pmax volume as non-branded when you have excluded branded in the UI.
 *
 * Required OAuth Scopes:
 * - https://www.googleapis.com/auth/spreadsheets
 *
 * Usage:
 * 1. Configure SHEET_URL, date range, BRAND_TOKENS, TIME_GRANULARITY, INCLUDE_BY_CAMPAIGN_TYPE
 * 2. Run the script manually or schedule it
 * 3. Check the output in the specified Google Sheet
 */

// ===== CONFIGURATION =====
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/1ysIP5kMNCuBI5uVzf6V-Xw6knCCSGMCEXw1B5i8zURY/edit?usp=sharing';

// Date range: use lookback OR explicit start/end (yyyy-MM-dd)
const LOOKBACK_DAYS = 90;
const START_DATE = ''; // e.g. '2024-01-01'
const END_DATE = '';   // e.g. '2024-12-31'

// Use only brand-specific phrases (e.g. "foodsisters"). Avoid single common words like "sister"
// or they will match generic queries (e.g. "book for sister") and over-count branded.
const BRAND_TOKENS = [
  'foodsisters', 'foodsister'
];

const TIME_GRANULARITY = 'month'; // 'month' (default) or 'week'
const INCLUDE_BY_CAMPAIGN_TYPE = true; // If true, add Raw + Charts tabs for Search, Pmax, Shopping

// Pmax: the API returns branded terms as ADDED/NONE even when excluded in the UI.
// Set true to report all Pmax search-term volume as non-branded (Pmax branded = 0).
// Set false to classify Pmax terms using BRAND_TOKENS (useful for historical comparison).
const PMAX_TREAT_ALL_AS_NON_BRANDED = false;

// Pmax Consumer Spotlight: Add tabs showing category-level data from campaign_search_term_insight.
// This shows search categories (themes) instead of individual terms. No cost metrics available.
const INCLUDE_PMAX_CATEGORIES = true;

// ===== HELPERS =====

function getSheetId(sheetIdentifier) {
  if (!sheetIdentifier || !sheetIdentifier.includes('/')) {
    return sheetIdentifier;
  }
  const match = sheetIdentifier.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (!match) {
    throw new Error('Invalid Google Sheet URL or ID format');
  }
  return match[1];
}

function getDateRangeClause() {
  const timeZone = AdsApp.currentAccount().getTimeZone();
  let startDate;
  let endDate;
  if (START_DATE && END_DATE) {
    startDate = new Date(START_DATE);
    endDate = new Date(END_DATE);
  } else {
    endDate = new Date();
    startDate = new Date();
    startDate.setDate(endDate.getDate() - LOOKBACK_DAYS);
  }
  const fStart = Utilities.formatDate(startDate, timeZone, 'yyyy-MM-dd');
  const fEnd = Utilities.formatDate(endDate, timeZone, 'yyyy-MM-dd');
  return "segments.date BETWEEN '" + fStart + "' AND '" + fEnd + "'";
}

function buildBrandPatterns() {
  const pattern = new RegExp(
    BRAND_TOKENS
      .map(function (token) {
        return token
          .replace(/[- ]/g, '[-_ ]?')
          .replace(/([0-9])/g, ' ?$1');
      })
      .join('|'),
    'i'
  );
  return [pattern];
}

let BRAND_PATTERNS = [];

function isBranded(text) {
  if (!text || typeof text !== 'string') return false;
  return BRAND_PATTERNS.some(function (p) { return p.test(text); });
}

function emptyMetrics() {
  return {
    impressions: 0,
    clicks: 0,
    cost: 0,
    conversions: 0,
    conversionsValue: 0
  };
}

function emptyPeriodData() {
  return {
    branded: emptyMetrics(),
    nonBranded: emptyMetrics()
  };
}

// For Pmax Categories - includes blank category for unidentifiable search categories
function emptyPeriodDataWithBlank() {
  return {
    branded: emptyMetrics(),
    nonBranded: emptyMetrics(),
    blank: emptyMetrics()
  };
}

function addRowToPeriodData(periodData, periodKey, isBrandedRow, rowMetrics) {
  if (!periodData[periodKey]) {
    periodData[periodKey] = emptyPeriodData();
  }
  const target = isBrandedRow ? periodData[periodKey].branded : periodData[periodKey].nonBranded;
  target.impressions += rowMetrics.impressions;
  target.clicks += rowMetrics.clicks;
  target.cost += rowMetrics.cost;
  target.conversions += rowMetrics.conversions;
  target.conversionsValue += rowMetrics.conversionsValue;
}

function addRowToPeriodDataWithBlank(periodData, periodKey, category, rowMetrics) {
  // category: 'branded', 'nonBranded', or 'blank'
  if (!periodData[periodKey]) {
    periodData[periodKey] = emptyPeriodDataWithBlank();
  }
  const target = periodData[periodKey][category];
  target.impressions += rowMetrics.impressions;
  target.clicks += rowMetrics.clicks;
  target.cost += rowMetrics.cost;
  target.conversions += rowMetrics.conversions;
  target.conversionsValue += rowMetrics.conversionsValue;
}

// ===== DATA FETCHES =====

function processSearchTermView(channelType) {
  const dateClause = getDateRangeClause();
  const timeSegment = TIME_GRANULARITY === 'week' ? 'segments.week' : 'segments.month';
  const query = [
    'SELECT',
    '  search_term_view.search_term,',
    '  ' + timeSegment + ',',
    '  metrics.impressions,',
    '  metrics.clicks,',
    '  metrics.cost_micros,',
    '  metrics.conversions,',
    '  metrics.conversions_value',
    'FROM search_term_view',
    "WHERE " + dateClause,
    "  AND campaign.advertising_channel_type = '" + channelType + "'"
  ].join('\n');

  const totals = { branded: emptyMetrics(), nonBranded: emptyMetrics() };
  const periodData = {};

  const report = AdsApp.search(query);
  while (report.hasNext()) {
    try {
      const row = report.next();
      const text = row.searchTermView && row.searchTermView.searchTerm ? row.searchTermView.searchTerm : '';
      const m = row.metrics || {};
      const cost = Number(m.costMicros) || 0;
      const costActual = cost / 1000000;
      const rowMetrics = {
        impressions: Number(m.impressions) || 0,
        clicks: Number(m.clicks) || 0,
        cost: costActual,
        conversions: Number(m.conversions) || 0,
        conversionsValue: Number(m.conversionsValue) || 0
      };
      const periodKey = TIME_GRANULARITY === 'week' ? (row.segments && row.segments.week) : (row.segments && row.segments.month);
      if (!periodKey) continue;

      const branded = isBranded(text);
      if (branded) {
        totals.branded.impressions += rowMetrics.impressions;
        totals.branded.clicks += rowMetrics.clicks;
        totals.branded.cost += rowMetrics.cost;
        totals.branded.conversions += rowMetrics.conversions;
        totals.branded.conversionsValue += rowMetrics.conversionsValue;
      } else {
        totals.nonBranded.impressions += rowMetrics.impressions;
        totals.nonBranded.clicks += rowMetrics.clicks;
        totals.nonBranded.cost += rowMetrics.cost;
        totals.nonBranded.conversions += rowMetrics.conversions;
        totals.nonBranded.conversionsValue += rowMetrics.conversionsValue;
      }
      addRowToPeriodData(periodData, periodKey, branded, rowMetrics);
    } catch (e) {
      Logger.log('Error processing row: ' + e);
    }
  }
  return { totals: totals, periodData: periodData };
}

function processCampaignSearchTermView() {
  const dateClause = getDateRangeClause();
  const timeSegment = TIME_GRANULARITY === 'week' ? 'segments.week' : 'segments.month';
  // Must SELECT segments.search_term_targeting_status to filter excluded terms in code
  // (API returns all terms including excluded; we skip EXCLUDED / ADDED_EXCLUDED to match UI).
  const query = [
    'SELECT',
    '  campaign_search_term_view.search_term,',
    '  segments.search_term_targeting_status,',
    '  ' + timeSegment + ',',
    '  metrics.impressions,',
    '  metrics.clicks,',
    '  metrics.cost_micros,',
    '  metrics.conversions,',
    '  metrics.conversions_value',
    'FROM campaign_search_term_view',
    'WHERE ' + dateClause,
    "  AND campaign.advertising_channel_type = 'PERFORMANCE_MAX'"
  ].join('\n');

  const totals = { branded: emptyMetrics(), nonBranded: emptyMetrics() };
  const periodData = {};

  function isExcludedTargetingStatus(segments) {
    if (!segments) return false;
    const raw = segments.searchTermTargetingStatus !== undefined
      ? segments.searchTermTargetingStatus
      : segments.search_term_targeting_status;
    if (raw === undefined) return false;
    const s = String(raw).toUpperCase();
    return s === 'EXCLUDED' || s === 'ADDED_EXCLUDED';
  }

  const DEBUG_PMAX_ROWS = 0; // Log first N rows to see segments structure; set > 0 to enable
  let rowIndex = 0;
  let skippedExcluded = 0;

  const report = AdsApp.search(query);
  while (report.hasNext()) {
    try {
      const row = report.next();
      rowIndex++;

      const excluded = isExcludedTargetingStatus(row.segments);
      if (excluded) skippedExcluded++;

      if (DEBUG_PMAX_ROWS > 0 && rowIndex <= DEBUG_PMAX_ROWS) {
        const text = row.campaignSearchTermView && row.campaignSearchTermView.searchTerm ? row.campaignSearchTermView.searchTerm : '';
        Logger.log('[Pmax row ' + rowIndex + '] searchTerm="' + text + '" segments=' + JSON.stringify(row.segments) + ' excluded=' + excluded);
      }

      if (excluded) continue;

      const text = row.campaignSearchTermView && row.campaignSearchTermView.searchTerm ? row.campaignSearchTermView.searchTerm : '';
      const m = row.metrics || {};
      const cost = Number(m.costMicros) || 0;
      const costActual = cost / 1000000;
      const rowMetrics = {
        impressions: Number(m.impressions) || 0,
        clicks: Number(m.clicks) || 0,
        cost: costActual,
        conversions: Number(m.conversions) || 0,
        conversionsValue: Number(m.conversionsValue) || 0
      };
      const periodKey = TIME_GRANULARITY === 'week' ? (row.segments && row.segments.week) : (row.segments && row.segments.month);
      if (!periodKey) continue;

      const branded = PMAX_TREAT_ALL_AS_NON_BRANDED ? false : isBranded(text);
      if (branded) {
        totals.branded.impressions += rowMetrics.impressions;
        totals.branded.clicks += rowMetrics.clicks;
        totals.branded.cost += rowMetrics.cost;
        totals.branded.conversions += rowMetrics.conversions;
        totals.branded.conversionsValue += rowMetrics.conversionsValue;
      } else {
        totals.nonBranded.impressions += rowMetrics.impressions;
        totals.nonBranded.clicks += rowMetrics.clicks;
        totals.nonBranded.cost += rowMetrics.cost;
        totals.nonBranded.conversions += rowMetrics.conversions;
        totals.nonBranded.conversionsValue += rowMetrics.conversionsValue;
      }
      addRowToPeriodData(periodData, periodKey, branded, rowMetrics);
    } catch (e) {
      Logger.log('Error processing row: ' + e);
    }
  }

  if (DEBUG_PMAX_ROWS > 0) {
    Logger.log('[Pmax] Total rows: ' + rowIndex + ', skipped (excluded): ' + skippedExcluded);
  }

  return { totals: totals, periodData: periodData };
}

// Pmax Consumer Spotlight: campaign_search_term_insight provides category-level data.
// No cost metrics available; only impressions, clicks, conversions, conversion value.
// Note: This resource doesn't support segments.month/week/date with campaign_id filter,
// so we run separate queries for each time period.
function processCampaignSearchTermInsight() {
  const timeZone = AdsApp.currentAccount().getTimeZone();

  // First get all Pmax campaign IDs
  const campaignQuery = [
    'SELECT campaign.id, campaign.name',
    'FROM campaign',
    "WHERE campaign.advertising_channel_type = 'PERFORMANCE_MAX'",
    "  AND campaign.status != 'REMOVED'"
  ].join('\n');

  const campaignIds = [];
  const campaignReport = AdsApp.search(campaignQuery);
  while (campaignReport.hasNext()) {
    const row = campaignReport.next();
    if (row.campaign && row.campaign.id) {
      campaignIds.push(row.campaign.id);
    }
  }

  if (campaignIds.length === 0) {
    Logger.log('[Pmax Categories] No Pmax campaigns found.');
    return { totals: { branded: emptyMetrics(), nonBranded: emptyMetrics() }, periodData: {} };
  }

  Logger.log('[Pmax Categories] Found ' + campaignIds.length + ' Pmax campaign(s).');

  // Calculate period date ranges based on config
  function getPeriodRanges() {
    let startDate, endDate;
    if (START_DATE && END_DATE) {
      startDate = new Date(START_DATE);
      endDate = new Date(END_DATE);
    } else {
      endDate = new Date();
      startDate = new Date();
      startDate.setDate(endDate.getDate() - LOOKBACK_DAYS);
    }

    const periods = [];
    if (TIME_GRANULARITY === 'week') {
      // Generate week ranges (Monday to Sunday)
      let current = new Date(startDate);
      // Move to Monday of the first week
      const dayOfWeek = current.getDay();
      const diff = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
      current.setDate(current.getDate() + diff);

      while (current <= endDate) {
        const weekStart = new Date(current);
        const weekEnd = new Date(current);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const periodKey = Utilities.formatDate(weekStart, timeZone, 'yyyy-MM-dd');
        periods.push({
          key: periodKey,
          start: Utilities.formatDate(weekStart, timeZone, 'yyyy-MM-dd'),
          end: Utilities.formatDate(weekEnd > endDate ? endDate : weekEnd, timeZone, 'yyyy-MM-dd')
        });
        current.setDate(current.getDate() + 7);
      }
    } else {
      // Generate month ranges
      let current = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
      while (current <= endDate) {
        const monthStart = new Date(current);
        const monthEnd = new Date(current.getFullYear(), current.getMonth() + 1, 0); // Last day of month
        const periodKey = Utilities.formatDate(monthStart, timeZone, 'yyyy-MM');
        periods.push({
          key: periodKey,
          start: Utilities.formatDate(monthStart < startDate ? startDate : monthStart, timeZone, 'yyyy-MM-dd'),
          end: Utilities.formatDate(monthEnd > endDate ? endDate : monthEnd, timeZone, 'yyyy-MM-dd')
        });
        current.setMonth(current.getMonth() + 1);
      }
    }
    return periods;
  }

  const periodRanges = getPeriodRanges();
  Logger.log('[Pmax Categories] Processing ' + periodRanges.length + ' period(s).');

  const totals = { branded: emptyMetrics(), nonBranded: emptyMetrics(), blank: emptyMetrics() };
  const periodData = {};

  // Query each period separately (no segments needed)
  for (let p = 0; p < periodRanges.length; p++) {
    const period = periodRanges[p];
    const dateClause = "segments.date BETWEEN '" + period.start + "' AND '" + period.end + "'";

    for (let i = 0; i < campaignIds.length; i++) {
      const campaignId = campaignIds[i];
      const query = [
        'SELECT',
        '  campaign_search_term_insight.category_label,',
        '  metrics.impressions,',
        '  metrics.clicks,',
        '  metrics.conversions,',
        '  metrics.conversions_value',
        'FROM campaign_search_term_insight',
        'WHERE ' + dateClause,
        '  AND campaign_search_term_insight.campaign_id = ' + campaignId
      ].join('\n');

      try {
        const report = AdsApp.search(query);
        while (report.hasNext()) {
          const row = report.next();
          const categoryLabel = row.campaignSearchTermInsight && row.campaignSearchTermInsight.categoryLabel
            ? row.campaignSearchTermInsight.categoryLabel
            : '';
          const m = row.metrics || {};
          const rowMetrics = {
            impressions: Number(m.impressions) || 0,
            clicks: Number(m.clicks) || 0,
            cost: 0, // Not available in campaign_search_term_insight
            conversions: Number(m.conversions) || 0,
            conversionsValue: Number(m.conversionsValue) || 0
          };

          // Determine category: blank, branded, or nonBranded
          let category;
          if (!categoryLabel || categoryLabel.trim() === '') {
            category = 'blank';
          } else if (isBranded(categoryLabel)) {
            category = 'branded';
          } else {
            category = 'nonBranded';
          }

          // Update totals
          totals[category].impressions += rowMetrics.impressions;
          totals[category].clicks += rowMetrics.clicks;
          totals[category].conversions += rowMetrics.conversions;
          totals[category].conversionsValue += rowMetrics.conversionsValue;

          addRowToPeriodDataWithBlank(periodData, period.key, category, rowMetrics);
        }
      } catch (e) {
        Logger.log('[Pmax Categories] Error querying campaign ' + campaignId + ' for period ' + period.key + ': ' + e);
      }
    }
  }

  return { totals: totals, periodData: periodData };
}

function mergePeriodData(target, source) {
  let key;
  for (key in source) {
    if (!source.hasOwnProperty(key)) continue;
    if (!target[key]) {
      target[key] = emptyPeriodData();
    }
    target[key].branded.impressions += source[key].branded.impressions;
    target[key].branded.clicks += source[key].branded.clicks;
    target[key].branded.cost += source[key].branded.cost;
    target[key].branded.conversions += source[key].branded.conversions;
    target[key].branded.conversionsValue += source[key].branded.conversionsValue;
    target[key].nonBranded.impressions += source[key].nonBranded.impressions;
    target[key].nonBranded.clicks += source[key].nonBranded.clicks;
    target[key].nonBranded.cost += source[key].nonBranded.cost;
    target[key].nonBranded.conversions += source[key].nonBranded.conversions;
    target[key].nonBranded.conversionsValue += source[key].nonBranded.conversionsValue;
  }
}

function mergeTotals(target, source) {
  target.branded.impressions += source.branded.impressions;
  target.branded.clicks += source.branded.clicks;
  target.branded.cost += source.branded.cost;
  target.branded.conversions += source.branded.conversions;
  target.branded.conversionsValue += source.branded.conversionsValue;
  target.nonBranded.impressions += source.nonBranded.impressions;
  target.nonBranded.clicks += source.nonBranded.clicks;
  target.nonBranded.cost += source.nonBranded.cost;
  target.nonBranded.conversions += source.nonBranded.conversions;
  target.nonBranded.conversionsValue += source.nonBranded.conversionsValue;
}

// ===== FORMAT PERIOD LABEL =====
function formatPeriodLabel(periodKey) {
  if (TIME_GRANULARITY === 'week') {
    return 'w/c ' + periodKey;
  }
  const parts = String(periodKey).split('-');
  if (parts.length >= 2) {
    const year = parts[0];
    const monthNum = parseInt(parts[1], 10);
    const d = new Date(year, monthNum - 1, 1);
    return Utilities.formatDate(d, AdsApp.currentAccount().getTimeZone(), 'MMM yyyy');
  }
  return periodKey;
}

// ===== BUILD RAW TAB ROWS =====
function buildRawTabRows(periodData) {
  const periods = Object.keys(periodData).sort();
  const rows = [];
  rows.push(['Period', 'Segment', 'Impressions', 'Clicks', 'Cost', 'Conversions', 'Conversion Value', 'CPA', 'ROAS']);
  periods.forEach(function (periodKey) {
    const p = periodData[periodKey];
    const label = formatPeriodLabel(periodKey);
    const cpaB = p.branded.conversions > 0 ? p.branded.cost / p.branded.conversions : 0;
    const cpaN = p.nonBranded.conversions > 0 ? p.nonBranded.cost / p.nonBranded.conversions : 0;
    const roasB = p.branded.cost > 0 ? p.branded.conversionsValue / p.branded.cost : 0;
    const roasN = p.nonBranded.cost > 0 ? p.nonBranded.conversionsValue / p.nonBranded.cost : 0;
    rows.push([label, 'Branded', p.branded.impressions, p.branded.clicks, p.branded.cost, p.branded.conversions, p.branded.conversionsValue, cpaB, roasB]);
    rows.push([label, 'Non-branded', p.nonBranded.impressions, p.nonBranded.clicks, p.nonBranded.cost, p.nonBranded.conversions, p.nonBranded.conversionsValue, cpaN, roasN]);
  });
  return rows;
}

// Build raw rows for categories (no cost metrics available, includes blank)
function buildRawTabRowsNoCost(periodData) {
  const periods = Object.keys(periodData).sort();
  const rows = [];
  rows.push(['Period', 'Segment', 'Impressions', 'Clicks', 'Conversions', 'Conversion Value']);
  periods.forEach(function (periodKey) {
    const p = periodData[periodKey];
    const label = formatPeriodLabel(periodKey);
    rows.push([label, 'Branded', p.branded.impressions, p.branded.clicks, p.branded.conversions, p.branded.conversionsValue]);
    rows.push([label, 'Non-branded', p.nonBranded.impressions, p.nonBranded.clicks, p.nonBranded.conversions, p.nonBranded.conversionsValue]);
    // Blank category for unidentifiable search categories
    if (p.blank) {
      rows.push([label, 'Blank', p.blank.impressions, p.blank.clicks, p.blank.conversions, p.blank.conversionsValue]);
    }
  });
  return rows;
}

// ===== SHEET: INFO + RAW TABS =====
function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
  } else {
    sh.clear();
  }
  return sh;
}

function writeInfoAndRawTabs(ss, combined, byType, dateRangeStr) {
  const account = AdsApp.currentAccount();
  const timeZone = account.getTimeZone();
  const currency = account.getCurrencyCode();

  const infoSheet = getOrCreateSheet(ss, 'Info');
  const infoData = [
    ['Account Name', account.getName()],
    ['Account ID', account.getCustomerId()],
    ['Currency', currency],
    ['Date Range', dateRangeStr],
    ['Time Granularity', TIME_GRANULARITY],
    ['Brand Tokens', BRAND_TOKENS.join(', ')],
    ['Run Timestamp', Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss')]
  ];
  infoSheet.getRange(1, 1, infoData.length, 2).setValues(infoData);

  const rawCombined = getOrCreateSheet(ss, 'Raw - Combined');
  const combinedRows = buildRawTabRows(combined.periodData);
  if (combinedRows.length > 1) {
    rawCombined.getRange(1, 1, combinedRows.length, combinedRows[0].length).setValues(combinedRows);
  }

  if (INCLUDE_BY_CAMPAIGN_TYPE && byType) {
    ['Search', 'Pmax', 'Shopping'].forEach(function (channel) {
      const data = byType[channel];
      if (!data) return;
      const sh = getOrCreateSheet(ss, 'Raw - ' + channel);
      const rows = buildRawTabRows(data.periodData);
      if (rows.length > 1) {
        sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
      }
    });
  }
}

function writePmaxCategoriesTab(ss, pmaxCategoriesData) {
  if (!INCLUDE_PMAX_CATEGORIES || !pmaxCategoriesData) return;
  const sh = getOrCreateSheet(ss, 'Raw - Pmax Categories');
  const rows = buildRawTabRowsNoCost(pmaxCategoriesData.periodData);
  if (rows.length > 1) {
    sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  }
}

// ===== CHARTS =====
// Build chart data rows for one metric (Period, Branded, Non-branded)
function buildChartDataRows(periodData, valueType) {
  const periods = Object.keys(periodData).sort();
  const rows = [];
  periods.forEach(function (periodKey) {
    const label = formatPeriodLabel(periodKey);
    const p = periodData[periodKey];
    let valB = p.branded[valueType];
    let valN = p.nonBranded[valueType];
    if (valueType === 'cpa') {
      valB = p.branded.conversions > 0 ? p.branded.cost / p.branded.conversions : 0;
      valN = p.nonBranded.conversions > 0 ? p.nonBranded.cost / p.nonBranded.conversions : 0;
    } else if (valueType === 'roas') {
      valB = p.branded.cost > 0 ? p.branded.conversionsValue / p.branded.cost : 0;
      valN = p.nonBranded.cost > 0 ? p.nonBranded.conversionsValue / p.nonBranded.cost : 0;
    }
    rows.push([label, valB, valN]);
  });
  return rows;
}

// Build branded ratio data (branded / total as percentage)
function buildBrandedRatioRows(periodData, valueType, includeBlank) {
  const periods = Object.keys(periodData).sort();
  const rows = [];
  periods.forEach(function (periodKey) {
    const label = formatPeriodLabel(periodKey);
    const p = periodData[periodKey];
    let valB, valN, valBlank;
    if (valueType === 'cpa') {
      valB = p.branded.conversions > 0 ? p.branded.cost / p.branded.conversions : 0;
      valN = p.nonBranded.conversions > 0 ? p.nonBranded.cost / p.nonBranded.conversions : 0;
      valBlank = (includeBlank && p.blank && p.blank.conversions > 0) ? p.blank.cost / p.blank.conversions : 0;
    } else if (valueType === 'roas') {
      valB = p.branded.cost > 0 ? p.branded.conversionsValue / p.branded.cost : 0;
      valN = p.nonBranded.cost > 0 ? p.nonBranded.conversionsValue / p.nonBranded.cost : 0;
      valBlank = (includeBlank && p.blank && p.blank.cost > 0) ? p.blank.conversionsValue / p.blank.cost : 0;
    } else {
      valB = p.branded ? p.branded[valueType] || 0 : 0;
      valN = p.nonBranded ? p.nonBranded[valueType] || 0 : 0;
      valBlank = (includeBlank && p.blank) ? p.blank[valueType] || 0 : 0;
    }
    const total = valB + valN + (includeBlank ? valBlank : 0);
    const ratio = total > 0 ? valB / total : 0;
    rows.push([label, ratio]);
  });
  return rows;
}

function writeChartsForView(ss, periodData, chartTabName, currency) {
  const color1 = '#4285F4';
  const color2 = '#FBBC05';
  const colorRatio = '#34A853';  // Green for ratio line
  const currFmt = '"' + currency + '" #,##0.00';
  const periodLabel = TIME_GRANULARITY === 'week' ? 'Week' : 'Month';

  let chartSheet = ss.getSheetByName(chartTabName);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(chartTabName);
  }
  chartSheet.clear();

  const metrics = [
    { valueType: 'impressions', title: 'Impressions', format: '#,##0' },
    { valueType: 'clicks', title: 'Clicks', format: '#,##0' },
    { valueType: 'cost', title: 'Cost (' + currency + ')', format: currFmt },
    { valueType: 'conversions', title: 'Conversions', format: '#,##0.00' },
    { valueType: 'conversionsValue', title: 'Conversion Value (' + currency + ')', format: currFmt },
    { valueType: 'cpa', title: 'Cost per Conversion (' + currency + ')', format: currFmt },
    { valueType: 'roas', title: 'ROAS', format: '#,##0.00' }
  ];

  let startRow = 1;
  const chartWidth = 500;
  const ratioChartWidth = 400;
  const chartHeight = 300;
  const rowHeight = 25;
  const dataColStart = 1;  // Column A for bar chart data
  const ratioColStart = 5; // Column E for ratio data (leaving gap)

  metrics.forEach(function (m, idx) {
    const rows = buildChartDataRows(periodData, m.valueType);
    if (rows.length === 0) return;

    // Bar chart data (columns A-C)
    const header = [['Period', 'Branded', 'Non-branded']];
    const allRows = header.concat(rows);
    const numRows = allRows.length;
    const dataStartRow = startRow + 1;
    const numDataRows = numRows - 1;

    chartSheet.getRange(startRow, dataColStart, numRows, 3).setValues(allRows);
    if (m.format && numDataRows >= 1) {
      chartSheet.getRange(dataStartRow, dataColStart + 1, numDataRows, 2).setNumberFormat(m.format);
    }

    // Ratio data (columns E-F)
    const ratioRows = buildBrandedRatioRows(periodData, m.valueType, false);
    const ratioHeader = [['Period', '% Branded']];
    const allRatioRows = ratioHeader.concat(ratioRows);
    chartSheet.getRange(startRow, ratioColStart, numRows, 2).setValues(allRatioRows);
    chartSheet.getRange(dataStartRow, ratioColStart + 1, numDataRows, 1).setNumberFormat('0.0%');

    // Bar chart (positioned below data, column A)
    const dataRange = chartSheet.getRange(dataStartRow, dataColStart, numDataRows, 3);
    const barChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange)
      .setPosition(startRow + numRows, 1, 0, 0)
      .setOption('title', m.title + ' by ' + periodLabel + ' (Branded vs Non-branded)')
      .setOption('legend', { position: 'bottom' })
      .setOption('isStacked', true)
      .setOption('series', {
        0: { labelInLegend: 'Branded', color: color1 },
        1: { labelInLegend: 'Non-branded', color: color2 }
      })
      .setOption('colors', [color1, color2])
      .setOption('vAxis', { title: m.title })
      .setOption('hAxis', { title: periodLabel, slantedText: true, slantedTextAngle: 45 })
      .setOption('width', chartWidth)
      .setOption('height', chartHeight)
      .build();
    chartSheet.insertChart(barChart);

    // Ratio line chart (positioned to the right of bar chart)
    const ratioXRange = chartSheet.getRange(startRow, ratioColStart, numRows, 1);
    const ratioYRange = chartSheet.getRange(startRow, ratioColStart + 1, numRows, 1);
    const ratioChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(ratioXRange)
      .addRange(ratioYRange)
      .setPosition(startRow + numRows, 7, 0, 0)  // Column G (to the right)
      .setOption('title', 'Branded Dependency Ratio - ' + m.title)
      .setOption('legend', { position: 'none' })
      .setOption('pointSize', 5)
      .setOption('series', {
        0: { color: colorRatio }
      })
      .setOption('colors', [colorRatio])
      .setOption('vAxis', { title: '% Branded', format: 'percent', minValue: 0, maxValue: 1 })
      .setOption('hAxis', { title: periodLabel, slantedText: true, slantedTextAngle: 45 })
      .setOption('width', ratioChartWidth)
      .setOption('height', chartHeight)
      .build();
    chartSheet.insertChart(ratioChart);

    startRow += numRows + Math.ceil(chartHeight / rowHeight) + 2;
  });
}

function writeAllCharts(ss, combined, byType) {
  const currency = AdsApp.currentAccount().getCurrencyCode();
  writeChartsForView(ss, combined.periodData, 'Charts - Combined', currency);
  if (INCLUDE_BY_CAMPAIGN_TYPE && byType) {
    if (byType.Search) writeChartsForView(ss, byType.Search.periodData, 'Charts - Search', currency);
    if (byType.Pmax) writeChartsForView(ss, byType.Pmax.periodData, 'Charts - Pmax', currency);
    if (byType.Shopping) writeChartsForView(ss, byType.Shopping.periodData, 'Charts - Shopping', currency);
  }
}

// Build chart data rows for categories view (includes blank)
function buildChartDataRowsWithBlank(periodData, valueType) {
  const periods = Object.keys(periodData).sort();
  const rows = [];
  periods.forEach(function (periodKey) {
    const label = formatPeriodLabel(periodKey);
    const p = periodData[periodKey];
    const valB = p.branded ? p.branded[valueType] || 0 : 0;
    const valN = p.nonBranded ? p.nonBranded[valueType] || 0 : 0;
    const valBlank = p.blank ? p.blank[valueType] || 0 : 0;
    rows.push([label, valB, valN, valBlank]);
  });
  return rows;
}

// Charts for Pmax Categories (no cost metrics, includes blank, stacked column charts + ratio line charts)
function writeChartsForCategoriesView(ss, periodData, chartTabName, currency) {
  const color1 = '#4285F4';  // Blue for branded
  const color2 = '#FBBC05';  // Yellow for non-branded
  const color3 = '#BEBEBE';  // Gray for blank
  const colorRatio = '#34A853';  // Green for ratio line
  const currFmt = '"' + currency + '" #,##0.00';
  const periodLabel = TIME_GRANULARITY === 'week' ? 'Week' : 'Month';

  let chartSheet = ss.getSheetByName(chartTabName);
  if (!chartSheet) {
    chartSheet = ss.insertSheet(chartTabName);
  }
  chartSheet.clear();

  // Only 4 metrics - no cost or CPA
  const metrics = [
    { valueType: 'impressions', title: 'Impressions', format: '#,##0' },
    { valueType: 'clicks', title: 'Clicks', format: '#,##0' },
    { valueType: 'conversions', title: 'Conversions', format: '#,##0.00' },
    { valueType: 'conversionsValue', title: 'Conversion Value (' + currency + ')', format: currFmt }
  ];

  let startRow = 1;
  const chartWidth = 500;
  const ratioChartWidth = 400;
  const chartHeight = 300;
  const rowHeight = 25;
  const dataColStart = 1;  // Column A for bar chart data
  const ratioColStart = 6; // Column F for ratio data (leaving gap after D)

  metrics.forEach(function (m, idx) {
    const rows = buildChartDataRowsWithBlank(periodData, m.valueType);
    if (rows.length === 0) return;

    // Bar chart data (columns A-D): Period, Branded, Non-branded, Blank
    const header = [['Period', 'Branded', 'Non-branded', 'Blank']];
    const allRows = header.concat(rows);
    const numRows = allRows.length;
    const numDataRows = rows.length;

    chartSheet.getRange(startRow, dataColStart, numRows, 4).setValues(allRows);
    if (m.format && numDataRows >= 1) {
      chartSheet.getRange(startRow + 1, dataColStart + 1, numDataRows, 3).setNumberFormat(m.format);
    }

    // Ratio data (columns F-G) - includeBlank = true for categories
    const ratioRows = buildBrandedRatioRows(periodData, m.valueType, true);
    const ratioHeader = [['Period', '% Branded']];
    const allRatioRows = ratioHeader.concat(ratioRows);
    chartSheet.getRange(startRow, ratioColStart, numRows, 2).setValues(allRatioRows);
    chartSheet.getRange(startRow + 1, ratioColStart + 1, numDataRows, 1).setNumberFormat('0.0%');

    // Bar chart (positioned below data, column A)
    const xAxisRange = chartSheet.getRange(startRow, dataColStart, numRows, 1);
    const dataSeriesRange = chartSheet.getRange(startRow, dataColStart + 1, numRows, 3);
    
    const barChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(xAxisRange)
      .addRange(dataSeriesRange)
      .setPosition(startRow + numRows, 1, 0, 0)
      .setOption('title', m.title + ' by ' + periodLabel + ' (Branded vs Non-branded) - Consumer Spotlight')
      .setOption('legend', { position: 'bottom' })
      .setOption('isStacked', true)
      .setOption('series', {
        0: { labelInLegend: 'Branded', color: color1 },
        1: { labelInLegend: 'Non-branded', color: color2 },
        2: { labelInLegend: 'Blank', color: color3 }
      })
      .setOption('colors', [color1, color2, color3])
      .setOption('vAxis', { title: m.title })
      .setOption('hAxis', { title: periodLabel, slantedText: true, slantedTextAngle: 45 })
      .setOption('width', chartWidth)
      .setOption('height', chartHeight)
      .build();
    chartSheet.insertChart(barChart);

    // Ratio line chart (positioned to the right of bar chart)
    const ratioXRange = chartSheet.getRange(startRow, ratioColStart, numRows, 1);
    const ratioYRange = chartSheet.getRange(startRow, ratioColStart + 1, numRows, 1);
    const ratioChart = chartSheet.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(ratioXRange)
      .addRange(ratioYRange)
      .setPosition(startRow + numRows, 8, 0, 0)  // Column H (to the right)
      .setOption('title', 'Branded Dependency Ratio - ' + m.title)
      .setOption('legend', { position: 'none' })
      .setOption('pointSize', 5)
      .setOption('series', {
        0: { color: colorRatio }
      })
      .setOption('colors', [colorRatio])
      .setOption('vAxis', { title: '% Branded', format: 'percent', minValue: 0, maxValue: 1 })
      .setOption('hAxis', { title: periodLabel, slantedText: true, slantedTextAngle: 45 })
      .setOption('width', ratioChartWidth)
      .setOption('height', chartHeight)
      .build();
    chartSheet.insertChart(ratioChart);

    startRow += numRows + Math.ceil(chartHeight / rowHeight) + 2;
  });
}

function writePmaxCategoriesCharts(ss, pmaxCategoriesData) {
  if (!INCLUDE_PMAX_CATEGORIES || !pmaxCategoriesData) return;
  const currency = AdsApp.currentAccount().getCurrencyCode();
  writeChartsForCategoriesView(ss, pmaxCategoriesData.periodData, 'Charts - Pmax Categories', currency);
}

// ===== TAB ORDER =====
function reorderTabs(ss) {
  const order = [
    'Info',
    'Raw - Combined',
    'Raw - Search',
    'Raw - Pmax',
    'Raw - Pmax Categories',
    'Raw - Shopping',
    'Charts - Combined',
    'Charts - Search',
    'Charts - Pmax',
    'Charts - Pmax Categories',
    'Charts - Shopping'
  ];
  order.forEach(function (name, idx) {
    const tab = ss.getSheetByName(name);
    if (tab) {
      ss.setActiveSheet(tab);
      ss.moveActiveSheet(idx + 1);
    }
  });
}

// ===== MAIN =====
function main() {
  try {
    BRAND_PATTERNS = buildBrandPatterns();

    const dateClause = getDateRangeClause();
    const dateRangeStr = START_DATE && END_DATE ? START_DATE + ' to ' + END_DATE : 'Last ' + LOOKBACK_DAYS + ' days';

    const searchData = processSearchTermView('SEARCH');
    const shoppingData = processSearchTermView('SHOPPING');
    const pmaxData = processCampaignSearchTermView();

    const combined = {
      totals: { branded: emptyMetrics(), nonBranded: emptyMetrics() },
      periodData: {}
    };
    mergeTotals(combined.totals, searchData.totals);
    mergeTotals(combined.totals, shoppingData.totals);
    mergeTotals(combined.totals, pmaxData.totals);
    mergePeriodData(combined.periodData, searchData.periodData);
    mergePeriodData(combined.periodData, shoppingData.periodData);
    mergePeriodData(combined.periodData, pmaxData.periodData);

    let byType = null;
    if (INCLUDE_BY_CAMPAIGN_TYPE) {
      byType = {
        Search: searchData,
        Pmax: pmaxData,
        Shopping: shoppingData
      };
    }

    // Pmax Categories (Consumer Spotlight) - separate data source
    let pmaxCategoriesData = null;
    if (INCLUDE_PMAX_CATEGORIES) {
      pmaxCategoriesData = processCampaignSearchTermInsight();
    }

    const sheetId = getSheetId(SHEET_URL);
    const ss = SpreadsheetApp.openById(sheetId);

    writeInfoAndRawTabs(ss, combined, byType, dateRangeStr);
    writePmaxCategoriesTab(ss, pmaxCategoriesData);
    writeAllCharts(ss, combined, byType);
    writePmaxCategoriesCharts(ss, pmaxCategoriesData);
    reorderTabs(ss);

    Logger.log('Script completed successfully.');
  } catch (e) {
    Logger.log('Script failed: ' + e.toString());
    throw e;
  }
}
