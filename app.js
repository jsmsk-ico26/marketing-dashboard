/**
 * マーケティング・ダッシュボード
 * 広告代理店向けマーケティング分析レポート生成Webアプリ
 */

// ===================================
// 定数 & 設定
// ===================================
const STORAGE_KEY_CLIENTS = 'mktg_dashboard_clients';
const STORAGE_KEY_ACTIVE = 'mktg_dashboard_active_client';

// 会社共通のマスター顧客名簿スプレッドシート
const MASTER_CLIENT_SHEET_ID = '1Zg9kByoLMidbe5i4FEGzgQPGExHTZYEeh6WouGaSINA';

// Google ログイン設定
const GOOGLE_CLIENT_ID = '726058308400-251vl121cbmsc305ju1m6e3a42f5gqv7.apps.googleusercontent.com';
const ALLOWED_DOMAINS = ['gws.ico-ad.co.jp']; // 自社GWSドメインのみに制限

// アプリ状態管理
const AppState = {
    isAuthenticated: false,
    user: null,
    // 顧客管理
    clients: [],
    activeClientId: null,
    showClientModal: false,
    showClientManager: false,
    // データ
    data: {
        ga4: [],
        adCost: [],
        measures: []
    },
    mergedData: [],
    filteredData: [],
    isLoading: false,
    activeTab: 'overview',
    charts: {},
    // 期間フィルター
    dateFilter: {
        startDate: '',
        endDate: '',
        isActive: false
    },
    // 期間比較モード
    compareMode: {
        enabled: false,
        periodA: { start: '', end: '', label: '期間A' },
        periodB: { start: '', end: '', label: '期間B' }
    }
};

// ===================================
// 顧客管理
// ===================================

/**
 * localStorage + 共有スプレッドシートから顧客リストを読み込み
 */
async function loadClients() {
    let localClients = [];
    let masterClients = [];

    // 1. ローカル保存（各自が追加した分）を読み込み
    try {
        const saved = localStorage.getItem(STORAGE_KEY_CLIENTS);
        if (saved) {
            localClients = JSON.parse(saved);
        }
    } catch (e) {
        console.warn('ローカル顧客データの読み込みに失敗:', e);
    }

    // 2. マスター名簿（会社共通）をスプレッドシートから読み込み
    try {
        const res = await fetch(`https://docs.google.com/spreadsheets/d/${MASTER_CLIENT_SHEET_ID}/gviz/tq?tqx=out:csv`);
        const csvText = await res.text();
        const rawMasterData = parseCSV(csvText);

        // スプシの列名に合わせてマッピング
        masterClients = rawMasterData.map((row, index) => ({
            id: 'master-' + index,
            name: row['顧客名'] || row[Object.keys(row)[0]], // 1列目
            siteName: row['サイト名'] || row[Object.keys(row)[1]] || '共通', // 2列目
            spreadsheetId: extractSpreadsheetId(row['スプレッドシートID'] || row[Object.keys(row)[2]]), // 3列目
            createdAt: 'マスター共有',
            isMaster: true
        })).filter(c => c.spreadsheetId);
    } catch (e) {
        console.warn('会社共有マスターの読み込みに失敗:', e);
    }

    // 両方を統合（重複はスプシIDで排除）
    const allClients = [...masterClients];
    localClients.forEach(lc => {
        if (!allClients.find(mc => mc.spreadsheetId === lc.spreadsheetId)) {
            allClients.push(lc);
        }
    });

    AppState.clients = allClients;

    // アクティブ顧客を復元
    const savedActive = localStorage.getItem(STORAGE_KEY_ACTIVE);
    if (savedActive && AppState.clients.find(c => c.id === savedActive)) {
        AppState.activeClientId = savedActive;
    } else if (AppState.clients.length > 0) {
        AppState.activeClientId = AppState.clients[0].id;
    }
}

/**
 * 顧客リストをlocalStorageに保存
 */
function saveClients() {
    try {
        localStorage.setItem(STORAGE_KEY_CLIENTS, JSON.stringify(AppState.clients));
        localStorage.setItem(STORAGE_KEY_ACTIVE, AppState.activeClientId || '');
    } catch (e) {
        console.warn('顧客データの保存に失敗:', e);
    }
}

/**
 * 現在アクティブな顧客情報を取得
 */
function getActiveClient() {
    return AppState.clients.find(c => c.id === AppState.activeClientId) || null;
}

/**
 * スプレッドシートURLからIDを抽出
 */
function extractSpreadsheetId(urlOrId) {
    if (!urlOrId) return '';
    // URL形式の場合: https://docs.google.com/spreadsheets/d/XXXXX/edit...
    const match = urlOrId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (match) return match[1];
    // 既にIDのみの場合
    return urlOrId.trim();
}

// ===================================
// JSONレポート出力
// ===================================

async function exportJSON() {
    if (AppState.filteredData.length === 0) {
        alert('出力対象のデータがありません。');
        return;
    }

    const client = getActiveClient();
    const currentData = AppState.filteredData;
    const currentMetrics = calculatePeriodMetrics(currentData);
    const insights = generateInsights();

    // 期間の算出
    const start = new Date(AppState.dateFilter.startDate);
    const end = new Date(AppState.dateFilter.endDate);

    // 前月期間
    const lastMonthStart = new Date(start);
    lastMonthStart.setMonth(lastMonthStart.getMonth() - 1);
    const lastMonthEnd = new Date(end);
    lastMonthEnd.setMonth(lastMonthEnd.getMonth() - 1);

    // 前年同月期間
    const lastYearStart = new Date(start);
    lastYearStart.setFullYear(lastYearStart.getFullYear() - 1);
    const lastYearEnd = new Date(end);
    lastYearEnd.setFullYear(lastYearEnd.getFullYear() - 1);

    const filterByDate = (s, e) => AppState.mergedData.filter(d => {
        const dt = new Date(d.date);
        return dt >= s && dt <= e;
    });

    const lastMonthData = filterByDate(lastMonthStart, lastMonthEnd);
    const lastYearData = filterByDate(lastYearStart, lastYearEnd);

    const report = {
        meta: {
            exportDate: new Date().toISOString(),
            clientName: client?.name || 'Unknown',
            siteName: client?.siteName || '',
            period: {
                start: AppState.dateFilter.startDate,
                end: AppState.dateFilter.endDate,
                days: currentData.length
            }
        },
        summary: currentMetrics,
        analysis: insights.map(i => ({ title: i.title.replace(/<[^>]*>?/gm, ''), content: i.content.replace(/<[^>]*>?/gm, '') })),
        comparisons: {
            lastMonth: {
                period: { start: lastMonthStart.toLocaleDateString(), end: lastMonthEnd.toLocaleDateString() },
                metrics: calculatePeriodMetrics(lastMonthData),
                diff: calculatePeriodDiff(currentMetrics, calculatePeriodMetrics(lastMonthData))
            },
            lastYearSameMonth: {
                period: { start: lastYearStart.toLocaleDateString(), end: lastYearEnd.toLocaleDateString() },
                metrics: calculatePeriodMetrics(lastYearData),
                diff: calculatePeriodDiff(currentMetrics, calculatePeriodMetrics(lastYearData))
            }
        },
        details: currentData.map(d => ({
            date: d.date,
            sessions: d.sessions,
            conversions: d.conversions,
            cvr: d.cvr,
            cpa: d.cpa,
            adCost: d.adCost,
            measures: d.measures || []
        }))
    };

    const blob = new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Report_${client?.name || 'Client'}_${AppState.dateFilter.startDate.replace(/\//g, '')}-${AppState.dateFilter.endDate.replace(/\//g, '')}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/**
 * 新規顧客を登録
 */
function addClient(clientName, siteName, spreadsheetUrl) {
    if (!clientName || !spreadsheetUrl) {
        alert('顧客名とスプレッドシートURLを入力してください。');
        return false;
    }

    const spreadsheetId = extractSpreadsheetId(spreadsheetUrl);
    if (!spreadsheetId) {
        alert('スプレッドシートのURLまたはIDが正しくありません。');
        return false;
    }

    const newClient = {
        id: 'client-' + Date.now(),
        name: clientName.trim(),
        siteName: (siteName || '').trim() || 'メインサイト',
        spreadsheetId: spreadsheetId,
        createdAt: new Date().toISOString().split('T')[0]
    };

    AppState.clients.push(newClient);
    AppState.activeClientId = newClient.id;
    saveClients();
    return true;
}

/**
 * 顧客を削除（確認付き）
 */
function removeClient(clientId) {
    const client = AppState.clients.find(c => c.id === clientId);
    if (!client) return;
    if (client.id === DEFAULT_CLIENT.id) {
        alert('デモ顧客は削除できません。');
        return;
    }
    if (!confirm(`「${client.name} / ${client.siteName}」を削除しますか？\n※データはスプレッドシートに残るので、再登録すれば復元できます。`)) return;

    AppState.clients = AppState.clients.filter(c => c.id !== clientId);
    if (AppState.activeClientId === clientId) {
        AppState.activeClientId = AppState.clients[0]?.id || null;
    }
    saveClients();
    if (AppState.activeClientId) switchClient(AppState.activeClientId);
    else renderApp();
}

/**
 * 顧客を切り替えてデータを再取得
 */
async function switchClient(clientId) {
    if (AppState.activeClientId === clientId && AppState.mergedData.length > 0) return;
    AppState.activeClientId = clientId;
    // 期間フィルターをリセット
    AppState.dateFilter = { startDate: '', endDate: '', isActive: false };
    AppState.compareMode = { enabled: false, periodA: { start: '', end: '', label: '期間A' }, periodB: { start: '', end: '', label: '期間B' } };
    saveClients();
    await fetchSheetData();
}

// ===================================
// データ取得 & パース
// ===================================

/**
 * CSVテキストをパースしオブジェクト配列に変換
 */
function parseCSV(csvText) {
    const lines = csvText.trim().split('\n');
    if (lines.length < 2) return [];

    const headers = lines[0].split(',').map(h => h.replace(/"/g, '').trim());
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',').map(v => v.replace(/"/g, '').trim());
        const row = {};
        headers.forEach((h, idx) => {
            row[h] = values[idx] || '';
        });
        rows.push(row);
    }
    return rows;
}

/**
 * スプレッドシートの実データ（フォールバック用）
 * CORSエラー時やオフライン時にこのデータを使用
 */
const FALLBACK_DATA = {
    ga4: [
        { '日付': '2026/03/01', 'セッション': '1200', 'ユーザー': '1000', 'エンゲージメント率': '65.50%', 'キーイベント（成約）': '12', 'CVR(%)': '1.00%' },
        { '日付': '2026/03/02', 'セッション': '1150', 'ユーザー': '950', 'エンゲージメント率': '64.20%', 'キーイベント（成約）': '10', 'CVR(%)': '0.90%' },
        { '日付': '2026/03/03', 'セッション': '1300', 'ユーザー': '1100', 'エンゲージメント率': '63.80%', 'キーイベント（成約）': '11', 'CVR(%)': '0.80%' },
        { '日付': '2026/03/04', 'セッション': '1250', 'ユーザー': '1050', 'エンゲージメント率': '72.10%', 'キーイベント（成約）': '25', 'CVR(%)': '2.00%' },
        { '日付': '2026/03/05', 'セッション': '1400', 'ユーザー': '1200', 'エンゲージメント率': '75.30%', 'キーイベント（成約）': '32', 'CVR(%)': '2.30%' }
    ],
    adCost: [
        { '日付': '2026/03/01', '媒体': 'Google広告', '費用(円)': '15000', '表示回数': '50000', 'クリック数': '500', 'CTR(%)': '1.00%', 'CPC(円)': '30' },
        { '日付': '2026/03/02', '媒体': 'Google広告', '費用(円)': '15000', '表示回数': '48000', 'クリック数': '450', 'CTR(%)': '0.90%', 'CPC(円)': '33' },
        { '日付': '2026/03/03', '媒体': 'Google広告', '費用(円)': '15000', '表示回数': '52000', 'クリック数': '550', 'CTR(%)': '1.10%', 'CPC(円)': '27' },
        { '日付': '2026/03/04', '媒体': 'Google広告', '費用(円)': '20000', '表示回数': '60000', 'クリック数': '700', 'CTR(%)': '1.20%', 'CPC(円)': '29' },
        { '日付': '2026/03/05', '媒体': 'Google広告', '費用(円)': '20000', '表示回数': '65000', 'クリック数': '800', 'CTR(%)': '1.20%', 'CPC(円)': '25' }
    ],
    measures: [
        { '実施日': '2026/03/03', 'カテゴリ': 'LP改善', '対象': 'コンバージョンボタン', '変更内容': '色をグレーからオレンジに変更', '狙い・仮説': 'ボタンの視認性を高め、クリックの心理的ハードルを下げる' },
        { '実施日': '2026/03/04', 'カテゴリ': '広告運用', '対象': '予算増額', '変更内容': '日予算を1.5万円→2万円へ', '狙い・仮説': 'LP改善の効果を確認できたため、流入数を増やして最大化を図る' }
    ]
};

/**
 * Google Sheetsから全タブのデータを取得（アクティブ顧客のスプレッドシートから）
 */
async function fetchSheetData() {
    AppState.isLoading = true;
    renderApp();

    const client = getActiveClient();
    let usedFallback = false;

    if (!client) {
        AppState.isLoading = false;
        renderApp();
        return;
    }

    const sheetsBase = `https://docs.google.com/spreadsheets/d/${client.spreadsheetId}/gviz/tq?tqx=out:csv`;

    try {
        const [ga4Res, adRes, measureRes] = await Promise.all([
            fetch(`${sheetsBase}&sheet=GA4`),
            fetch(`${sheetsBase}&sheet=${encodeURIComponent('広告費')}`),
            fetch(`${sheetsBase}&sheet=${encodeURIComponent('施策ログ')}`)
        ]);

        const [ga4Text, adText, measureText] = await Promise.all([
            ga4Res.text(),
            adRes.text(),
            measureRes.text()
        ]);

        AppState.data.ga4 = parseCSV(ga4Text);
        AppState.data.adCost = parseCSV(adText);
        AppState.data.measures = parseCSV(measureText);

        if (AppState.data.ga4.length === 0) throw new Error('パースされたデータが空です');

    } catch (error) {
        console.warn('Google Sheetsからの取得に失敗。フォールバックデータを使用します:', error.message);
        // デモ顧客の場合のみフォールバックデータを使用
        if (client.id === DEFAULT_CLIENT.id) {
            AppState.data.ga4 = FALLBACK_DATA.ga4;
            AppState.data.adCost = FALLBACK_DATA.adCost;
            AppState.data.measures = FALLBACK_DATA.measures;
        } else {
            AppState.data.ga4 = [];
            AppState.data.adCost = [];
            AppState.data.measures = [];
        }
        usedFallback = true;
    }

    // データを日付キーで統合
    mergeDataByDate();

    AppState.isLoading = false;
    AppState.usedFallback = usedFallback;
    renderApp();

    // チャート描画（DOM描画後に実行）
    setTimeout(() => {
        renderCharts();
    }, 200);
}

/**
 * 日付をキーにして3タブのデータを統合
 */
function mergeDataByDate() {
    const dateMap = {};

    // GA4データ
    AppState.data.ga4.forEach(row => {
        // ヘッダー名（日本語/英語）の柔軟な解決
        const date = row['日付'] || row['Date'] || row['date'];
        if (!date) return;

        const dKey = date.replace(/-/g, '/'); // 書式を統一
        if (!dateMap[dKey]) dateMap[dKey] = { date: dKey };

        dateMap[dKey].sessions = parseInt(row['セッション'] || row['Sessions'] || row['sessions']) || 0;
        dateMap[dKey].users = parseInt(row['ユーザー'] || row['Total users'] || row['activeUsers']) || 0;
        dateMap[dKey].engagementRate = parseFloat(row['エンゲージメント率'] || row['Engagement rate'] || row['engagementRate']) || 0;
        dateMap[dKey].conversions = parseInt(row['キーイベント（成約）'] || row['Conversions'] || row['conversions'] || row['events']) || 0;
        dateMap[dKey].cvr = parseFloat(row['CVR(%)'] || row['Session conversion rate'] || row['cvr']) || 0;

        // GA4連携時の広告費（Google Ads）があれば自動セット
        const adCostFromGA4 = parseInt(row['advertiserAdCost'] || row['Google Ads cost'] || 0);
        if (adCostFromGA4 > 0) {
            dateMap[dKey].adCost = adCostFromGA4;
            dateMap[dKey].medium = 'Google Ads (Auto)';
        }
    });

    // 広告費データ
    AppState.data.adCost.forEach(row => {
        const date = row['日付'] || row['Date'] || row['date'];
        if (!date) return;

        const dKey = date.replace(/-/g, '/');
        if (!dateMap[dKey]) dateMap[dKey] = { date: dKey };

        dateMap[dKey].medium = row['媒体'] || row['Medium'] || row['medium'] || '';
        dateMap[dKey].adCost = parseInt(row['費用(円)'] || row['Cost'] || row['cost']) || 0;
        dateMap[dKey].impressions = parseInt(row['表示回数'] || row['Impressions'] || row['impressions']) || 0;
        dateMap[dKey].clicks = parseInt(row['クリック数'] || row['Clicks'] || row['clicks']) || 0;
        dateMap[dKey].ctr = parseFloat(row['CTR(%)'] || row['CTR'] || row['ctr']) || 0;
        dateMap[dKey].cpc = parseInt(row['CPC(円)'] || row['CPC'] || row['cpc']) || 0;
    });

    // 施策ログ
    AppState.data.measures.forEach(row => {
        const date = row['実施日'];
        if (!dateMap[date]) dateMap[date] = { date };
        if (!dateMap[date].measures) dateMap[date].measures = [];
        dateMap[date].measures.push({
            category: row['カテゴリ'],
            target: row['対象'],
            change: row['変更内容'],
            hypothesis: row['狙い・仮説']
        });
    });

    // 日付順にソート
    AppState.mergedData = Object.values(dateMap).sort((a, b) =>
        new Date(a.date) - new Date(b.date)
    );

    // CPA / ROAS 計算
    AppState.mergedData.forEach(d => {
        d.cpa = d.conversions > 0 ? Math.round(d.adCost / d.conversions) : 0;
        d.roas = d.adCost > 0 ? ((d.conversions * 30000) / d.adCost * 100).toFixed(1) : 0; // 仮の売上単価3万円
    });

    // フィルター済みデータを初期化（全期間）
    applyDateFilter();

    // 期間フィルターの初期値を設定
    if (AppState.mergedData.length > 0 && !AppState.dateFilter.startDate) {
        AppState.dateFilter.startDate = AppState.mergedData[0].date;
        AppState.dateFilter.endDate = AppState.mergedData[AppState.mergedData.length - 1].date;
    }
}

/**
 * 期間フィルターを適用してfilteredDataを更新
 */
function applyDateFilter() {
    if (AppState.dateFilter.isActive && AppState.dateFilter.startDate && AppState.dateFilter.endDate) {
        const start = new Date(AppState.dateFilter.startDate);
        const end = new Date(AppState.dateFilter.endDate);
        end.setHours(23, 59, 59); // endDateも含む
        AppState.filteredData = AppState.mergedData.filter(d => {
            const dt = new Date(d.date);
            return dt >= start && dt <= end;
        });
    } else {
        AppState.filteredData = [...AppState.mergedData];
    }
}

/**
 * 指定期間のデータ平均値を計算するユーティリティ
 */
function calculatePeriodMetrics(dataSlice) {
    if (!dataSlice || dataSlice.length === 0) return null;
    const safeNum = (val) => (typeof val === 'number' && !isNaN(val)) ? val : 0;

    const totalSessions = dataSlice.reduce((s, d) => s + safeNum(d.sessions), 0);
    const totalUsers = dataSlice.reduce((s, d) => s + safeNum(d.users), 0);
    const totalConversions = dataSlice.reduce((s, d) => s + safeNum(d.conversions), 0);
    const totalCost = dataSlice.reduce((s, d) => s + safeNum(d.adCost), 0);
    const totalClicks = dataSlice.reduce((s, d) => s + safeNum(d.clicks), 0);
    const avgCVR = dataSlice.reduce((s, d) => s + safeNum(d.cvr), 0) / dataSlice.length;
    const avgEngRate = dataSlice.reduce((s, d) => s + safeNum(d.engagementRate), 0) / dataSlice.length;
    const avgCPA = totalConversions > 0 ? Math.round(totalCost / totalConversions) : 0;
    const avgROAS = totalCost > 0 ? parseFloat(((totalConversions * 30000) / totalCost * 100).toFixed(1)) : 0;

    return {
        days: dataSlice.length,
        totalSessions,
        totalUsers,
        totalConversions,
        totalCost,
        totalClicks,
        avgCVR: parseFloat(avgCVR.toFixed(2)),
        avgEngRate: parseFloat(avgEngRate.toFixed(1)),
        avgCPA,
        avgROAS,
        avgSessionsPerDay: Math.round(totalSessions / dataSlice.length),
        avgConversionsPerDay: parseFloat((totalConversions / dataSlice.length).toFixed(1))
    };
}

/**
 * 2つの期間メトリクスの差分を計算
 */
function calculatePeriodDiff(metricsA, metricsB) {
    if (!metricsA || !metricsB) return null;
    const pctChange = (a, b) => b !== 0 ? parseFloat(((a - b) / b * 100).toFixed(1)) : 0;

    return {
        cvr: pctChange(metricsB.avgCVR, metricsA.avgCVR),
        cpa: pctChange(metricsB.avgCPA, metricsA.avgCPA),
        roas: pctChange(metricsB.avgROAS, metricsA.avgROAS),
        sessions: pctChange(metricsB.avgSessionsPerDay, metricsA.avgSessionsPerDay),
        engRate: pctChange(metricsB.avgEngRate, metricsA.avgEngRate),
        conversions: pctChange(metricsB.avgConversionsPerDay, metricsA.avgConversionsPerDay),
        totalCost: pctChange(metricsB.totalCost, metricsA.totalCost)
    };
}

// ===================================
// 分析ロジック（Fact / Interpretation / Action）
// ===================================

function generateInsights() {
    const data = AppState.filteredData;
    if (data.length < 2) return [];

    const insights = [];
    const safeNum = (val) => (typeof val === 'number' && !isNaN(val)) ? val : 0;

    // 1. 期間内の施策を特定
    const sDate = (AppState.dateFilter.startDate || '').replace(/\//g, '-');
    const eDate = (AppState.dateFilter.endDate || '').replace(/\//g, '-');
    const start = sDate ? new Date(sDate) : new Date();
    const end = eDate ? new Date(eDate) : new Date();

    const measureInRange = [...AppState.data.measures].filter(m => {
        const d = new Date((m['実施日'] || '').replace(/\//g, '-'));
        return d >= start && d <= end;
    }).sort((a, b) => new Date(b['実施日']) - new Date(a['実施日']))[0];

    if (measureInRange) {
        // --- 施策ベースの分析 ---
        const mDate = new Date(measureInRange['実施日']);
        const before = data.filter(d => new Date(d.date) < mDate);
        const after = data.filter(d => new Date(d.date) >= mDate); // 実施日当日を含む

        if (before.length > 0 && after.length > 0) {
            const mBefore = calculatePeriodMetrics(before);
            const mAfter = calculatePeriodMetrics(after);
            const diff = calculatePeriodDiff(mBefore, mAfter);

            insights.push({
                type: 'fact',
                title: '📊 施策分析 (Fact)',
                content: `実施施策: <strong>${measureInRange['対象']}</strong> (${measureInRange['実施日']})<br>
                実施前後の比較で、CVRは${mBefore.avgCVR}%から${mAfter.avgCVR}%へ(<strong>${diff.cvr > 0 ? '+' : ''}${diff.cvr}%</strong>)、
                CPAは¥${mBefore.avgCPA.toLocaleString()}から¥${mAfter.avgCPA.toLocaleString()}へ(<strong>${diff.cpa}%</strong>)変化しました。`
            });

            insights.push({
                type: 'interpretation',
                title: '🔍 施策の解釈 (Interpretation)',
                content: `${measureInRange['対象']}における「${measureInRange['変更内容']}」は、${diff.cvr > 0 ? '成約率の向上に寄与しており、当初の狙い通り' : '成約率への影響は限定的であり、ユーザーへの訴求力が不足している'}と考えられます。
                広告費のWoW ${diff.totalCost}%に対し、獲得数WoW ${diff.conversions}%となっており、${diff.cpa < 0 ? '獲得効率は改善' : '獲得効率は低下'}傾向にあります。`
            });
        }
    }

    const mNow = calculatePeriodMetrics(data);

    // 2. 期間比較分析 (PoP)
    const periodDays = data.length;

    const prevStart = new Date(start);
    prevStart.setDate(prevStart.getDate() - periodDays);
    const prevEnd = new Date(start);
    prevEnd.setDate(prevEnd.getDate() - 1);

    const prevData = AppState.mergedData.filter(d => {
        const dt = new Date(d.date.replace(/\//g, '-'));
        return dt >= prevStart && dt <= prevEnd;
    });

    if (prevData.length > 0) {
        const mPrev = calculatePeriodMetrics(prevData);
        const diff = calculatePeriodDiff(mPrev, mNow);

        insights.push({
            type: 'fact',
            title: `📈 前期間比分析 (${periodDays}日前比)`,
            content: `現在の期間 (${periodDays}日間) とその前期間を比較すると、セッション数は <strong>${diff.sessions > 0 ? '+' : ''}${diff.sessions}%</strong>、
            CVRは <strong>${diff.cvr > 0 ? '+' : ''}${diff.cvr}%</strong> となりました。
            結果として、1日平均の成約数は ${mPrev.avgConversionsPerDay}件から${mNow.avgConversionsPerDay}件へ推移しています。`
        });

        const isGood = diff.cvr > 0 && diff.cpa < 0;
        insights.push({
            type: 'interpretation',
            title: '💡 状況の解釈',
            content: isGood ?
                `全体として<strong>非常に良好な推移</strong>です。流入数と質のバランスが最適化されており、獲得効率が高い水準で安定しています。` :
                `一部の指標において改善の余地があります。${diff.cvr < 0 ? '流入は増えていますが成約率が低下しており、ターゲット層のズレが発生している可能性があります。' : '効率は維持されていますが、さらなるスケールには流入数の底上げが必要です。'}`
        });
    } else {
        // 前期間がない場合の現状分析
        insights.push({
            type: 'fact',
            title: '📊 現状分析 (Current Facts)',
            content: `選択された期間 (${periodDays}日間) において、
            平均CVR <strong>${mNow.avgCVR}%</strong>、平均CPA <strong>¥${mNow.avgCPA.toLocaleString()}</strong> を記録しています。
            期間中の広告宣伝費は合計 ¥${mNow.totalCost.toLocaleString()} です。`
        });
    }

    // 3. Action (常に生成)
    insights.push({
        type: 'action',
        title: '🎯 次の一手 (Action)',
        content: `<strong>① 運用:</strong> ${mNow && mNow.avgROAS > 100 ? 'ROASが維持されているため、広告予算の維持または微増(5-10%)を推奨。' : 'ROASが目標を下回っているため、低パフォーマンスな媒体の予算削減を検討。'}<br>
        <strong>② 改善:</strong> 直近の数値変動を考慮し、LPの特定要素（見出し、フォームの使いやすさ等）の微調整による検証を継続。<br>
        <strong>③ 計測:</strong> 獲得の質を維持するため、特定媒体からの流入が成約に繋がっているかを詳細に監視。`
    });

    return insights;
}

// ===================================
// KPI計算
// ===================================

function calculateKPIs() {
    const data = AppState.filteredData;
    if (data.length === 0) return null;

    const latest = data[data.length - 1];
    const previous = data.length > 1 ? data[data.length - 2] : null;

    // 安全なプロパティアクセス（undefinedの場合は0にフォールバック）
    const safeNum = (val) => (typeof val === 'number' && !isNaN(val)) ? val : 0;

    const totalConversions = data.reduce((s, d) => s + safeNum(d.conversions), 0);
    const totalCost = data.reduce((s, d) => s + safeNum(d.adCost), 0);
    const totalClicks = data.reduce((s, d) => s + safeNum(d.clicks), 0);
    const totalSessions = data.reduce((s, d) => s + safeNum(d.sessions), 0);
    const avgCVR = data.reduce((s, d) => s + safeNum(d.cvr), 0) / data.length;
    const avgCPA = totalConversions > 0 ? Math.round(totalCost / totalConversions) : 0;
    const avgROAS = totalCost > 0 ? ((totalConversions * 30000) / totalCost * 100).toFixed(1) : '0';

    const latestCVR = safeNum(latest.cvr);
    const latestCPA = safeNum(latest.cpa);
    const latestSessions = safeNum(latest.sessions);
    const latestEngRate = safeNum(latest.engagementRate);

    const cvrChange = previous && safeNum(previous.cvr) > 0 ? ((latestCVR - safeNum(previous.cvr)) / safeNum(previous.cvr) * 100).toFixed(1) : 0;
    const cpaPrevious = previous ? safeNum(previous.cpa) : 0;
    const cpaChange = cpaPrevious > 0 ? ((latestCPA - cpaPrevious) / cpaPrevious * 100).toFixed(1) : 0;

    return {
        cvr: { value: latestCVR.toFixed(2) + '%', change: cvrChange, label: 'CVR（成約率）', tooltip: 'コンバージョン率。サイト訪問者のうち成約に至った割合' },
        cpa: { value: '¥' + latestCPA.toLocaleString(), change: cpaChange, label: 'CPA（顧客獲得単価）', tooltip: '1件の成約を獲得するのにかかった広告費' },
        roas: { value: avgROAS + '%', change: 0, label: 'ROAS（広告費用対効果）', tooltip: '広告費に対する売上の割合。100%以上で黒字' },
        sessions: { value: latestSessions.toLocaleString(), change: previous && safeNum(previous.sessions) > 0 ? ((latestSessions - safeNum(previous.sessions)) / safeNum(previous.sessions) * 100).toFixed(1) : 0, label: 'セッション数', tooltip: 'サイトへの訪問回数' },
        engagement: { value: latestEngRate.toFixed(1) + '%', change: previous && safeNum(previous.engagementRate) > 0 ? ((latestEngRate - safeNum(previous.engagementRate)) / safeNum(previous.engagementRate) * 100).toFixed(1) : 0, label: 'エンゲージメント率', tooltip: 'サイト内で積極的に行動したユーザーの割合' }
    };
}

// ===================================
// チャート描画
// ===================================

function renderCharts() {
    // 既存チャートを破棄
    Object.values(AppState.charts).forEach(chart => chart.destroy());
    AppState.charts = {};

    const data = AppState.filteredData;
    if (data.length === 0) return;

    const labels = data.map(d => {
        const date = new Date(d.date);
        return `${date.getMonth() + 1}/${date.getDate()}`;
    });

    // 施策実施日のアノテーション
    const measureAnnotations = {};
    AppState.data.measures.forEach((m, idx) => {
        const date = new Date(m['実施日']);
        const label = `${date.getMonth() + 1}/${date.getDate()}`;
        const labelIdx = labels.indexOf(label);
        if (labelIdx >= 0) {
            measureAnnotations[`line${idx}`] = {
                type: 'line',
                xMin: labelIdx,
                xMax: labelIdx,
                borderColor: 'rgba(251, 191, 36, 0.6)',
                borderWidth: 2,
                borderDash: [6, 4],
                label: {
                    display: true,
                    content: m['カテゴリ'] === 'LP改善' ? '🔧 LP改善' : '💰 予算変更',
                    position: 'start',
                    backgroundColor: 'rgba(251, 191, 36, 0.15)',
                    color: '#fbbf24',
                    font: { size: 10, weight: '600', family: "'Inter', 'Noto Sans JP', sans-serif" },
                    padding: { top: 4, bottom: 4, left: 8, right: 8 },
                    borderRadius: 6,
                }
            };
        }
    });

    // 共通チャートオプション
    const commonOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                labels: {
                    color: '#94a3b8',
                    font: { family: "'Inter', 'Noto Sans JP', sans-serif", size: 11 },
                    padding: 16,
                    usePointStyle: true,
                    pointStyleWidth: 8,
                }
            },
            tooltip: {
                backgroundColor: 'rgba(15, 23, 42, 0.95)',
                titleColor: '#f1f5f9',
                bodyColor: '#cbd5e1',
                borderColor: 'rgba(99, 102, 241, 0.3)',
                borderWidth: 1,
                padding: 12,
                cornerRadius: 10,
                titleFont: { family: "'Inter', 'Noto Sans JP', sans-serif", weight: '600' },
                bodyFont: { family: "'Inter', 'Noto Sans JP', sans-serif" },
            }
        },
        scales: {
            x: {
                grid: { color: 'rgba(51, 65, 85, 0.3)', drawBorder: false },
                ticks: { color: '#64748b', font: { family: "'Inter', 'Noto Sans JP', sans-serif", size: 11 } }
            },
            y: {
                grid: { color: 'rgba(51, 65, 85, 0.3)', drawBorder: false },
                ticks: { color: '#64748b', font: { family: "'Inter', 'Noto Sans JP', sans-serif", size: 11 } }
            }
        },
        interaction: {
            mode: 'index',
            intersect: false,
        }
    };

    // --- CVR推移チャート ---
    const cvrCtx = document.getElementById('chart-cvr');
    if (cvrCtx) {
        AppState.charts.cvr = new Chart(cvrCtx, {
            type: 'line',
            data: {
                labels,
                datasets: [{
                    label: 'CVR (%)',
                    data: data.map(d => d.cvr),
                    borderColor: '#6366f1',
                    backgroundColor: 'rgba(99, 102, 241, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    pointBackgroundColor: '#6366f1',
                    pointBorderColor: '#1e293b',
                    pointBorderWidth: 3,
                    pointHoverRadius: 8,
                }]
            },
            options: {
                ...commonOptions,
                plugins: {
                    ...commonOptions.plugins,
                    annotation: { annotations: measureAnnotations }
                },
                scales: {
                    ...commonOptions.scales,
                    y: {
                        ...commonOptions.scales.y,
                        ticks: {
                            ...commonOptions.scales.y.ticks,
                            callback: v => v + '%'
                        }
                    }
                }
            }
        });
    }

    // --- CPA推移チャート ---
    const cpaCtx = document.getElementById('chart-cpa');
    if (cpaCtx) {
        AppState.charts.cpa = new Chart(cpaCtx, {
            type: 'line',
            data: {
                labels,
                datasets: [{
                    label: 'CPA (円)',
                    data: data.map(d => d.cpa),
                    borderColor: '#f59e0b',
                    backgroundColor: 'rgba(245, 158, 11, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    pointBackgroundColor: '#f59e0b',
                    pointBorderColor: '#1e293b',
                    pointBorderWidth: 3,
                    pointHoverRadius: 8,
                }]
            },
            options: {
                ...commonOptions,
                plugins: {
                    ...commonOptions.plugins,
                    annotation: { annotations: measureAnnotations }
                },
                scales: {
                    ...commonOptions.scales,
                    y: {
                        ...commonOptions.scales.y,
                        ticks: {
                            ...commonOptions.scales.y.ticks,
                            callback: v => '¥' + v.toLocaleString()
                        }
                    }
                }
            }
        });
    }

    // --- セッション×コンバージョン複合チャート ---
    const comboCtx = document.getElementById('chart-combo');
    if (comboCtx) {
        AppState.charts.combo = new Chart(comboCtx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    {
                        label: 'セッション数',
                        data: data.map(d => d.sessions),
                        backgroundColor: 'rgba(99, 102, 241, 0.4)',
                        borderColor: '#6366f1',
                        borderWidth: 1,
                        borderRadius: 6,
                        type: 'bar',
                        yAxisID: 'y',
                    },
                    {
                        label: '成約数',
                        data: data.map(d => d.conversions),
                        borderColor: '#10b981',
                        backgroundColor: 'rgba(16, 185, 129, 0.1)',
                        borderWidth: 3,
                        fill: false,
                        tension: 0.4,
                        type: 'line',
                        yAxisID: 'y1',
                        pointRadius: 6,
                        pointBackgroundColor: '#10b981',
                        pointBorderColor: '#1e293b',
                        pointBorderWidth: 3,
                    }
                ]
            },
            options: {
                ...commonOptions,
                plugins: {
                    ...commonOptions.plugins,
                    annotation: { annotations: measureAnnotations }
                },
                scales: {
                    x: commonOptions.scales.x,
                    y: {
                        ...commonOptions.scales.y,
                        position: 'left',
                        title: {
                            display: true,
                            text: 'セッション数',
                            color: '#64748b',
                            font: { size: 11, family: "'Inter', 'Noto Sans JP', sans-serif" }
                        }
                    },
                    y1: {
                        ...commonOptions.scales.y,
                        position: 'right',
                        grid: { drawOnChartArea: false },
                        title: {
                            display: true,
                            text: '成約数',
                            color: '#64748b',
                            font: { size: 11, family: "'Inter', 'Noto Sans JP', sans-serif" }
                        }
                    }
                }
            }
        });
    }

    // --- 広告費 vs ROAS ---
    const costCtx = document.getElementById('chart-cost');
    if (costCtx) {
        AppState.charts.cost = new Chart(costCtx, {
            type: 'bar',
            data: {
                labels,
                datasets: [
                    {
                        label: '広告費 (円)',
                        data: data.map(d => d.adCost),
                        backgroundColor: 'rgba(236, 72, 153, 0.4)',
                        borderColor: '#ec4899',
                        borderWidth: 1,
                        borderRadius: 6,
                        yAxisID: 'y',
                    },
                    {
                        label: 'ROAS (%)',
                        data: data.map(d => parseFloat(d.roas)),
                        borderColor: '#8b5cf6',
                        backgroundColor: 'rgba(139, 92, 246, 0.1)',
                        borderWidth: 3,
                        fill: false,
                        tension: 0.4,
                        type: 'line',
                        yAxisID: 'y1',
                        pointRadius: 6,
                        pointBackgroundColor: '#8b5cf6',
                        pointBorderColor: '#1e293b',
                        pointBorderWidth: 3,
                    }
                ]
            },
            options: {
                ...commonOptions,
                scales: {
                    x: commonOptions.scales.x,
                    y: {
                        ...commonOptions.scales.y,
                        position: 'left',
                        ticks: {
                            ...commonOptions.scales.y.ticks,
                            callback: v => '¥' + (v / 1000).toFixed(0) + 'K'
                        },
                        title: {
                            display: true,
                            text: '広告費',
                            color: '#64748b',
                            font: { size: 11, family: "'Inter', 'Noto Sans JP', sans-serif" }
                        }
                    },
                    y1: {
                        ...commonOptions.scales.y,
                        position: 'right',
                        grid: { drawOnChartArea: false },
                        ticks: {
                            ...commonOptions.scales.y.ticks,
                            callback: v => v + '%'
                        },
                        title: {
                            display: true,
                            text: 'ROAS',
                            color: '#64748b',
                            font: { size: 11, family: "'Inter', 'Noto Sans JP', sans-serif" }
                        }
                    }
                }
            }
        });
    }
}

// ===================================
// PDF生成
// ===================================

async function generatePDF() {
    const content = document.getElementById('pdf-capture-area');
    if (!content) return;

    // PDF生成中オーバーレイ表示
    const overlay = document.createElement('div');
    overlay.className = 'pdf-generating-overlay';
    overlay.innerHTML = '<div class="loading-spinner"></div><div class="pdf-generating-text">PDFを生成中...</div>';
    document.body.appendChild(overlay);

    // PDF用のスタイルに切り替え
    document.body.classList.add('pdf-mode');

    try {
        await new Promise(resolve => setTimeout(resolve, 500));

        const canvas = await html2canvas(content, {
            scale: 2,
            useCORS: true,
            backgroundColor: '#ffffff',
            logging: false,
            windowWidth: 1200,
        });

        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({
            orientation: 'portrait',
            unit: 'mm',
            format: 'a4'
        });

        const pageWidth = 210;
        const pageHeight = 297;
        const margin = 10;
        const contentWidth = pageWidth - (margin * 2);

        const imgData = canvas.toDataURL('image/png');
        const imgWidth = contentWidth;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        // ヘッダー
        pdf.setFillColor(79, 70, 229);
        pdf.rect(0, 0, pageWidth, 12, 'F');
        pdf.setTextColor(255, 255, 255);
        pdf.setFontSize(9);
        pdf.text('マーケティング分析レポート', margin, 8);
        pdf.text(new Date().toLocaleDateString('ja-JP'), pageWidth - margin, 8, { align: 'right' });

        // コンテンツを複数ページに分割
        let currentY = 16;
        let remainingHeight = imgHeight;
        let sourceY = 0;

        const maxContentHeight = pageHeight - 24; // ヘッダー/フッター分引く

        while (remainingHeight > 0) {
            const heightOnPage = Math.min(remainingHeight, maxContentHeight);

            pdf.addImage(imgData, 'PNG', margin, currentY, imgWidth, imgHeight, undefined, 'FAST', 0);

            // フッター
            pdf.setFillColor(241, 245, 249);
            pdf.rect(0, pageHeight - 8, pageWidth, 8, 'F');
            pdf.setTextColor(100, 116, 139);
            pdf.setFontSize(7);
            pdf.text('機密 | 事業所向けマーケティング・ダッシュボード', pageWidth / 2, pageHeight - 3, { align: 'center' });

            if (remainingHeight > maxContentHeight) {
                pdf.addPage();
                currentY = 8;
                remainingHeight -= maxContentHeight;
                sourceY += maxContentHeight;
            } else {
                break;
            }
        }

        pdf.save(`marketing_report_${new Date().toISOString().split('T')[0]}.pdf`);

    } catch (error) {
        console.error('PDF生成エラー:', error);
        alert('PDF生成中にエラーが発生しました。');
    } finally {
        document.body.classList.remove('pdf-mode');
        overlay.remove();
    }
}

// ===================================
// 認証
// ===================================

/**
 * Google Identity Services の初期化
 */
function initGoogleLogin() {
    if (typeof google === 'undefined') {
        setTimeout(initGoogleLogin, 100);
        return;
    }

    google.accounts.id.initialize({
        client_id: GOOGLE_CLIENT_ID,
        callback: handleCredentialResponse,
        auto_select: false,
        cancel_on_tap_outside: true
    });

    // ログイン画面が表示されている場合のみボタンをレンダリング
    const btnContainer = document.getElementById('google-login-btn-container');
    if (btnContainer) {
        google.accounts.id.renderButton(btnContainer, {
            theme: 'outline',
            size: 'large',
            width: 320,
            text: 'signin_with',
            shape: 'pill'
        });
    }
}

/**
 * Googleからの証明書レスポンスを処理
 */
function handleCredentialResponse(response) {
    // JWTトークンのデコード（簡易版）
    try {
        const base64Url = response.credential.split('.')[1];
        const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
        const jsonPayload = decodeURIComponent(atob(base64).split('').map(c => {
            return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
        }).join(''));

        const profile = JSON.parse(jsonPayload);

        // ドメイン制限チェック
        if (ALLOWED_DOMAINS.length > 0) {
            const domain = profile.email.split('@')[1];
            if (!ALLOWED_DOMAINS.includes(domain)) {
                alert(`ログイン権限がありません：${profile.email}\n社内アカウント（@${ALLOWED_DOMAINS.join(', @')}）でログインしてください。`);
                return;
            }
        }

        AppState.isAuthenticated = true;
        AppState.user = {
            name: profile.name,
            email: profile.email,
            avatar: profile.picture ? `<img src="${profile.picture}" class="avatar-img" />` : profile.name.charAt(0),
            officeId: 'Authenticated'
        };

        loadClients().then(() => {
            renderApp();
            fetchSheetData();
        });
    } catch (e) {
        console.error('ログイン処理エラー:', e);
        alert('ログイン処理中にエラーが発生しました。');
    }
}

function handleLogout() {
    if (typeof google !== 'undefined') {
        google.accounts.id.disableAutoSelect();
    }
    AppState.isAuthenticated = false;
    AppState.user = null;
    AppState.data = { ga4: [], adCost: [], measures: [] };
    AppState.mergedData = [];
    AppState.filteredData = [];
    Object.values(AppState.charts).forEach(chart => chart.destroy());
    AppState.charts = {};
    renderApp();

    // ログアウト直後にGoogleボタンを再描画
    setTimeout(initGoogleLogin, 100);
}

// ===================================
// UI レンダリング
// ===================================

function renderApp() {
    const app = document.getElementById('app');

    if (AppState.isLoading) {
        app.innerHTML = `
            <div class="bg-particles"></div>
            <div class="loading-overlay">
                <div class="loading-spinner"></div>
                <div class="loading-text">スプレッドシートからデータを取得中...</div>
            </div>
        `;
        return;
    }

    if (!AppState.isAuthenticated) {
        app.innerHTML = renderLoginScreen();
        return;
    }

    app.innerHTML = renderDashboard();

    // モーダルをDOMに追加（renderDashboardの後）
    if (AppState.showClientModal) {
        const modal = document.createElement('div');
        modal.innerHTML = renderClientModal();
        document.body.appendChild(modal.firstElementChild);
    }
    if (AppState.showClientManager) {
        const panel = document.createElement('div');
        panel.innerHTML = renderClientManagerPanel();
        document.body.appendChild(panel.firstElementChild);
    }
}

function renderLoginScreen() {
    setTimeout(initGoogleLogin, 100); // UI描画後にGoogleボタンを初期化
    return `
        <div class="bg-particles"></div>
        <div class="login-container">
            <div class="login-card">
                <div style="font-size: 2.5rem; margin-bottom: 16px;">📊</div>
                <h1>マーケティング・ダッシュボード</h1>
                <p>広告代理店向け WEB広告効果 × LP改善施策 分析レポート</p>
                
                <div id="google-login-btn-container" style="display: flex; justify-content: center; margin: 32px 0;">
                    <!-- Googleボタンがここに自動生成されます -->
                </div>

                <p style="margin-top: 24px; font-size: 0.7rem; color: #475569;">
                    自社のアカウント（Google Workspace）のみアクセス可能です
                </p>
            </div>
        </div>
    `;
}

function renderDashboard() {
    const kpis = calculateKPIs();
    const insights = generateInsights();
    const data = AppState.filteredData;

    return `
        <div class="bg-particles"></div>
        <div class="dashboard-wrapper" id="pdf-capture-area">
            ${renderHeader()}
            <div class="dashboard-content">
                ${AppState.usedFallback ? `
                <div style="background: rgba(245, 158, 11, 0.1); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 10px; padding: 12px 18px; margin-bottom: 16px; display: flex; align-items: center; gap: 10px; font-size: 0.8rem; color: #fbbf24; animation: fadeIn 0.5s ease-out;">
                    <span>⚠️</span>
                    <span>Google Sheetsへの直接接続ができなかったため、取得済みのデータを表示しています。Webサーバー経由でアクセスするとリアルタイムデータを取得できます。</span>
                </div>
                ` : ''}

                <!-- 期間フィルター -->
                ${renderDateFilterPanel()}

                <!-- タブ切り替え -->
                <div class="tab-bar">
                    <button class="tab-btn ${AppState.activeTab === 'overview' ? 'active' : ''}" onclick="switchTab('overview')" id="tab-overview">📊 概要</button>
                    <button class="tab-btn ${AppState.activeTab === 'detail' ? 'active' : ''}" onclick="switchTab('detail')" id="tab-detail">📋 詳細データ</button>
                    <button class="tab-btn ${AppState.activeTab === 'analysis' ? 'active' : ''}" onclick="switchTab('analysis')" id="tab-analysis">🧠 分析レポート</button>
                    <button class="tab-btn ${AppState.activeTab === 'compare' ? 'active' : ''}" onclick="switchTab('compare')" id="tab-compare">⚖️ 期間比較</button>
                </div>

                ${AppState.activeTab === 'overview' ? renderOverviewTab(kpis, data) : ''}
                ${AppState.activeTab === 'detail' ? renderDetailTab(data) : ''}
                ${AppState.activeTab === 'analysis' ? renderAnalysisTab(insights) : ''}
                ${AppState.activeTab === 'compare' ? renderCompareTab() : ''}
            </div>
            ${renderFooter()}
        </div>
    `;
}

/**
 * 期間フィルターパネルのUI
 */
function renderDateFilterPanel() {
    // 日付をinput[type=date]用のyyyy-mm-dd形式に変換
    const toInputDate = (dStr) => {
        if (!dStr) return '';
        const d = new Date(dStr.replace(/-/g, '/'));
        if (isNaN(d.getTime())) return '';
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const r = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${r}`;
    };
    const minDate = AppState.mergedData.length > 0 ? toInputDate(AppState.mergedData[0].date) : '';
    const maxDate = AppState.mergedData.length > 0 ? toInputDate(AppState.mergedData[AppState.mergedData.length - 1].date) : '';

    return `
        <div class="date-filter-panel">
            <div class="date-filter-header">
                <span class="date-filter-icon">📅</span>
                <span class="date-filter-title">分析期間</span>
                <span class="date-filter-count">${AppState.filteredData.length} 日間のデータ</span>
            </div>
            <div class="date-filter-controls">
                <div class="date-input-group">
                    <label>開始日</label>
                    <input type="date" id="filter-start-date" value="${toInputDate(AppState.dateFilter.startDate)}" min="${minDate}" max="${maxDate}" onchange="updateDateFilter()">
                </div>
                <span class="date-separator">〜</span>
                <div class="date-input-group">
                    <label>終了日</label>
                    <input type="date" id="filter-end-date" value="${toInputDate(AppState.dateFilter.endDate)}" min="${minDate}" max="${maxDate}" onchange="updateDateFilter()">
                </div>
                <div class="date-filter-actions">
                    <button class="date-preset-btn" onclick="setDatePreset('all')" id="preset-all">全期間</button>
                    <button class="date-preset-btn" onclick="setDatePreset('last7')" id="preset-7d">直近7日</button>
                    <button class="date-preset-btn" onclick="setDatePreset('last14')" id="preset-14d">直近14日</button>
                    
                    <select class="month-selector" onchange="setDatePreset('month', this.value)">
                        <option value="">年月を選択</option>
                        ${(() => {
            const months = new Set();
            AppState.mergedData.forEach(d => {
                const dt = new Date(d.date.replace(/-/g, '/'));
                if (!isNaN(dt.getTime())) {
                    months.add(`${dt.getFullYear()}/${String(dt.getMonth() + 1).padStart(2, '0')}`);
                }
            });
            return Array.from(months)
                .sort((a, b) => b.localeCompare(a))
                .slice(0, 13)
                .map(m => `<option value="${m}">${m.replace('/', '年')}月</option>`)
                .join('');
        })()}
                    </select>
                </div>
            </div>
        </div>
    `;
}

function renderHeader() {
    const client = getActiveClient();
    const clientOptions = AppState.clients.map(c =>
        `<option value="${c.id}" ${c.id === AppState.activeClientId ? 'selected' : ''}>${c.name} / ${c.siteName}</option>`
    ).join('');

    return `
        <header class="header">
            <div class="header-left">
                <div class="header-logo">M</div>
                <div>
                    <div class="header-title">マーケティング・ダッシュボード</div>
                    <div class="header-subtitle">広告代理店向け 分析ツール</div>
                </div>
            </div>
            <div class="header-center">
                <div class="client-selector">
                    <label class="client-selector-label">🏢 顧客:</label>
                    <select class="client-dropdown" id="client-dropdown" onchange="switchClient(this.value)">
                        ${clientOptions}
                    </select>
                    <button class="btn-add-client" onclick="openClientModal()" title="新規顧客を登録">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><line x1="12" y1="5" x2="12" y2="19"></line><line x1="5" y1="12" x2="19" y2="12"></line></svg>
                    </button>
                    <button class="btn-manage-clients" onclick="openClientManager()" title="顧客管理">
                        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="3"></circle><path d="M12 1v6m0 6v6m-7-7h6m6 0h6"></path></svg>
                        管理
                    </button>
                </div>
            </div>
            <div class="header-right">
                <div class="header-date">
                    📅 ${AppState.mergedData.length > 0 ? AppState.mergedData[0].date + ' 〜 ' + AppState.mergedData[AppState.mergedData.length - 1].date : '---'}
                </div>
                <button class="btn-refresh" onclick="refreshData()" id="btn-refresh">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>
                    更新
                </button>
                <button class="btn-pdf" onclick="generatePDF()" id="btn-pdf-export">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
                    PDF
                </button>
                <button class="btn-json" onclick="exportJSON()" id="btn-json-export" title="JSON形式でデータをエクスポート">
                    <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 16 12 12 8 16"></polyline><line x1="12" y1="12" x2="12" y2="21"></line><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"></path><polyline points="16 16 12 12 8 16"></polyline></svg>
                    JSON
                </button>
                <div class="user-avatar" title="${AppState.user?.name || ''}">${AppState.user?.avatar || '?'}</div>
                <button class="btn-logout" onclick="handleLogout()" id="btn-logout">ログアウト</button>
            </div>
        </header>
    `;
}

function renderOverviewTab(kpis, data) {
    if (!kpis) return '<p style="text-align:center;color:#94a3b8;padding:60px;">データが読み込まれていません。</p>';

    return `
        <!-- KPI カード -->
        <div class="kpi-grid">
            ${renderKPICard(kpis.cvr)}
            ${renderKPICard(kpis.cpa)}
            ${renderKPICard(kpis.roas)}
            ${renderKPICard(kpis.sessions)}
            ${renderKPICard(kpis.engagement)}
        </div>

        <!-- チャート -->
        <div class="chart-grid">
            <div class="chart-card" style="animation-delay: 0.2s;">
                <h3><span class="icon">📈</span> CVR（成約率）推移</h3>
                <div class="chart-container">
                    <canvas id="chart-cvr"></canvas>
                </div>
            </div>
            <div class="chart-card" style="animation-delay: 0.3s;">
                <h3><span class="icon">💰</span> CPA（顧客獲得単価）推移</h3>
                <div class="chart-container">
                    <canvas id="chart-cpa"></canvas>
                </div>
            </div>
            <div class="chart-card" style="animation-delay: 0.4s;">
                <h3><span class="icon">👥</span> セッション数 × 成約数</h3>
                <div class="chart-container">
                    <canvas id="chart-combo"></canvas>
                </div>
            </div>
            <div class="chart-card" style="animation-delay: 0.5s;">
                <h3><span class="icon">📊</span> 広告費 × ROAS</h3>
                <div class="chart-container">
                    <canvas id="chart-cost"></canvas>
                </div>
            </div>
        </div>

        <!-- 施策タイムライン -->
        ${renderTimeline()}
    `;
}

function renderKPICard(kpi) {
    const isPositive = parseFloat(kpi.change) > 0;
    const isCPA = kpi.label.includes('CPA');
    // CPAは下がる方が良いので色を反転
    const changeClass = isCPA ? (isPositive ? 'negative' : 'positive') : (isPositive ? 'positive' : 'negative');
    const changeValue = parseFloat(kpi.change);

    return `
        <div class="kpi-card">
            <div class="kpi-label">
                <span class="tooltip-trigger" data-tooltip="${kpi.tooltip}">${kpi.label}</span>
            </div>
            <div class="kpi-value">${kpi.value}</div>
            ${changeValue !== 0 ? `
                <span class="kpi-change ${changeClass}">
                    ${changeValue > 0 ? '↑' : '↓'} ${Math.abs(changeValue)}%
                </span>
            ` : '<span class="kpi-change" style="color:#94a3b8;">— 期間平均</span>'}
        </div>
    `;
}

function renderTimeline() {
    if (AppState.data.measures.length === 0) return '';

    const items = AppState.data.measures.map(m => `
        <div class="timeline-item">
            <div class="timeline-content">
                <div class="timeline-date">${m['実施日']}</div>
                <span class="timeline-category ${m['カテゴリ'] === 'LP改善' ? 'lp' : 'ad'}">${m['カテゴリ']}</span>
                <div class="timeline-title">${m['対象']} — ${m['変更内容']}</div>
                <div class="timeline-desc">💡 ${m['狙い・仮説']}</div>
            </div>
        </div>
    `).join('');

    return `
        <div class="timeline-section">
            <h2>🔄 施策タイムライン</h2>
            <div class="timeline">
                ${items}
            </div>
        </div>
    `;
}

function renderDetailTab(data) {
    if (!data || data.length === 0) return '<p style="text-align:center;color:#94a3b8;padding:60px;">選択された期間にデータがありません。期間を変更してください。</p>';

    const rows = data.map((d, i) => {
        const prev = i > 0 ? data[i - 1] : null;
        const cvrChange = prev ? ((d.cvr - prev.cvr) / prev.cvr * 100).toFixed(1) : '-';
        const hasMeasure = d.measures && d.measures.length > 0;

        return `
            <tr class="${hasMeasure ? 'highlight-row' : ''}">
                <td style="font-weight:600;">${d.date} ${hasMeasure ? '⚡' : ''}</td>
                <td>${d.sessions.toLocaleString()}</td>
                <td>${d.users.toLocaleString()}</td>
                <td>${d.engagementRate.toFixed(1)}%</td>
                <td>${d.clicks.toLocaleString()}</td>
                <td>¥${d.adCost.toLocaleString()}</td>
                <td style="font-weight:700;">${d.cvr.toFixed(2)}%
                    ${cvrChange !== '-' ? `<span class="change-badge ${parseFloat(cvrChange) >= 0 ? 'up' : 'down'}">${parseFloat(cvrChange) >= 0 ? '↑' : '↓'}${Math.abs(cvrChange)}%</span>` : ''}
                </td>
                <td>¥${d.cpa.toLocaleString()}</td>
                <td>${d.roas}%</td>
                <td style="max-width:200px;font-size:0.75rem;color:#94a3b8;">${hasMeasure ? d.measures.map(m => m.change).join(', ') : '—'}</td>
            </tr>
        `;
    }).join('');

    return `
        <div class="data-table-section">
            <h2>📋 統合データテーブル（日付キーで統合済）</h2>
            <div class="table-wrapper">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>日付</th>
                            <th><span class="tooltip-trigger" data-tooltip="サイトへの訪問回数">セッション</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="ユニークユーザー数">ユーザー</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="積極的に行動した訪問者の割合">ER</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="広告のクリック数">クリック</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="広告に使った費用">広告費</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="成約率（コンバージョン率）">CVR</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="顧客獲得単価">CPA</span></th>
                            <th><span class="tooltip-trigger" data-tooltip="広告費用対効果（100%以上で黒字）">ROAS</span></th>
                            <th>施策メモ</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${rows}
                    </tbody>
                </table>
            </div>
            <p style="margin-top:12px;font-size:0.75rem;color:#64748b;">⚡ = 施策実施日 | ER = エンゲージメント率 | ROAS算出は売上単価30,000円で仮計算</p>
        </div>
    `;
}

function renderAnalysisTab(insights) {
    if (insights.length === 0) {
        return '<p style="text-align:center;color:#94a3b8;padding:60px;">十分なデータがないため分析を生成できません。</p>';
    }

    const insightCards = insights.map(ins => `
        <div class="insight-card">
            <div class="insight-tag ${ins.type}">${ins.title}</div>
            <div class="insight-content">${ins.content}</div>
        </div>
    `).join('');

    return `
        <div class="insight-section">
            <h2>🧠 AI分析レポート — Fact / Interpretation / Action</h2>
            <p style="font-size:0.8rem;color:#94a3b8;margin-bottom:24px;margin-top:-12px;">
                施策実施日前後のデータ変化を自動分析し、3段構成でインサイトを生成しています。
            </p>
            ${insightCards}
        </div>

        <!-- 施策効果サマリー -->
        <div class="insight-section">
            <h2>📊 施策効果サマリー</h2>
            ${renderMeasureEffectSummary()}
        </div>
    `;
}

function renderMeasureEffectSummary() {
    const data = AppState.filteredData;
    const allMeasures = AppState.data.measures;
    if (data.length < 2) return '<p style="text-align:center;color:#64748b;padding:20px;">分析対象のデータが不足しています。</p>';

    // 期間内の施策を特定、なければ直近の過去施策を1つ
    const start = new Date(AppState.dateFilter.startDate);
    const end = new Date(AppState.dateFilter.endDate);

    let targetMeasure = [...allMeasures].filter(m => {
        const d = new Date(m['実施日']);
        return d >= start && d <= end;
    }).sort((a, b) => new Date(b['実施日']) - new Date(a['実施日']))[0];

    // 期間内に施策がない場合は、期間前の直近の施策を探す
    if (!targetMeasure) {
        targetMeasure = [...allMeasures].filter(m => new Date(m['実施日']) <= end)
            .sort((a, b) => new Date(b['実施日']) - new Date(a['実施日']))[0];
    }

    if (!targetMeasure) return '<p style="text-align:center;color:#64748b;padding:20px;">適用可能な施策データが見つかりません。</p>';

    const measureDate = new Date(targetMeasure['実施日']);

    // 施策前後のデータ抽出（AppState.mergedData全域から探す）
    const beforeData = AppState.mergedData.filter(d => new Date(d.date) < measureDate);
    const afterData = AppState.mergedData.filter(d => new Date(d.date) >= measureDate); // 当日含む

    if (beforeData.length === 0 || afterData.length === 0) {
        return `<p style="text-align:center;color:#64748b;padding:20px;">施策日(${targetMeasure['実施日']})の前後データが不足しているため比較できません。</p>`;
    }

    const metrics = [
        {
            label: 'CVR（成約率）',
            before: (beforeData.reduce((s, d) => s + d.cvr, 0) / beforeData.length).toFixed(2) + '%',
            after: (afterData.reduce((s, d) => s + d.cvr, 0) / afterData.length).toFixed(2) + '%',
            change: (((afterData.reduce((s, d) => s + d.cvr, 0) / afterData.length) - (beforeData.reduce((s, d) => s + d.cvr, 0) / beforeData.length)) / (beforeData.reduce((s, d) => s + d.cvr, 0) / beforeData.length) * 100).toFixed(1),
            isInverse: false
        },
        {
            label: 'CPA（顧客獲得単価）',
            before: '¥' + Math.round(beforeData.reduce((s, d) => s + d.cpa, 0) / beforeData.length).toLocaleString(),
            after: '¥' + Math.round(afterData.reduce((s, d) => s + d.cpa, 0) / afterData.length).toLocaleString(),
            change: (((afterData.reduce((s, d) => s + d.cpa, 0) / afterData.length) - (beforeData.reduce((s, d) => s + d.cpa, 0) / beforeData.length)) / (beforeData.reduce((s, d) => s + d.cpa, 0) / beforeData.length) * 100).toFixed(1),
            isInverse: true
        },
        {
            label: '成約数 (1日平均)',
            before: (beforeData.reduce((s, d) => s + d.conversions, 0) / beforeData.length).toFixed(1),
            after: (afterData.reduce((s, d) => s + d.conversions, 0) / afterData.length).toFixed(1),
            change: (((afterData.reduce((s, d) => s + d.conversions, 0) / afterData.length) - (beforeData.reduce((s, d) => s + d.conversions, 0) / beforeData.length)) / (beforeData.reduce((s, d) => s + d.conversions, 0) / beforeData.length) * 100).toFixed(1),
            isInverse: false
        }
    ];

    const rows = metrics.map(m => {
        const chg = parseFloat(m.change);
        const isGood = m.isInverse ? chg < 0 : chg > 0;
        return `
            <tr>
                <td style="font-weight:600;">${m.label}</td>
                <td style="color:#94a3b8;">${m.before}</td>
                <td style="color:#e2e8f0;font-weight:600;">${m.after}</td>
                <td>
                    <span class="change-badge ${isGood ? 'up' : 'down'}">
                        ${chg > 0 ? '↑' : '↓'} ${Math.abs(chg)}%
                    </span>
                </td>
            </tr>
        `;
    }).join('');

    return `
        <div class="table-wrapper">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>指標</th>
                        <th>施策前平均</th>
                        <th>施策後平均</th>
                        <th>変化率</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>
        <p style="margin-top:12px;font-size:0.75rem;color:#64748b;">
            📌 最新の施策: <strong>${targetMeasure['対象']} (${targetMeasure['実施日']})</strong> を基準に分析
        </p>
    `;
}

function renderFooter() {
    return `
        <footer class="footer">
            <p>© 2026 マーケティング・ダッシュボード — 機密情報 | 事業所内利用限定</p>
            <p style="margin-top:4px;">データソース: Google Sheets API | 分析ロジック: WoW%(前週比)計算</p>
        </footer>
    `;
}

// ===================================
// 期間比較タブ
// ===================================

function renderCompareTab() {
    const toInputDate = (dStr) => {
        if (!dStr) return '';
        const d = new Date(dStr.replace(/-/g, '/'));
        if (isNaN(d.getTime())) return '';
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const r = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${r}`;
    };
    const minDate = AppState.mergedData.length > 0 ? toInputDate(AppState.mergedData[0].date) : '';
    const maxDate = AppState.mergedData.length > 0 ? toInputDate(AppState.mergedData[AppState.mergedData.length - 1].date) : '';

    const cm = AppState.compareMode;

    // 比較データの計算
    let comparisonHTML = '';
    if (cm.periodA.start && cm.periodA.end && cm.periodB.start && cm.periodB.end) {
        const parse = (s) => s ? new Date(s.replace(/-/g, '/')) : null;
        const startA = parse(cm.periodA.start);
        const endA = parse(cm.periodA.end);
        if (endA) endA.setHours(23, 59, 59, 999);

        const startB = parse(cm.periodB.start);
        const endB = parse(cm.periodB.end);
        if (endB) endB.setHours(23, 59, 59, 999);

        const dataA = AppState.mergedData.filter(d => {
            const dt = new Date(d.date.replace(/-/g, '/'));
            return dt >= startA && dt <= endA;
        });
        const dataB = AppState.mergedData.filter(d => {
            const dt = new Date(d.date.replace(/-/g, '/'));
            return dt >= startB && dt <= endB;
        });

        const metricsA = calculatePeriodMetrics(dataA);
        const metricsB = calculatePeriodMetrics(dataB);

        if (metricsA && metricsB) {
            const diff = calculatePeriodDiff(metricsA, metricsB);
            comparisonHTML = renderComparisonResults(metricsA, metricsB, diff, cm);
        } else {
            comparisonHTML = `<p style="text-align:center;color:#94a3b8;padding:40px;">選択した期間にデータがありません。別の期間を選択してください。</p>`;
        }
    } else {
        comparisonHTML = `
            <div style="text-align:center;color:#94a3b8;padding:60px;">
                <div style="font-size:3rem;margin-bottom:16px;">⚖️</div>
                <p style="font-size:1rem;margin-bottom:8px;">2つの期間を設定してください</p>
                <p style="font-size:0.8rem;">期間Aと期間Bの開始日・終了日を選ぶと、主要指標の比較結果が自動表示されます。</p>
            </div>
        `;
    }

    return `
        <div class="insight-section" style="animation-delay:0.1s;">
            <h2>⚖️ 期間A vs 期間B 比較分析</h2>
            <p style="font-size:0.8rem;color:#94a3b8;margin-bottom:24px;margin-top:-12px;">
                任意の2つの期間を選択し、主要マーケティング指標の変化を比較できます。
            </p>

            <!-- 期間選択UI -->
            <div class="compare-period-selector">
                <div class="compare-period-box period-a">
                    <div class="compare-period-label">📘 期間A（基準）</div>
                    <div class="compare-date-row">
                        <div class="date-input-group">
                            <label>開始日</label>
                            <input type="date" id="compare-a-start" value="${toInputDate(cm.periodA.start)}" min="${minDate}" max="${maxDate}" onchange="updateCompareMode()">
                        </div>
                        <span class="date-separator">〜</span>
                        <div class="date-input-group">
                            <label>終了日</label>
                            <input type="date" id="compare-a-end" value="${toInputDate(cm.periodA.end)}" min="${minDate}" max="${maxDate}" onchange="updateCompareMode()">
                        </div>
                    </div>
                </div>
                <div class="compare-vs">VS</div>
                <div class="compare-period-box period-b">
                    <div class="compare-period-label">📗 期間B（比較対象）</div>
                    <div class="compare-date-row">
                        <div class="date-input-group">
                            <label>開始日</label>
                            <input type="date" id="compare-b-start" value="${toInputDate(cm.periodB.start)}" min="${minDate}" max="${maxDate}" onchange="updateCompareMode()">
                        </div>
                        <span class="date-separator">〜</span>
                        <div class="date-input-group">
                            <label>終了日</label>
                            <input type="date" id="compare-b-end" value="${toInputDate(cm.periodB.end)}" min="${minDate}" max="${maxDate}" onchange="updateCompareMode()">
                        </div>
                    </div>
                </div>
            </div>

            ${comparisonHTML}
        </div>
    `;
}

function renderComparisonResults(metricsA, metricsB, diff, cm) {
    const metrics = [
        { label: 'CVR（成約率）', a: metricsA.avgCVR + '%', b: metricsB.avgCVR + '%', change: diff.cvr, inverse: false, tooltip: 'コンバージョン率の期間平均' },
        { label: 'CPA（顧客獲得単価）', a: '¥' + metricsA.avgCPA.toLocaleString(), b: '¥' + metricsB.avgCPA.toLocaleString(), change: diff.cpa, inverse: true, tooltip: '1件の成約にかかった平均広告費' },
        { label: 'ROAS（広告費用対効果）', a: metricsA.avgROAS + '%', b: metricsB.avgROAS + '%', change: diff.roas, inverse: false, tooltip: '広告費に対する売上割合' },
        { label: 'セッション/日', a: metricsA.avgSessionsPerDay.toLocaleString(), b: metricsB.avgSessionsPerDay.toLocaleString(), change: diff.sessions, inverse: false, tooltip: '1日平均のサイト訪問数' },
        { label: 'エンゲージメント率', a: metricsA.avgEngRate + '%', b: metricsB.avgEngRate + '%', change: diff.engRate, inverse: false, tooltip: '積極的に行動したユーザーの割合' },
        { label: '成約数/日', a: metricsA.avgConversionsPerDay, b: metricsB.avgConversionsPerDay, change: diff.conversions, inverse: false, tooltip: '1日平均の成約件数' }
    ];

    const rows = metrics.map(m => {
        const chg = m.change;
        const isGood = m.inverse ? chg < 0 : chg > 0;
        const arrow = chg > 0 ? '↑' : chg < 0 ? '↓' : '→';
        const badgeClass = chg === 0 ? '' : (isGood ? 'up' : 'down');
        return `
            <tr>
                <td style="font-weight:600;">
                    <span class="tooltip-trigger" data-tooltip="${m.tooltip}">${m.label}</span>
                </td>
                <td style="color:#93c5fd;font-weight:500;">${m.a}</td>
                <td style="color:#86efac;font-weight:500;">${m.b}</td>
                <td>
                    <span class="change-badge ${badgeClass}">
                        ${arrow} ${Math.abs(chg)}%
                    </span>
                </td>
            </tr>
        `;
    }).join('');

    // 期間のラベル生成
    const fmtDate = (d) => { const dt = new Date(d); return `${dt.getMonth() + 1}/${dt.getDate()}`; };
    const labelA = `${fmtDate(cm.periodA.start)}〜${fmtDate(cm.periodA.end)}`;
    const labelB = `${fmtDate(cm.periodB.start)}〜${fmtDate(cm.periodB.end)}`;

    return `
        <div class="table-wrapper" style="margin-top:24px;">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>指標</th>
                        <th style="color:#93c5fd;">📘 期間A（${labelA}）<br><span style="font-weight:400;font-size:0.65rem;">${metricsA.days}日間</span></th>
                        <th style="color:#86efac;">📗 期間B（${labelB}）<br><span style="font-weight:400;font-size:0.65rem;">${metricsB.days}日間</span></th>
                        <th>変化率（A→B）</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows}
                </tbody>
            </table>
        </div>

        <!-- 比較サマリーカード -->
        <div class="compare-summary-cards">
            <div class="compare-summary-card">
                <div class="compare-summary-emoji">💴</div>
                <div class="compare-summary-label">広告費（合計）</div>
                <div class="compare-summary-values">
                    <span style="color:#93c5fd;">A: ¥${metricsA.totalCost.toLocaleString()}</span>
                    <span style="color:#86efac;">B: ¥${metricsB.totalCost.toLocaleString()}</span>
                </div>
            </div>
            <div class="compare-summary-card">
                <div class="compare-summary-emoji">🎯</div>
                <div class="compare-summary-label">成約数（合計）</div>
                <div class="compare-summary-values">
                    <span style="color:#93c5fd;">A: ${metricsA.totalConversions}件</span>
                    <span style="color:#86efac;">B: ${metricsB.totalConversions}件</span>
                </div>
            </div>
            <div class="compare-summary-card">
                <div class="compare-summary-emoji">👥</div>
                <div class="compare-summary-label">ユーザー数（合計）</div>
                <div class="compare-summary-values">
                    <span style="color:#93c5fd;">A: ${metricsA.totalUsers.toLocaleString()}</span>
                    <span style="color:#86efac;">B: ${metricsB.totalUsers.toLocaleString()}</span>
                </div>
            </div>
        </div>

        <p style="margin-top:16px;font-size:0.75rem;color:#64748b;">
            📌 変化率は「期間A → 期間B」の方向で計算。ROAS算出は仮の売上単価30,000円で計算。
        </p>

        <!-- 比較インサイト -->
        <div style="margin-top:32px;">
            <h3>🧠 比較分析インサイト</h3>
            <div class="insight-card" style="margin-top:16px; border-left: 4px solid #6366f1;">
                <div class="insight-content" style="font-size:0.95rem;">
                    ${generateComparisonInsights(metricsA, metricsB, diff)}
                </div>
            </div>
        </div>
    `;
}

/**
 * 2つの期間の差分からAI分析テキストを生成
 */
function generateComparisonInsights(mA, mB, diff) {
    let summary = "";

    // CVRの変化
    if (diff.cvr > 10) {
        summary += `期間Aに対して期間Bの<strong>CVRが${diff.cvr}%向上</strong>しており、非常に高い改善効果が見られます。`;
    } else if (diff.cvr < -10) {
        summary += `期間Aに比べ期間Bでは<strong>CVRが${Math.abs(diff.cvr)}%低下</strong>しています。流入ユーザーの質、もしくはサイト改修による離脱の影響を確認してください。`;
    }

    // CPAの変化
    if (diff.cpa < -10) {
        summary += `獲得効率(CPA)も<strong>${Math.abs(diff.cpa)}%改善</strong>されており、収益性が高まっています。`;
    } else if (diff.cpa > 10) {
        summary += `CPAが<strong>${diff.cpa}%上昇</strong>しています。クリック単価の高騰やCVR低下が原因と考えられます。`;
    }

    // セッションとCVの変化
    if (diff.sessions > 20 && diff.conversions < 5) {
        summary += `流入(セッション)は大幅に増えていますが、成約数に繋がっていません。<strong>集客ターゲットのズレ</strong>が発生している可能性があります。`;
    }

    if (!summary) {
        summary = "主要指標に大きな変動はありません。安定した運用状況ですが、さらなる拡大に向けて新しいターゲット設定やLP改修の余地を検討しましょう。";
    }

    return `
        <p style="margin-bottom:12px;">期間A（${mA.days}日間）と期間B（${mB.days}日間）を比較した結果：</p>
        <p>${summary}</p>
        <div style="margin-top:16px; padding:12px; background:rgba(99,102,241,0.1); border-radius:8px; font-size:0.8rem;">
            <strong>💡 アドバイス:</strong> ${diff.cvr > 0 ? '現在の施策を継続しつつ、獲得単価のさらなる最適化を狙いましょう。' : 'まずは直近の変更点（広告文、LPのファーストビュー等）を見直し、低下原因の仮説を立てることを推奨します。'}
        </div>
    `;
}

// ===================================
// 期間フィルター・比較操作
// ===================================

function updateDateFilter() {
    const startEl = document.getElementById('filter-start-date');
    const endEl = document.getElementById('filter-end-date');
    if (startEl && endEl) {
        AppState.dateFilter.startDate = startEl.value.replace(/-/g, '/');
        AppState.dateFilter.endDate = endEl.value.replace(/-/g, '/');
        AppState.dateFilter.isActive = true;
        applyDateFilter();
        renderApp();
        if (AppState.activeTab === 'overview') {
            setTimeout(() => renderCharts(), 150);
        }
    }
}

function setDatePreset(preset) {
    const allData = AppState.mergedData;
    if (allData.length === 0) return;

    const lastDate = new Date(allData[allData.length - 1].date);
    let startDate;

    switch (preset) {
        case 'all':
            AppState.dateFilter.isActive = false;
            AppState.dateFilter.startDate = allData[0].date;
            AppState.dateFilter.endDate = allData[allData.length - 1].date;
            break;
        case 'last7':
            startDate = new Date(lastDate);
            startDate.setDate(startDate.getDate() - 6);
            AppState.dateFilter.isActive = true;
            AppState.dateFilter.startDate = `${startDate.getFullYear()}/${String(startDate.getMonth() + 1).padStart(2, '0')}/${String(startDate.getDate()).padStart(2, '0')}`;
            AppState.dateFilter.endDate = allData[allData.length - 1].date;
            break;
        case 'last14':
            startDate = new Date(lastDate);
            startDate.setDate(startDate.getDate() - 13);
            AppState.dateFilter.isActive = true;
            AppState.dateFilter.startDate = `${startDate.getFullYear()}/${String(startDate.getMonth() + 1).padStart(2, '0')}/${String(startDate.getDate()).padStart(2, '0')}`;
            AppState.dateFilter.endDate = allData[allData.length - 1].date;
            break;
        case 'last30':
            startDate = new Date(lastDate);
            startDate.setDate(startDate.getDate() - 29);
            AppState.dateFilter.isActive = true;
            AppState.dateFilter.startDate = `${startDate.getFullYear()}/${String(startDate.getMonth() + 1).padStart(2, '0')}/${String(startDate.getDate()).padStart(2, '0')}`;
            AppState.dateFilter.endDate = allData[allData.length - 1].date;
            break;
        case 'month':
            if (!arguments[1]) return;
            const [year, month] = arguments[1].split('/').map(Number);
            const firstDay = new Date(year, month - 1, 1);
            const lastDay = new Date(year, month, 0); // その月の末日

            AppState.dateFilter.isActive = true;
            AppState.dateFilter.startDate = `${year}/${String(month).padStart(2, '0')}/01`;
            AppState.dateFilter.endDate = `${year}/${String(month).padStart(2, '0')}/${String(lastDay.getDate()).padStart(2, '0')}`;
            break;
    }

    applyDateFilter();
    renderApp();
    if (AppState.activeTab === 'overview') {
        setTimeout(() => renderCharts(), 150);
    }
}

function updateCompareMode() {
    const aStart = document.getElementById('compare-a-start');
    const aEnd = document.getElementById('compare-a-end');
    const bStart = document.getElementById('compare-b-start');
    const bEnd = document.getElementById('compare-b-end');

    if (aStart) AppState.compareMode.periodA.start = aStart.value.replace(/-/g, '/');
    if (aEnd) AppState.compareMode.periodA.end = aEnd.value.replace(/-/g, '/');
    if (bStart) AppState.compareMode.periodB.start = bStart.value.replace(/-/g, '/');
    if (bEnd) AppState.compareMode.periodB.end = bEnd.value.replace(/-/g, '/');

    AppState.compareMode.enabled = true;
    renderApp();
}

// ===================================
// 顧客管理UI
// ===================================

function openClientModal() {
    AppState.showClientModal = true;
    renderApp();
}

function closeClientModal() {
    AppState.showClientModal = false;
    // 既存のモーダルDOMを削除
    const modal = document.getElementById('client-modal-overlay');
    if (modal) modal.remove();
}

function openClientManager() {
    AppState.showClientManager = true;
    renderApp();
}

function closeClientManager() {
    AppState.showClientManager = false;
    const panel = document.getElementById('client-manager-overlay');
    if (panel) panel.remove();
}

async function submitNewClient() {
    const name = document.getElementById('new-client-name')?.value;
    const site = document.getElementById('new-client-site')?.value;
    const url = document.getElementById('new-client-url')?.value;

    if (addClient(name, site, url)) {
        closeClientModal();
        await loadClients(); // リストを再読込して反映
        switchClient(AppState.activeClientId);
    }
}

function renderClientModal() {
    return `
        <div class="modal-overlay" id="client-modal-overlay" onclick="if(event.target===this)closeClientModal()">
            <div class="modal-card">
                <div class="modal-header">
                    <h2>➕ 新規顧客登録</h2>
                    <button class="modal-close" onclick="closeClientModal()">✕</button>
                </div>
                <div class="modal-body">
                    <div class="modal-field">
                        <label>🏢 顧客名 <span style="color:#f87171;">*</span></label>
                        <input type="text" id="new-client-name" placeholder="例: ABC商事" />
                    </div>
                    <div class="modal-field">
                        <label>🌐 対象サイト名</label>
                        <input type="text" id="new-client-site" placeholder="例: コーポレートサイト / 採用LP" />
                    </div>
                    <div class="modal-field">
                        <label>📊 スプレッドシートURL <span style="color:#f87171;">*</span></label>
                        <input type="text" id="new-client-url" placeholder="https://docs.google.com/spreadsheets/d/..../edit" />
                        <p class="modal-hint">※ スプレッドシートの共有URLを貼り付けてください。<br>タブ構成は「GA4」「広告費」「施策ログ」の3タブが必要です。</p>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="modal-btn-cancel" onclick="closeClientModal()">キャンセル</button>
                    <button class="modal-btn-submit" onclick="submitNewClient()">登録して分析開始</button>
                </div>
            </div>
        </div>
    `;
}

function renderClientManagerPanel() {
    const clientRows = AppState.clients.map(c => {
        const isDemo = c.id === DEFAULT_CLIENT.id;
        const isActive = c.id === AppState.activeClientId;
        return `
            <div class="client-row ${isActive ? 'active' : ''}">
                <div class="client-row-info">
                    <div class="client-row-name">🏢 ${c.name}</div>
                    <div class="client-row-site">${c.siteName}</div>
                    <div class="client-row-meta">登録: ${c.createdAt} | ID: ${c.spreadsheetId.substring(0, 12)}...</div>
                </div>
                <div class="client-row-actions">
                    ${isActive ? '<span class="client-active-badge">✓ 選択中</span>' : `<button class="client-switch-btn" onclick="closeClientManager();switchClient('${c.id}')">切替</button>`}
                    ${!isDemo ? `<button class="client-delete-btn" onclick="removeClient('${c.id}')">削除</button>` : ''}
                </div>
            </div>
        `;
    }).join('');

    return `
        <div class="modal-overlay" id="client-manager-overlay" onclick="if(event.target===this)closeClientManager()">
            <div class="modal-card" style="max-width:600px;">
                <div class="modal-header">
                    <h2>📋 顧客管理</h2>
                    <button class="modal-close" onclick="closeClientManager()">✕</button>
                </div>
                <div class="modal-body" style="padding:0;">
                    <div class="client-list">
                        ${clientRows}
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="modal-btn-submit" onclick="closeClientManager();openClientModal()">➕ 新規顧客を追加</button>
                </div>
            </div>
        </div>
    `;
}

// ===================================
// ユーティリティ
// ===================================

function switchTab(tab) {
    AppState.activeTab = tab;
    renderApp();
    if (tab === 'overview') {
        setTimeout(() => renderCharts(), 100);
    }
}

async function refreshData() {
    const btn = document.getElementById('btn-refresh');
    if (btn) btn.classList.add('loading');
    await fetchSheetData();
    if (btn) btn.classList.remove('loading');
}

// ===================================
// 初期化
// ===================================
document.addEventListener('DOMContentLoaded', () => {
    renderApp();
});
