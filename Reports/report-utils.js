// OneDrive monitoring report utilities
document.addEventListener('DOMContentLoaded', function() {
    // データテーブルの初期化
    initializeTables();
    // グラフの初期化
    initializeCharts();
});

// テーブル初期化
function initializeTables() {
    const tables = document.querySelectorAll('.data-table');
    tables.forEach(initializeTable);
}

function initializeTable(table) {
    const headers = table.querySelectorAll('th');
    // ソート機能の追加
    headers.forEach((header, index) => {
        header.addEventListener('click', () => sortTable(table, index));
    });

    // エクスポート機能の追加
    addExportButtons(table);
}

// エクスポートボタンの追加
function addExportButtons(table) {
    const container = document.createElement('div');
    container.className = 'export-container';

    // CSVエクスポート
    const csvBtn = createExportButton('csv-export', 'CSVエクスポート', () => 
        exportTableToCSV(table, `onedrive_report_${formatDate(new Date())}.csv`));
    
    // 印刷ボタン
    const printBtn = createExportButton('print-button', '印刷', () => 
        printTable(table));

    container.append(csvBtn, printBtn);
    table.parentNode.insertBefore(container, table);
}

function createExportButton(className, text, onClick) {
    const button = document.createElement('button');
    button.className = className;
    button.textContent = text;
    button.addEventListener('click', onClick);
    return button;
}

// テーブルソート機能
function sortTable(table, column) {
    const tbody = table.querySelector('tbody');
    const rows = Array.from(tbody.querySelectorAll('tr'));
    const header = table.querySelectorAll('th')[column];
    const isAscending = !header.classList.contains('sort-asc');

    // ソート方向の更新
    table.querySelectorAll('th').forEach(th => 
        th.classList.remove('sort-asc', 'sort-desc'));
    header.classList.add(isAscending ? 'sort-asc' : 'sort-desc');

    // データのソート処理
    rows.sort((a, b) => {
        const aValue = a.cells[column].textContent.trim();
        const bValue = b.cells[column].textContent.trim();
        return compareValues(aValue, bValue, isAscending);
    });

    // ソート結果の適用
    tbody.append(...rows);
}

// 値の比較
function compareValues(a, b, ascending) {
    const numA = parseFloat(a);
    const numB = parseFloat(b);
    
    if (!isNaN(numA) && !isNaN(numB)) {
        return ascending ? numA - numB : numB - numA;
    }
    return ascending ? 
        a.localeCompare(b, 'ja') : 
        b.localeCompare(a, 'ja');
}

// CSVエクスポート機能
function exportTableToCSV(table, filename) {
    const rows = Array.from(table.querySelectorAll('tr'));
    const csv = rows.map(row => 
        Array.from(row.querySelectorAll('th,td'))
            .map(cell => `"${cell.textContent.replace(/"/g, '""')}"`)
            .join(',')
    ).join('\n');

    downloadFile(csv, filename, 'text/csv;charset=utf-8;');
}

// ファイルダウンロード
function downloadFile(content, filename, type) {
    const blob = new Blob(['\uFEFF' + content], { type });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
}

// 印刷機能
function printTable(table) {
    const printWindow = window.open('', '_blank');
    const style = `
        <style>
            table { border-collapse: collapse; width: 100%; }
            th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
            th { background-color: #f5f5f5; }
            @media print {
                button { display: none; }
                table { page-break-inside: avoid; }
            }
        </style>
    `;
    printWindow.document.write(`
        <html>
            <head>${style}</head>
            <body>
                ${table.outerHTML}
                <script>window.onload = () => window.print();</script>
            </body>
        </html>
    `);
}

// 日付フォーマット
function formatDate(date) {
    return date.toISOString().slice(0,10).replace(/-/g, '');
}

// グラフ初期化
function initializeCharts() {
    const ctx = document.getElementById('usageChart')?.getContext('2d');
    if (ctx) {
        new Chart(ctx, {
            type: 'bar',
            data: {
                datasets: [{
                    label: 'ストレージ使用量',
                    backgroundColor: '#3498db'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: 'OneDriveストレージ使用状況'
                    }
                }
            }
        });
    }
}

// ローディング表示の制御
function showLoading() {
    document.querySelector('.loading').classList.add('active');
}

function hideLoading() {
    document.querySelector('.loading').classList.remove('active');
}

// チャートの更新
function updateCharts() {
    const charts = Object.values(Chart.instances);
    charts.forEach(chart => {
        chart.update();
    });
}

// テーブルの更新後にチャートも更新
function updateTableAndCharts() {
    // ...existing table update code...
    updateCharts();
}

// レスポンシブ対応の強化
window.addEventListener('resize', function() {
    updateCharts();
});

// エラー処理の改善
window.addEventListener('error', function(e) {
    console.error('Runtime error:', e);
    hideLoading();
    alert('エラーが発生しました。ページを更新してください。');
});

// 印刷用メディアクエリのスタイル動的追加
const printStyles = document.createElement('style');
printStyles.textContent = `
    @media print {
        body { background: white; }
        .container { box-shadow: none; }
        .controls, .csv-download { display: none; }
        .section { break-inside: avoid; }
        table { page-break-inside: avoid; }
        @page { margin: 2cm; }
    }
`;
document.head.appendChild(printStyles);