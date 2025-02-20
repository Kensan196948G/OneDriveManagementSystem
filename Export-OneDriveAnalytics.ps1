# OneDrive KFM analytics data exporter
[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ReportPath = "$PSScriptRoot\Reports",
    [Parameter(Mandatory = $false)]
    [int]$HistoryDays = 30
)

# データの収集と整形
function Get-OneDriveAnalytics {
    $analytics = @{
        UsageStats = @()
        TrendData = @()
        ComplianceData = @()
        ErrorStats = @()
    }

    # 使用状況データの収集
    Get-ChildItem "$ReportPath\OneDriveUsage_*.csv" |
        Sort-Object LastWriteTime |
        Select-Object -Last $HistoryDays |
        ForEach-Object {
            $dateStr = $_.BaseName -replace 'OneDriveUsage_'
            $date = [datetime]::ParseExact($dateStr, 'yyyyMMdd', $null)
            
            $usageData = Import-Csv $_
            $analytics.UsageStats += $usageData | ForEach-Object {
                [PSCustomObject]@{
                    Date = $date
                    UserName = $_.Name
                    UsedStorage = [double]($_.UsedStorageGB)
                    TotalStorage = [double]($_.AllocatedStorageGB)
                    UsagePercentage = [double]($_.UsagePercentage)
                }
            }
        }

    # トレンドデータの生成
    $analytics.TrendData = $analytics.UsageStats |
        Group-Object Date |
        ForEach-Object {
            [PSCustomObject]@{
                Date = $_.Name
                AverageUsage = ($_.Group | Measure-Object UsedStorage -Average).Average
                TotalUsage = ($_.Group | Measure-Object UsedStorage -Sum).Sum
                UserCount = $_.Count
            }
        }

    return $analytics
}

# 予測分析の実行
function Invoke-PredictiveAnalysis {
    param(
        $historicalData
    )

    $n = $historicalData.Count
    if ($n -lt 2) { return @() }

    # データを数値配列に変換
    $dates = $historicalData | ForEach-Object { $_.Date.ToOADate() }
    $usage = $historicalData | ForEach-Object { $_.TotalUsage }

    # 線形回帰の計算
    $sumX = ($dates | Measure-Object -Sum).Sum
    $sumY = ($usage | Measure-Object -Sum).Sum
    $sumXY = 0
    $sumX2 = 0

    for ($i = 0; $i -lt $n; $i++) {
        $sumXY += $dates[$i] * $usage[$i]
        $sumX2 += $dates[$i] * $dates[$i]
    }

    $slope = ($n * $sumXY - $sumX * $sumY) / ($n * $sumX2 - $sumX * $sumX)
    $intercept = ($sumY - $slope * $sumX) / $n

    # 将来90日分の予測を生成
    $predictions = @()
    $lastDate = [DateTime]($historicalData[-1].Date)

    1..90 | ForEach-Object {
        $futureDate = $lastDate.AddDays($_)
        $predictedValue = $slope * $futureDate.ToOADate() + $intercept
        $predictions += [PSCustomObject]@{
            Date = $futureDate
            PredictedUsage = [Math]::Max(0, $predictedValue)
            Confidence = 1 - ($_/90) # 信頼度は時間とともに低下
        }
    }

    return $predictions
}

# メイン処理
try {
    # データの収集と分析
    $analytics = Get-OneDriveAnalytics
    $predictions = Invoke-PredictiveAnalysis $analytics.TrendData

    # HTMLレポートの生成
    $htmlTemplate = @"
<!DOCTYPE html>
<html>
<head>
    <title>OneDrive Analytics Report</title>
    <meta charset="UTF-8">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <link rel="stylesheet" href="report-utils.css">
</head>
<body>
    <div class="report-container">
        <h1>OneDrive Analytics Report</h1>
        <div class="chart-container">
            <canvas id="usageChart"></canvas>
        </div>
        <div class="chart-container">
            <canvas id="predictionChart"></canvas>
        </div>
        <!-- データテーブル -->
    </div>
    <script src="report-utils.js"></script>
</body>
</html>
"@

    # レポートの出力
    $reportDate = Get-Date -Format "yyyyMMdd"
    $htmlTemplate | Out-File "$ReportPath\OneDriveAnalytics_$reportDate.html" -Encoding UTF8
    $analytics.UsageStats | Export-Csv "$ReportPath\OneDriveAnalytics_$reportDate.csv" -NoTypeInformation -Encoding UTF8
    $predictions | Export-Csv "$ReportPath\OneDrivePredictions_$reportDate.csv" -NoTypeInformation -Encoding UTF8

    Write-Host "Analytics report generated successfully."
}
catch {
    Write-Error "データ分析処理に失敗しました: $_"
    exit 1
}