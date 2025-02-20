# Main WPF application for OneDrive KFM Management Tool
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName LiveCharts.Wpf
Add-Type -AssemblyName LiveCharts.Core

# Load required modules for Power BI integration
Import-Module MicrosoftPowerBIMgmt

# XAML definition
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:lvc="clr-namespace:LiveCharts.Wpf;assembly=LiveCharts.Wpf"
    Title="OneDrive KFM Management Dashboard" Height="800" Width="1200"
    WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <!-- Top Menu -->
        <Menu Grid.Row="0">
            <MenuItem Header="ファイル">
                <MenuItem x:Name="MenuExportPDF" Header="PDFエクスポート"/>
                <MenuItem x:Name="MenuExportExcel" Header="Excelエクスポート"/>
                <Separator/>
                <MenuItem x:Name="MenuExit" Header="終了"/>
            </MenuItem>
            <MenuItem Header="設定">
                <MenuItem x:Name="MenuConfig" Header="環境設定"/>
                <MenuItem x:Name="MenuNotification" Header="通知設定"/>
                <MenuItem x:Name="MenuSchedule" Header="スケジュール設定"/>
            </MenuItem>
            <MenuItem Header="ツール">
                <MenuItem x:Name="MenuDeploy" Header="KFM展開"/>
                <MenuItem x:Name="MenuTroubleshoot" Header="トラブルシューティング"/>
            </MenuItem>
            <MenuItem Header="ヘルプ">
                <MenuItem x:Name="MenuAbout" Header="バージョン情報"/>
            </MenuItem>
        </Menu>

        <!-- Main Content -->
        <TabControl Grid.Row="1" Margin="5">
            <!-- Dashboard Tab -->
            <TabItem Header="ダッシュボード">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="300"/>
                    </Grid.ColumnDefinitions>

                    <!-- Main Dashboard Area -->
                    <Grid Grid.Column="0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Status Summary -->
                        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="5">
                            <Border Background="#FF4CAF50" CornerRadius="5" Margin="5" Padding="10">
                                <StackPanel>
                                    <TextBlock Text="正常" Foreground="White"/>
                                    <TextBlock x:Name="SuccessCount" Text="0" Foreground="White" FontSize="20"/>
                                </StackPanel>
                            </Border>
                            <Border Background="#FFFFC107" CornerRadius="5" Margin="5" Padding="10">
                                <StackPanel>
                                    <TextBlock Text="警告" Foreground="White"/>
                                    <TextBlock x:Name="WarningCount" Text="0" Foreground="White" FontSize="20"/>
                                </StackPanel>
                            </Border>
                            <Border Background="#FFF44336" CornerRadius="5" Margin="5" Padding="10">
                                <StackPanel>
                                    <TextBlock Text="エラー" Foreground="White"/>
                                    <TextBlock x:Name="ErrorCount" Text="0" Foreground="White" FontSize="20"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>

                        <!-- Storage Usage Chart -->
                        <Border Grid.Row="1" Margin="5" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="5">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Text="ストレージ使用状況" Margin="10,5" FontWeight="Bold"/>
                                <lvc:CartesianChart Grid.Row="1" x:Name="StorageChart"/>
                            </Grid>
                        </Border>

                        <!-- Trend Analysis Chart -->
                        <Border Grid.Row="2" Margin="5" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="5">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Text="使用量予測" Margin="10,5" FontWeight="Bold"/>
                                <lvc:CartesianChart Grid.Row="1" x:Name="TrendChart"/>
                            </Grid>
                        </Border>
                    </Grid>

                    <!-- Side Panel -->
                    <Grid Grid.Column="1">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Quick Actions -->
                        <StackPanel Grid.Row="0" Margin="5">
                            <Button x:Name="RefreshButton" Content="データ更新" Margin="2"/>
                            <Button x:Name="ExportButton" Content="レポート出力" Margin="2"/>
                            <Button x:Name="AlertButton" Content="アラート設定" Margin="2"/>
                        </StackPanel>

                        <!-- Recent Events -->
                        <Border Grid.Row="1" Margin="5" BorderBrush="#DDDDDD" BorderThickness="1" CornerRadius="5">
                            <Grid>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                </Grid.RowDefinitions>
                                <TextBlock Text="最近のイベント" Margin="10,5" FontWeight="Bold"/>
                                <ListBox Grid.Row="1" x:Name="EventsList" Margin="5"/>
                            </Grid>
                        </Border>
                    </Grid>
                </Grid>
            </TabItem>

            <!-- Monitoring Tab -->
            <TabItem Header="監視">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <!-- Filter Controls -->
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="5">
                        <TextBox x:Name="SearchBox" Width="200" Margin="2" ToolTip="検索"/>
                        <ComboBox x:Name="StatusFilter" Width="100" Margin="2"/>
                        <Button x:Name="ApplyFilter" Content="適用" Margin="2"/>
                    </StackPanel>

                    <!-- Data Grid -->
                    <DataGrid Grid.Row="1" x:Name="MonitoringGrid" AutoGenerateColumns="False" Margin="5">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="ユーザー名" Binding="{Binding Name}"/>
                            <DataGridTextColumn Header="メールアドレス" Binding="{Binding Email}"/>
                            <DataGridTextColumn Header="使用容量" Binding="{Binding UsedStorage}"/>
                            <DataGridTextColumn Header="使用率" Binding="{Binding UsagePercentage}"/>
                            <DataGridTextColumn Header="状態" Binding="{Binding Status}"/>
                            <DataGridTextColumn Header="最終更新" Binding="{Binding LastUpdate}"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>

            <!-- Settings Tab -->
            <TabItem Header="設定">
                <ScrollViewer>
                    <StackPanel Margin="10">
                        <!-- General Settings -->
                        <GroupBox Header="一般設定" Margin="0,5">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="テナントID:" Grid.Row="0" Grid.Column="0" Margin="5"/>
                                <TextBox x:Name="TenantIdBox" Grid.Row="0" Grid.Column="1" Margin="5"/>

                                <TextBlock Text="更新間隔:" Grid.Row="1" Grid.Column="0" Margin="5"/>
                                <ComboBox x:Name="RefreshInterval" Grid.Row="1" Grid.Column="1" Margin="5"/>

                                <TextBlock Text="ログ保持期間:" Grid.Row="2" Grid.Column="0" Margin="5"/>
                                <ComboBox x:Name="LogRetention" Grid.Row="2" Grid.Column="1" Margin="5"/>
                            </Grid>
                        </GroupBox>

                        <!-- Notification Settings -->
                        <GroupBox Header="通知設定" Margin="0,5">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="SMTPサーバー:" Grid.Row="0" Grid.Column="0" Margin="5"/>
                                <TextBox x:Name="SmtpServer" Grid.Row="0" Grid.Column="1" Margin="5"/>

                                <TextBlock Text="Teams Webhook:" Grid.Row="1" Grid.Column="0" Margin="5"/>
                                <TextBox x:Name="TeamsWebhook" Grid.Row="1" Grid.Column="1" Margin="5"/>

                                <TextBlock Text="通知先メール:" Grid.Row="2" Grid.Column="0" Margin="5"/>
                                <TextBox x:Name="NotificationEmail" Grid.Row="2" Grid.Column="1" Margin="5"/>
                            </Grid>
                        </GroupBox>

                        <!-- Alert Thresholds -->
                        <GroupBox Header="アラート閾値" Margin="0,5">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="150"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>

                                <TextBlock Text="使用率警告:" Grid.Row="0" Grid.Column="0" Margin="5"/>
                                <Slider x:Name="StorageWarning" Grid.Row="0" Grid.Column="1" Margin="5" 
                                        Minimum="0" Maximum="100" TickFrequency="5" IsSnapToTickEnabled="True"/>

                                <TextBlock Text="エラー数閾値:" Grid.Row="1" Grid.Column="0" Margin="5"/>
                                <Slider x:Name="ErrorThreshold" Grid.Row="1" Grid.Column="1" Margin="5"
                                        Minimum="0" Maximum="50" TickFrequency="5" IsSnapToTickEnabled="True"/>

                                <TextBlock Text="警告数閾値:" Grid.Row="2" Grid.Column="0" Margin="5"/>
                                <Slider x:Name="WarningThreshold" Grid.Row="2" Grid.Column="1" Margin="5"
                                        Minimum="0" Maximum="50" TickFrequency="5" IsSnapToTickEnabled="True"/>
                            </Grid>
                        </GroupBox>

                        <!-- Action Buttons -->
                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10">
                            <Button x:Name="SaveSettings" Content="設定を保存" Margin="5" Padding="10,5"/>
                            <Button x:Name="CancelSettings" Content="キャンセル" Margin="5" Padding="10,5"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@

# Create window
$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get controls
$controls = @{}
$xaml.SelectNodes("//*[@x:Name]") | ForEach-Object {
    $controls[$_.Name] = $window.FindName($_.Name)
}

# Initialize data
$script:monitoringData = @()
$script:chartData = @()

# Event handlers
$controls.RefreshButton.Add_Click({
    Update-DashboardData
})

$controls.SaveSettings.Add_Click({
    Save-Settings
})

# Functions
function Update-DashboardData {
    # Get monitoring data
    $monitoringData = & "$PSScriptRoot\Monitor-OneDriveStatus.ps1" -OutputPath "$PSScriptRoot\Reports" -LogPath "$PSScriptRoot\Logs"
    
    # Update status counts
    $controls.SuccessCount.Text = ($monitoringData | Where-Object Status -eq "Success").Count
    $controls.WarningCount.Text = ($monitoringData | Where-Object Status -eq "Warning").Count
    $controls.ErrorCount.Text = ($monitoringData | Where-Object Status -eq "Error").Count
    
    # Update storage chart
    Update-StorageChart $monitoringData
    
    # Update trend chart
    Update-TrendChart $monitoringData
    
    # Update monitoring grid
    $controls.MonitoringGrid.ItemsSource = $monitoringData
    
    # Update events list
    Update-EventsList
}

function Update-StorageChart {
    param($data)
    
    $storageData = $data | Select-Object Name, UsedStorage, TotalStorage
    
    $usedValues = New-Object LiveCharts.ChartValues[double]
    $remainingValues = New-Object LiveCharts.ChartValues[double]
    
    $storageData | ForEach-Object {
        $usedValues.Add($_.UsedStorage)
        $remainingValues.Add($_.TotalStorage - $_.UsedStorage)
    }
    
    $usedSeries = New-Object LiveCharts.Wpf.ColumnSeries
    $usedSeries.Title = "使用容量"
    $usedSeries.Values = $usedValues
    
    $remainingSeries = New-Object LiveCharts.Wpf.ColumnSeries
    $remainingSeries.Title = "残り容量"
    $remainingSeries.Values = $remainingValues
    
    $controls.StorageChart.Series.Clear()
    $controls.StorageChart.Series.Add($usedSeries)
    $controls.StorageChart.Series.Add($remainingSeries)
    
    $axisX = New-Object LiveCharts.Wpf.Axis
    $axisX.Labels = [string[]]($storageData.Name)
    $controls.StorageChart.AxisX = $axisX
}

function Update-TrendChart {
    param($data)
    
    $historicalData = Get-HistoricalData
    $prediction = Invoke-TrendAnalysis $historicalData
    
    $actualValues = New-Object LiveCharts.ChartValues[double]
    $predictedValues = New-Object LiveCharts.ChartValues[double]
    
    $historicalData | ForEach-Object {
        $actualValues.Add($_.UsedStorage)
    }
    
    $prediction | ForEach-Object {
        $predictedValues.Add($_.PredictedUsage)
    }
    
    $actualSeries = New-Object LiveCharts.Wpf.LineSeries
    $actualSeries.Title = "実際の使用量"
    $actualSeries.Values = $actualValues
    
    $predictedSeries = New-Object LiveCharts.Wpf.LineSeries
    $predictedSeries.Title = "予測使用量"
    $predictedSeries.Values = $predictedValues
    $predictedSeries.StrokeDashArray = New-Object System.Windows.Media.DoubleCollection @(2)
    
    $controls.TrendChart.Series.Clear()
    $controls.TrendChart.Series.Add($actualSeries)
    $controls.TrendChart.Series.Add($predictedSeries)
}

function Get-HistoricalData {
    # 過去30日分のデータを取得
    $historicalData = @()
    $startDate = (Get-Date).AddDays(-30)
    
    Get-ChildItem "$PSScriptRoot\Reports\OneDriveUsage_*.csv" | 
        Where-Object { $_.LastWriteTime -ge $startDate } |
        ForEach-Object {
            $historicalData += Import-Csv $_
        }
    
    return $historicalData
}

function Invoke-TrendAnalysis {
    param($historicalData)
    
    # 線形回帰による予測
    $dates = $historicalData | ForEach-Object { ($_.Date | Get-Date).ToOADate() }
    $usage = $historicalData.UsedStorage
    
    $n = $dates.Count
    $sumX = $dates | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumY = $usage | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumXY = 0
    $sumX2 = 0
    
    for ($i = 0; $i -lt $n; $i++) {
        $sumXY += $dates[$i] * $usage[$i]
        $sumX2 += $dates[$i] * $dates[$i]
    }
    
    $slope = ($n * $sumXY - $sumX * $sumY) / ($n * $sumX2 - $sumX * $sumX)
    $intercept = ($sumY - $slope * $sumX) / $n
    
    # 今後30日分の予測を生成
    $prediction = @()
    $today = (Get-Date).ToOADate()
    
    1..30 | ForEach-Object {
        $predictedDate = $today + $_
        $predictedUsage = $slope * $predictedDate + $intercept
        $prediction += [PSCustomObject]@{
            Date = [DateTime]::FromOADate($predictedDate)
            PredictedUsage = [Math]::Max(0, $predictedUsage)
        }
    }
    
    return $prediction
}

function Update-EventsList {
    # 最近のイベントを取得
    $events = Get-WinEvent -FilterHashtable @{
        LogName = 'Microsoft-Windows-OneDrive'
        Level = @(1,2,3)
        StartTime = (Get-Date).AddHours(-24)
    } -MaxEvents 50 -ErrorAction SilentlyContinue
    
    $controls.EventsList.Items.Clear()
    $events | ForEach-Object {
        $controls.EventsList.Items.Add("$($_.TimeCreated.ToString('HH:mm:ss')) - $($_.Message)")
    }
}

function Save-Settings {
    $settings = @{
        TenantId = $controls.TenantIdBox.Text
        RefreshInterval = $controls.RefreshInterval.SelectedValue
        LogRetention = $controls.LogRetention.SelectedValue
        SmtpServer = $controls.SmtpServer.Text
        TeamsWebhook = $controls.TeamsWebhook.Text
        NotificationEmail = $controls.NotificationEmail.Text
        StorageWarning = $controls.StorageWarning.Value
        ErrorThreshold = $controls.ErrorThreshold.Value
        WarningThreshold = $controls.WarningThreshold.Value
    }
    
    $settings | ConvertTo-Json | Set-Content "$PSScriptRoot\config\gui-settings.json"
    [Windows.MessageBox]::Show("設定を保存しました。", "設定の保存")
}

# Initialize
Update-DashboardData

# Show window
$window.ShowDialog()