Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# 应用Windows 10视觉样式
[System.Windows.Forms.Application]::EnableVisualStyles()
#[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# 连接到SAP Business One的UI API
Function SetApplication {
    try {
        $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
        $sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
        $SboGuiApi.Connect($sConnectionString)
        $app = $SboGuiApi.GetApplication()
        return $app
    }
    catch {
        throw "无法连接到SAP Business One: $($_.Exception.Message)"
    }
}

 # 找到ImportExcel模块路径
$modulePath =  "$PSScriptRoot\Modules\ImportExcel\7.8.10"

# 手动加载EPPlus程序集
Add-Type -Path "$modulePath\EPPlus.dll"
# 导入ImportExcel模块

if (-not (Get-Module -Name ImportExcel)) {
    Import-Module $modulePath\ImportExcel.psm1 -ErrorAction Stop
}

# 检查并安装必要的模块
if (-not (Get-Module  -Name ImportExcel)) {
    [System.Windows.Forms.MessageBox]::Show("ImportExcel模块未安装，请联系IT部门安装此模块。", "模块缺失", "OK", "Error")
    exit
}

# 获取脚本所在目录
$scriptPath = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }

# 主窗体设置 (Windows 10风格)
$form = New-Object System.Windows.Forms.Form
$form.Text = "SAP Business One SQL导出工具"
$form.Size = New-Object System.Drawing.Size(700, 550)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)

# 标题栏
$titleBar = New-Object System.Windows.Forms.Panel
$titleBar.Location = New-Object System.Drawing.Point(0, 0)
$titleBar.Size = New-Object System.Drawing.Size(700, 40)
$titleBar.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$form.Controls.Add($titleBar)

# 标题文本
$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Location = New-Object System.Drawing.Point(20, 10)
$titleLabel.Size = New-Object System.Drawing.Size(400, 25)
$titleLabel.Text = "SAP Business One SQL导出工具"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$titleBar.Controls.Add($titleLabel)

# SQL输入区域容器
$sqlContainer = New-Object System.Windows.Forms.GroupBox
$sqlContainer.Location = New-Object System.Drawing.Point(20, 60)
$sqlContainer.Size = New-Object System.Drawing.Size(650, 250)
$sqlContainer.Text = "SQL查询"
$sqlContainer.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$sqlContainer.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$form.Controls.Add($sqlContainer)

# SQL文本框（多行支持）
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(15, 25)
$textBox.Size = New-Object System.Drawing.Size(620, 200)
$textBox.Multiline = $true
$textBox.ScrollBars = "Both"
$textBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$sqlContainer.Controls.Add($textBox)

# 示例SQL（SBO特定）
$textBox.Text = @"
-- 示例SBO数据库查询
SELECT TOP 100 
    T0.[CardCode] AS '客户代码',
    T0.[CardName] AS '客户名称',
    T1.[DocNum] AS '单据编号',
    T1.[DocDate] AS '单据日期',
    T1.[DocTotal] AS '单据金额'
FROM 
    OCRD T0
    INNER JOIN OINV T1 ON T0.CardCode = T1.CardCode
WHERE
    T0.[CardType] = 'C' 
    AND T1.[DocStatus] = 'O'
ORDER BY
    T1.[DocDate] DESC
"@

# 数据库信息区域
$dbContainer = New-Object System.Windows.Forms.GroupBox
$dbContainer.Location = New-Object System.Drawing.Point(20, 320)
$dbContainer.Size = New-Object System.Drawing.Size(650, 80)
$dbContainer.Text = "数据库信息"
$dbContainer.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$dbContainer.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$form.Controls.Add($dbContainer)

# 数据库信息标签
$dbInfoLabel = New-Object System.Windows.Forms.Label
$dbInfoLabel.Location = New-Object System.Drawing.Point(15, 25)
$dbInfoLabel.Size = New-Object System.Drawing.Size(620, 40)
$dbInfoLabel.Text = "将使用当前SAP Business One连接的公司数据库..."
$dbInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$dbContainer.Controls.Add($dbInfoLabel)

# 选项区域
$optionsContainer = New-Object System.Windows.Forms.GroupBox
$optionsContainer.Location = New-Object System.Drawing.Point(20, 410)
$optionsContainer.Size = New-Object System.Drawing.Size(650, 60)
$optionsContainer.Text = "导出选项"
$optionsContainer.Font = New-Object System.Drawing.Font("Segoe UI", 9.5, [System.Drawing.FontStyle]::Bold)
$optionsContainer.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$form.Controls.Add($optionsContainer)

# 复选框
$checkBox = New-Object System.Windows.Forms.CheckBox
$checkBox.Location = New-Object System.Drawing.Point(15, 25)
$checkBox.Size = New-Object System.Drawing.Size(200, 30)
$checkBox.Text = "导出后自动打开Excel"
$checkBox.Checked = $true
$checkBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$optionsContainer.Controls.Add($checkBox)

# 导出按钮 (现代风格)
$button = New-Object System.Windows.Forms.Button
$button.Location = New-Object System.Drawing.Point(20, 480)
$button.Size = New-Object System.Drawing.Size(150, 40)
$button.Text = "导出数据"
$button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$button.ForeColor = [System.Drawing.Color]::White
$button.FlatStyle = "Flat"
$button.FlatAppearance.BorderSize = 0
$button.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$button.Cursor = [System.Windows.Forms.Cursors]::Hand
$form.Controls.Add($button)

# 状态标签
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(190, 480)
$statusLabel.Size = New-Object System.Drawing.Size(480, 40)
$statusLabel.Text = "就绪 - 输入SQL并点击导出按钮"
$statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
$statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(70, 70, 70)
$form.Controls.Add($statusLabel)

# 按钮悬停效果
$button.Add_MouseEnter({
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 90, 180)
})

$button.Add_MouseLeave({
    $button.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
})

# 按钮点击事件
$button.Add_Click({
    try {
        # 连接到SAP Business One
        $statusLabel.Text = "正在连接SAP Business One..."
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        $SBO_Application = SetApplication
        $company = $SBO_Application.Company
        
        # 显示数据库信息
        $dbInfo = "服务器: $($company.ServerName)`n数据库: $($company.DatabaseName)"
        $dbInfoLabel.Text = $dbInfo
        
        # 验证SQL输入
        if ([string]::IsNullOrWhiteSpace($textBox.Text)) {
            throw "SQL查询不能为空"
        }
        
        $statusLabel.Text = "正在执行SQL查询..."
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        # 构建连接字符串
        $connectionString = "Server=$($company.ServerName);Database=$($company.DatabaseName);Integrated Security=True;"
        
        # 执行SQL查询
        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()
        
        $command = $connection.CreateCommand()
        $command.CommandText = $textBox.Text
        
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataset) | Out-Null
        
        $connection.Close()
        
        if ($dataset.Tables[0] -eq $null -or $dataset.Tables[0].Rows.Count -eq 0) {
            $statusLabel.Text = "警告: 查询返回0行数据!"
            $statusLabel.ForeColor = [System.Drawing.Color]::DarkOrange
            return
        }
        
        # 生成带时间戳的文件名
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $excelPath = Join-Path $scriptPath "SBO_Export_$timestamp.xlsx"
        
        # 导出到Excel
        $statusLabel.Text = "正在导出数据到Excel..."
        $form.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        $dataset.Tables[0] | Export-Excel -Path $excelPath -WorksheetName "SBO数据" -AutoSize -FreezeTopRow -BoldTopRow -TableName "SBO_Data" -TableStyle "Medium6"
        
        $statusLabel.Text = "导出成功! 共导出 $($dataset.Tables[0].Rows.Count) 行数据`n文件位置: $excelPath"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 0)
        
        # 如果勾选则打开Excel文件
        if ($checkBox.Checked) {
            Start-Process $excelPath
        }
    }
    catch {
        $statusLabel.Text = "错误: $($_.Exception.Message)"
        $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(200, 0, 0)
        
        # 显示详细错误位置
        if ($_.InvocationInfo.PositionMessage) {
            $statusLabel.Text += "`n位置: $($_.InvocationInfo.PositionMessage)"
        }
        
        # 播放错误音效
        [System.Media.SystemSounds]::Hand.Play()
    }
})

# 显示窗体
[void]$form.ShowDialog()