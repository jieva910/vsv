# 设置参数
$sqlServer = "SZ-SAP01"
$database = "sapb1_CS"
$query = "SELECT t.CardCode, t.DocDate, t.CardName, t.DocTotal, t.DocCur FROM OPOR t WHERE t.DocDate > '20250601'"
$outputFile = "RESULT2.XLSX"
 Get-Date

  # 找到ImportExcel模块路径
$modulePath =  "$PSScriptRoot\Modules\ImportExcel\7.8.10"

# 手动加载EPPlus程序集
Add-Type -Path "$modulePath\EPPlus.dll"
# 导入ImportExcel模块

if (-not (Get-Module -Name ImportExcel)) {
    Import-Module $modulePath\ImportExcel.psm1 -ErrorAction Stop
}

# 创建连接字符串（集成身份验证）
$connectionString = "Server=$sqlServer;Database=$database;Integrated Security=True"

# 创建SQL连接对象
$connection = New-Object System.Data.SqlClient.SqlConnection
$connection.ConnectionString = $connectionString

try {
    # 打开数据库连接
    $connection.Open()
    
    # 执行SQL查询
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataset) | Out-Null
    
    # 获取结果表
    $results = $dataset.Tables[0]
    
    # 导出到Excel
    $results|Select-Object CardCode, DocDate, CardName, DocTotal, DocCur  | Export-Excel -Path $PSScriptRoot\$outputFile -WorksheetName "Results" -AutoSize -FreezeTopRow -BoldTopRow -ErrorAction Stop
    
    Write-Host "成功导出数据到 $outputFile" -ForegroundColor Green
}
catch {
    Write-Host "发生错误: $_" -ForegroundColor Red
}
finally {
    # 确保关闭连接
    if ($connection.State -eq 'Open') {
        $connection.Close()
    }
}
 Get-Date