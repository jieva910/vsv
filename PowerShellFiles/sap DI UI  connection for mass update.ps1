param(
    [Parameter(Mandatory=$true)]
    [string]$Path  # 接收BAT传递的路径参数
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class DPI { 
    [DllImport("user32.dll")] 
    public static extern bool SetProcessDPIAware();
}
"@

# 启用高DPI感知（Windows 8.1之前）
if ([Environment]::OSVersion.Version -lt [Version]"6.3") { 
    [DPI]::SetProcessDPIAware() | Out-Null 
}

# 加载系统视觉样式
[System.Windows.Forms.Application]::EnableVisualStyles()

# 创建全局SAP对象
$script:diCompany = $null
$script:uiApplication = $null
$script:isRunning = $false
$script:diConnectionStatus = "未连接"
$script:uiConnectionStatus = "未连接"
# 定义功能函数
function Batch_Close_SO {
    param($Company, $Application)
    # 这里是批量关闭销售订单的功能实现
    return "批量关闭销售订单功能执行成功"
}

# UI连接函数
function UI_DI_Conn {
    try {
        # 创建UI API对象
        $SboGuiApi = New-Object -comobject "SAPbouiCOM.SboGuiApi"
        $sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
         $script:uiCompany = New-Object -COMobject "SAPbobsCOM.Company"
        # 连接到正在运行的SAP实例
        $SboGuiApi.Connect($sConnectionString)
        $script:uiApplication = $SboGuiApi.GetApplication()
        
        # 获取DI连接上下文
        $sCookie = $script:uiCompany.GetContextCookie()
        $sConnectionContext = $script:uiApplication.Company.GetConnectionContext($sCookie)
        
        # 设置登录上下文
        $ret = $script:uiCompany.SetSboLoginContext($sConnectionContext)
        if ($ret -ne 0) {
            throw "设置登录上下文失败 (错误码 $ret): $($script:uiCompany.GetLastErrorDescription())"
        }
        
        # 连接到公司数据库
        $retVal = $script:uiCompany.Connect()
        if ($retVal -ne 0) {
            throw "UI连接失败 (错误码 $retVal): $($script:uiCompany.GetLastErrorDescription())"
        }
        
        return 0
    }
    catch {
        # 清理资源
        if ($script:uiApplication) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:uiApplication) | Out-Null
            $script:uiApplication = $null
        }
        throw $_.Exception.Message
    }
}

# 创建主窗体
$form = New-Object System.Windows.Forms.Form
$form.Text = "SAP Business One 脚本执行器 (SSO支持)"
$form.Size = New-Object System.Drawing.Size(700, 950)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# 创建TabControl
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 10)
$tabControl.Size = New-Object System.Drawing.Size(670, 800)
$form.Controls.Add($tabControl)

# 创建UI标签页
$tabPageUI = New-Object System.Windows.Forms.TabPage
$tabPageUI.Text = "UI"
$tabPageUI.Name = "tabPageUI"
$tabPageUI.Padding = New-Object System.Windows.Forms.Padding(3)
$tabControl.Controls.Add($tabPageUI)

# 创建DI标签页
$tabPageDI = New-Object System.Windows.Forms.TabPage
$tabPageDI.Text = "DI"
$tabPageDI.Name = "tabPageDI"
$tabPageDI.Padding = New-Object System.Windows.Forms.Padding(3)
$tabControl.Controls.Add($tabPageDI)

# ================== 在DI标签页中添加DI连接面板 ==================
$diConnPanel = New-Object System.Windows.Forms.Panel
$diConnPanel.Location = New-Object System.Drawing.Point(10, 10)
$diConnPanel.Size = New-Object System.Drawing.Size(630, 180)
$diConnPanel.BorderStyle = "FixedSingle"
$tabPageDI.Controls.Add($diConnPanel)

# License服务器
$lblLicense = New-Object System.Windows.Forms.Label
$lblLicense.Location = New-Object System.Drawing.Point(20, 20)
$lblLicense.Size = New-Object System.Drawing.Size(100, 20)
$lblLicense.Text = "License服务器:"
$diConnPanel.Controls.Add($lblLicense)

$SLDSvr = New-Object System.Windows.Forms.ComboBox
$SLDSvr.Location = New-Object System.Drawing.Point(130, 20)
$SLDSvr.Size = New-Object System.Drawing.Size(150, 20)
$SLDSvr.Items.AddRange(@("SZ-SAPLIC92:40000", "SZ-TSTSAPLIC92:40000"))
$SLDSvr.SelectedIndex = 0
$diConnPanel.Controls.Add($SLDSvr)

# 数据库服务器
$lblDbServer = New-Object System.Windows.Forms.Label
$lblDbServer.Location = New-Object System.Drawing.Point(20, 50)
$lblDbServer.Size = New-Object System.Drawing.Size(100, 20)
$lblDbServer.Text = "数据库服务器:"
$diConnPanel.Controls.Add($lblDbServer)

$txtDbServer = New-Object System.Windows.Forms.ComboBox
$txtDbServer.Location = New-Object System.Drawing.Point(130, 50)
$txtDbServer.Size = New-Object System.Drawing.Size(150, 20)
$txtDbServer.Items.AddRange(@("SZ-SAP01", "sz-sapstg91", "sz-sapstg92"))
$txtDbServer.SelectedIndex = 0
$diConnPanel.Controls.Add($txtDbServer)

# 数据库名称
$lblDbName = New-Object System.Windows.Forms.Label
$lblDbName.Location = New-Object System.Drawing.Point(20, 80)
$lblDbName.Size = New-Object System.Drawing.Size(100, 20)
$lblDbName.Text = "公司数据库:"
$diConnPanel.Controls.Add($lblDbName)

$txtDbName = New-Object System.Windows.Forms.ComboBox
$txtDbName.Location = New-Object System.Drawing.Point(130, 80)
$txtDbName.Size = New-Object System.Drawing.Size(150, 20)
$txtDbName.Items.AddRange(@("SAPB1_AS", "SAPB1_BY", "SAPB1_CS", "SAPB1_KT", "SAPB1-SZ", "SAPB1_WN", "SAPB1_YK", "SAPB1_CS_TST", "SAPB1_BY_TST", "SAPB1_SZ_TST"))
$txtDbName.SelectedIndex = 0
$diConnPanel.Controls.Add($txtDbName)

# SAP用户名
$lblSapUser = New-Object System.Windows.Forms.Label
$lblSapUser.Location = New-Object System.Drawing.Point(300, 20)
$lblSapUser.Size = New-Object System.Drawing.Size(100, 20)
$lblSapUser.Text = "SAP用户名:"
$diConnPanel.Controls.Add($lblSapUser)

$txtSapUser = New-Object System.Windows.Forms.TextBox
$txtSapUser.Location = New-Object System.Drawing.Point(410, 20)
$txtSapUser.Size = New-Object System.Drawing.Size(150, 20)
$txtSapUser.Text = "manager"
$diConnPanel.Controls.Add($txtSapUser)

# SAP密码
$lblSapPassword = New-Object System.Windows.Forms.Label
$lblSapPassword.Location = New-Object System.Drawing.Point(300, 50)
$lblSapPassword.Size = New-Object System.Drawing.Size(100, 20)
$lblSapPassword.Text = "SAP密码:"
$diConnPanel.Controls.Add($lblSapPassword)

$txtSapPassword = New-Object System.Windows.Forms.TextBox
$txtSapPassword.Location = New-Object System.Drawing.Point(410, 50)
$txtSapPassword.Size = New-Object System.Drawing.Size(150, 20)
$txtSapPassword.PasswordChar = '*'
$txtSapPassword.Text = "manager"
$diConnPanel.Controls.Add($txtSapPassword)

# 数据库类型
$lblDbType = New-Object System.Windows.Forms.Label
$lblDbType.Location = New-Object System.Drawing.Point(300, 80)
$lblDbType.Size = New-Object System.Drawing.Size(100, 20)
$lblDbType.Text = "数据库类型:"
$diConnPanel.Controls.Add($lblDbType)

$cmbDbType = New-Object System.Windows.Forms.ComboBox
$cmbDbType.Location = New-Object System.Drawing.Point(410, 80)
$cmbDbType.Size = New-Object System.Drawing.Size(150, 20)
$cmbDbType.Items.AddRange(@("MSSQL2014", "MSSQL2016", "MSSQL2019"))
$cmbDbType.SelectedIndex = 0
$diConnPanel.Controls.Add($cmbDbType)

# SSO连接复选框
$chkUseSSO = New-Object System.Windows.Forms.CheckBox
$chkUseSSO.Location = New-Object System.Drawing.Point(300, 110)
$chkUseSSO.Size = New-Object System.Drawing.Size(150, 20)
$chkUseSSO.Text = "使用SSO连接"
$chkUseSSO.Checked = $true
$chkUseSSO.Add_CheckedChanged({
    $txtSapUser.Enabled = -not $chkUseSSO.Checked
    $txtSapPassword.Enabled = -not $chkUseSSO.Checked
})
$diConnPanel.Controls.Add($chkUseSSO)

# 初始禁用用户名密码输入框
$txtSapUser.Enabled = $false
$txtSapPassword.Enabled = $false

# 连接按钮（DI连接）
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Location = New-Object System.Drawing.Point(20, 140)
$btnConnect.Size = New-Object System.Drawing.Size(80, 30)
$btnConnect.Text = "DI连接"
$btnConnect.Add_Click({
    $licenseServer = $SLDSvr.Text
    $dbServer = $txtDbServer.Text
    $dbName = $txtDbName.Text
    $sapUser = $txtSapUser.Text
    $sapPassword = $txtSapPassword.Text
    $dbType = $cmbDbType.SelectedItem.ToString()
    $useSSO = $chkUseSSO.Checked
    
    if ([string]::IsNullOrWhiteSpace($licenseServer) -or 
        [string]::IsNullOrWhiteSpace($dbServer) -or 
        [string]::IsNullOrWhiteSpace($dbName)) {
        [System.Windows.Forms.MessageBox]::Show("请填写所有连接信息", "错误", "OK", "Error")
        return
    }
    
    try {
       
        # 如果DI连接已存在，先断开
        if ($script:uiConnectionStatus -eq "已连接" ) {
            throw "请先断开 UI连接"
        }
        
       
        $diConnectionStatusLabel.Text = "连接中..."
        $diConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()
        
        # 创建新公司对象
        $script:diCompany = New-Object -comobject SAPbobsCOM.Company
        
        # 设置连接属性
        $script:diCompany.Server = $dbServer
        $script:diCompany.CompanyDB = $dbName
        $script:diCompany.SLDServer = $licenseServer
        $script:diCompany.LicenseServer = $licenseServer
        $script:diCompany.UseTrusted = [int]$useSSO
        
        # 设置数据库类型
        switch ($dbType) {
            "MSSQL2014" { $script:diCompany.DbServerType = 8 }
            "MSSQL2016" { $script:diCompany.DbServerType = 10 }
            "MSSQL2019" { $script:diCompany.DbServerType = 12 }
        }
        
        # 使用SSO连接
        if ($useSSO) {
            $script:diCompany.UserName = "\"
            $script:diCompany.Password = ""
        }
        # 使用传统用户名/密码连接
        else {
            $script:diCompany.UserName = $sapUser
            $script:diCompany.Password = $sapPassword
        }
        
        # 尝试连接
        $retVal = $script:diCompany.Connect()
        
        if ($retVal -ne 0) {
            $errorMsg = $script:diCompany.GetLastErrorDescription()
            throw "SAP连接失败 (错误码 $retVal): $errorMsg"
        }
        
        # 更新UI状态
        $script:diConnectionStatus = "已连接"
        $diConnectionStatusLabel.Text = $script:diConnectionStatus
        $diConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Green
        
        $btnConnect.Enabled = $false
        $btnDisconnect.Enabled = $true
        
        # 启用DI标签页中的所有功能按钮
        $tabPageDI.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] } | ForEach-Object {
            $_.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] } | ForEach-Object {
                $_.Enabled = $true
            }
        }
        
        # 显示成功消息
        $connType = if ($useSSO) { "SSO" } else { "用户名/密码" }
        [System.Windows.Forms.MessageBox]::Show("成功通过$connType连接到SAP Business One!", "连接成功", "OK", "Information")
    }
    catch {
        if ($script:diCompany) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:diCompany) | Out-Null
            $script:diCompany = $null
        }
        $diConnectionStatusLabel.Text = "连接失败"
        $diConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("连接失败: $($_.Exception.Message)", "错误", "OK", "Error")
    }
})
$diConnPanel.Controls.Add($btnConnect)

# 断开按钮（DI连接）
$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Location = New-Object System.Drawing.Point(110, 140)
$btnDisconnect.Size = New-Object System.Drawing.Size(80, 30)
$btnDisconnect.Text = "断开DI"
$btnDisconnect.Enabled = $false
$btnDisconnect.Add_Click({
    try {
        if ($script:diCompany -ne $null) {
            # 断开连接
            if ($script:diCompany.Connected) {
                $script:diCompany.Disconnect()
            }
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:diCompany) | Out-Null
            $script:diCompany = $null
        }
        
        # 更新UI状态
        $script:diConnectionStatus = "已断开"
        $diConnectionStatusLabel.Text = $script:diConnectionStatus
        $diConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
        
        $btnConnect.Enabled = $true
        $btnDisconnect.Enabled = $false
        
        # 禁用DI标签页中的所有功能按钮
        $tabPageDI.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] } | ForEach-Object {
            $_.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] } | ForEach-Object {
                $_.Enabled = $false
            }
        }
        
        # 显示断开消息
        [System.Windows.Forms.MessageBox]::Show("已断开SAP DI连接", "断开成功", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("断开连接时出错: $($_.Exception.Message)", "错误", "OK", "Error")
    }
})
$diConnPanel.Controls.Add($btnDisconnect)

# DI连接状态标签
$diConnectionStatusLabel = New-Object System.Windows.Forms.Label
$diConnectionStatusLabel.Location = New-Object System.Drawing.Point(200, 145)
$diConnectionStatusLabel.Size = New-Object System.Drawing.Size(200, 20)
$diConnectionStatusLabel.Text = $script:diConnectionStatus
$diConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
$diConnPanel.Controls.Add($diConnectionStatusLabel)

# 创建功能区域函数
function Create-Section {
    param(
        [System.Windows.Forms.Control]$container,
        [string]$sectionName,
        [string]$displayName,
        [int]$yPos,
        [array]$buttons
    )
    
    # 创建分组框
    $groupBox = New-Object System.Windows.Forms.GroupBox
    $groupBox.Location = New-Object System.Drawing.Point(20, $yPos)
    $groupBox.Size = New-Object System.Drawing.Size(610, 100)
    $groupBox.Text = $displayName
    
    # 按钮位置计数器
    $xPos = 20
    $buttonWidth = 150
    $buttonHeight = 30
    $spacing = 20
    
    # 创建按钮和进度条
    $progressBars = @()
    
    foreach ($button in $buttons) {
        # 创建按钮
        $btn = New-Object System.Windows.Forms.Button
        $btn.Location = New-Object System.Drawing.Point($xPos, 30)
        $btn.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
        $btn.Text = $button.DisplayName
        $btn.Tag = $button.ScriptPath
        $btn.Enabled = $false
        
        # 创建进度条
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Location = New-Object System.Drawing.Point($xPos, 65)
        $progressBar.Size = New-Object System.Drawing.Size($buttonWidth, 20)
        $progressBar.Style = "Marquee"
        $progressBar.Visible = $false
        $progressBar.Tag = "$($button.Name)-progress"
        
        # 按钮点击事件
        $btn.Add_Click({
            if ($script:isRunning) {
                [System.Windows.Forms.MessageBox]::Show("另一个脚本正在执行中", "警告", "OK", "Warning")
                return
            }
              $Company = $null
    $paramName = $null
            $scriptPath = $this.Tag
            $buttonName = $this.Text
            
            if (-not (Test-Path $scriptPath)) {
                [System.Windows.Forms.MessageBox]::Show("脚本文件不存在: $scriptPath", "错误", "OK", "Error")
                return
            }
            
            # 根据标签页确定使用哪个连接
            $tabPage = $this.Parent.Parent.Parent  # 按钮 -> GroupBox -> TabPage
            $useUI = ($tabPage.Name -eq "tabPageUI")
            
             if ($script:diConnectionStatus -eq '已连接' ) {
                $Company = $script:diCompany
                $paramName = "-Company"
            }
               if ($script:uiConnectionStatus -eq '已连接') {
                $Company = $script:uiCompany
                $paramName = "-Company"
            }
			
			# 验证连接对象
    if ($Company -eq $null) {
        [System.Windows.Forms.MessageBox]::Show("连接对象未初始化!`n请检查连接状态", "错误", "OK", "Error")
        return
    }
            # 更新UI状态
            $script:isRunning = $true
            $this.Parent.Controls | Where-Object { $_.Tag -eq "$($button.Name)-progress" } | ForEach-Object {
                $_.Visible = $true
            }
            $statusLabel.Text = "$buttonName 执行中..."
            
            # 禁用所有按钮
            $form.Controls | Where-Object { $_ -is [System.Windows.Forms.TabControl] } | ForEach-Object {
                $_.Controls | Where-Object { $_ -is [System.Windows.Forms.TabPage] } | ForEach-Object {
                    $_.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] } | ForEach-Object {
                        $_.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] } | ForEach-Object {
                            $_.Enabled = $false
                        }
                    }
                }
            }
            
            try {
                # 执行脚本
                $scriptBlock = [scriptblock]::Create((Get-Content -Path $scriptPath -Raw))
                
                # 将SAP连接对象传递给脚本
                $result = & $scriptBlock $paramName $Company
                
                # 显示执行结果
                [System.Windows.Forms.MessageBox]::Show("$buttonName 执行完成！`n输出结果:`n$result", "完成", "OK", "Information")
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("$buttonName 执行失败！`n错误信息:`n$($_.Exception.Message)", "错误", "OK", "Error")
            }
            finally {
                # 更新UI状态
                $script:isRunning = $false
                $this.Parent.Controls | Where-Object { $_.Tag -like "*-progress" } | ForEach-Object {
                    $_.Visible = $false
                }
                $statusLabel.Text = "就绪"
                
                # 启用所有按钮
                $form.Controls | Where-Object { $_ -is [System.Windows.Forms.TabControl] } | ForEach-Object {
                    $_.Controls | Where-Object { $_ -is [System.Windows.Forms.TabPage] } | ForEach-Object {
                        $_.Controls | Where-Object { $_ -is [System.Windows.Forms.GroupBox] } | ForEach-Object {
                            $_.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] } | ForEach-Object {
                                $_.Enabled = $true
                            }
                        }
                    }
                }
            }
        })
        
        # 添加到分组框
        $groupBox.Controls.Add($btn)
        $groupBox.Controls.Add($progressBar)
        
        # 移动到下一个位置
        $xPos += $buttonWidth + $spacing
        $progressBars += $progressBar
    }
    
    $container.Controls.Add($groupBox)
    return $groupBox
}

# ================== 在UI标签页中添加UI连接面板 ==================
$uiConnPanel = New-Object System.Windows.Forms.GroupBox
$uiConnPanel.Location = New-Object System.Drawing.Point(10, 10)
$uiConnPanel.Size = New-Object System.Drawing.Size(640, 100)
$uiConnPanel.Text = "UI API 连接"
$tabPageUI.Controls.Add($uiConnPanel)

# UI连接状态标签
$uiConnectionStatusLabel = New-Object System.Windows.Forms.Label
$uiConnectionStatusLabel.Location = New-Object System.Drawing.Point(400, 40)
$uiConnectionStatusLabel.Size = New-Object System.Drawing.Size(100, 20)
$uiConnectionStatusLabel.Text = $script:uiConnectionStatus
$uiConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
$uiConnPanel.Controls.Add($uiConnectionStatusLabel)

# UI连接按钮
$btnUIConnect = New-Object System.Windows.Forms.Button
$btnUIConnect.Location = New-Object System.Drawing.Point(500, 25)
$btnUIConnect.Size = New-Object System.Drawing.Size(80, 30)
$btnUIConnect.Text = "UI连接"
$btnUIConnect.Add_Click({
    try {
        # 如果UI连接已存在，先断开
        if ($script:uiApplication -ne $null) {
            $btnUIDisconnect.PerformClick()
        }
        
        # 检查DI连接
        if ($script:diCompany.Connected) {
            throw "请先断开 DI连接"
        }

        $uiConnectionStatusLabel.Text = "连接中..."
        $uiConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Blue
        $form.Refresh()

        # 执行UI连接
        $result = UI_DI_Conn
        
        # 更新UI状态
        $script:uiConnectionStatus = "已连接"
        $uiConnectionStatusLabel.Text = $script:uiConnectionStatus
        $uiConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Green
        
        $btnUIConnect.Enabled = $false
        $btnUIDisconnect.Enabled = $true
        
        # 启用UI标签页中的所有功能按钮
        $tabPageUI.Controls | Where-Object { 
            $_ -is [System.Windows.Forms.GroupBox] -and $_.Text -ne "UI API 连接" 
        } | ForEach-Object {
            $_.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] } | ForEach-Object {
                $_.Enabled = $true
            }
        }
        
        [System.Windows.Forms.MessageBox]::Show("UI连接成功! $Path", "成功", "OK", "Information")
    }
    catch {
        $uiConnectionStatusLabel.Text = "连接失败"
        $uiConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
        [System.Windows.Forms.MessageBox]::Show("UI连接失败: $_", "错误", "OK", "Error")
    }
})
$uiConnPanel.Controls.Add($btnUIConnect)

# UI断开按钮
$btnUIDisconnect = New-Object System.Windows.Forms.Button
$btnUIDisconnect.Location = New-Object System.Drawing.Point(500, 60)
$btnUIDisconnect.Size = New-Object System.Drawing.Size(80, 30)
$btnUIDisconnect.Text = "断开UI"
$btnUIDisconnect.Enabled = $false
$btnUIDisconnect.Add_Click({
    try {
        if ($script:uiApplication -ne $null) {
            # 清理UI对象 (COM对象不支持直接断开)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:uiApplication) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:uiCompany) | Out-Null
            $script:uiApplication = $null
             $script:uiCompany = $null
        }
        
        # 更新UI状态
        $script:uiConnectionStatus = "已断开"
        $uiConnectionStatusLabel.Text = $script:uiConnectionStatus
        $uiConnectionStatusLabel.ForeColor = [System.Drawing.Color]::Red
        
        $btnUIConnect.Enabled = $true
        $btnUIDisconnect.Enabled = $false
        
        # 禁用UI标签页中的所有功能按钮
        $tabPageUI.Controls | Where-Object { 
            $_ -is [System.Windows.Forms.GroupBox] -and $_.Text -ne "UI API 连接" 
        } | ForEach-Object {
            $_.Controls | Where-Object { $_ -is [System.Windows.Forms.Button] } | ForEach-Object {
                $_.Enabled = $false
            }
        }
        
        [System.Windows.Forms.MessageBox]::Show("UI连接已断开", "断开成功", "OK", "Information")
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("断开UI连接时出错: $($_.Exception.Message)", "错误", "OK", "Error")
    }
})
$uiConnPanel.Controls.Add($btnUIDisconnect)

# 定义功能按钮
$apButtons = @(
    @{Name = "AP"; DisplayName = "应付账款"; ScriptPath = "$Path\Scripts\AP_Script.ps1"}
)

$arButtons = @(
    @{Name = "AR"; DisplayName = "应收账款"; ScriptPath = "$Path\Scripts\AR_Script.ps1"}
)

$financeButtons = @(
    @{Name = "Accrual"; DisplayName = "Accrual Tool"; ScriptPath = "$Path\Scripts\Accrual_Tool.ps1"},
    @{Name = "Allocation"; DisplayName = "Allocation"; ScriptPath = "$Path\Scripts\Allocation.ps1"}
)

$procurementButtons = @(
    @{Name = "MassClosePO"; DisplayName = "Mass Close PO"; ScriptPath = "$Path\Scripts\Mass_Close_PO.ps1"},
    @{Name = "PriceUpdate"; DisplayName = "Special Price Update"; ScriptPath = "$Path\Scripts\Special_Price_Update.ps1"}
)

$salesButtons = @(
    @{Name = "BatchCloseSO"; DisplayName = "批量关闭销售订单"; ScriptPath = "$Path\Scripts\Batch_Close_SO.ps1"}
)

# 在UI标签页创建功能区域
Create-Section -container $tabPageUI -sectionName "AP" -displayName "应付账款" -yPos 120 -buttons $apButtons
Create-Section -container $tabPageUI -sectionName "AR" -displayName "应收账款" -yPos 230 -buttons $arButtons
Create-Section -container $tabPageUI -sectionName "Finance" -displayName "财务管理" -yPos 340 -buttons $financeButtons
Create-Section -container $tabPageUI -sectionName "Sales" -displayName "销售单据管理" -yPos 450 -buttons $salesButtons
Create-Section -container $tabPageUI -sectionName "Procurement" -displayName "采购管理" -yPos 560 -buttons $procurementButtons

# 在DI标签页创建功能区域（位置下移DI连接面板）
Create-Section -container $tabPageDI -sectionName "AP" -displayName "应付账款" -yPos 200 -buttons $apButtons
Create-Section -container $tabPageDI -sectionName "AR" -displayName "应收账款" -yPos 310 -buttons $arButtons
Create-Section -container $tabPageDI -sectionName "Finance" -displayName "财务管理" -yPos 420 -buttons $financeButtons
Create-Section -container $tabPageDI -sectionName "Sales" -displayName "销售单据管理" -yPos 530 -buttons $salesButtons
Create-Section -container $tabPageDI -sectionName "Procurement" -displayName "采购管理" -yPos 640 -buttons $procurementButtons

# 创建状态标签
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(20, 820)
$statusLabel.Size = New-Object System.Drawing.Size(400, 20)
$statusLabel.Text = "就绪"
$form.Controls.Add($statusLabel)

# 创建退出按钮
$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Location = New-Object System.Drawing.Point(550, 850)
$btnExit.Size = New-Object System.Drawing.Size(100, 30)
$btnExit.Text = "退出"
$btnExit.Add_Click({ 
    # 断开UI连接
    if ($script:uiApplication -ne $null) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:uiApplication) | Out-Null
    }
    
    # 断开DI连接
    if ($script:diCompany -ne $null) {
        if ($script:diCompany.Connected) {
            $script:diCompany.Disconnect()
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:diCompany) | Out-Null
    }
    
    $form.Close() 
})
$form.Controls.Add($btnExit)

# 显示窗体
[void]$form.ShowDialog()