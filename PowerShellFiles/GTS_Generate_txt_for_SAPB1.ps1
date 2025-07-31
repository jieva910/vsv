[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
$OpenFileDialog.ShowDialog() | Out-Null


$GTSHeader = "" | Select-Object -Property '金税编号',	'金税(数电)号码',	'行数',	'日期',	'月份',	'SAP发票号码',	'未税金额',	'税率',	'税额',	'客户名称',	'客户税号',	'客户的地址电话',	'客户银行账号',	'维苏威公司名称',	'维苏威公司税号',	'维苏威公司的地址电话',	'维苏威公司银行账号',	'备注'

$sourceFile = $OpenFileDialog.filename
$filename = get-date -Format "yyyyMMddhhmmss"
 
$outFile = "$($PSScriptRoot)\GTS$($filename).txt"
"SJJK0201~~已开发票传出" | Out-File -FilePath $outFile  -Encoding utf8

$col = 1
$usedCellType = 11
$SDNumberCounter=1
$excelApp = New-Object -ComObject Excel.Application 
 try {
        $excelApp.visible = $false;
        $excelApp.DisplayAlerts = $false 

        $workbook = $excelApp.Workbooks.Open($sourceFile) 
        $worksheetSummary = $workbook.WorkSheets.item("发票基础信息")
       $SDNumberCounter=  $excelApp.WorksheetFunction.CountIfs($worksheetSummary.Range("D:D"), "*00000*",$worksheetSummary.Range("O:O"),"正常",$worksheetSummary.Range("P:P"),"是")
       $maxdate = $worksheetSummary.Range("I2:I2000").value2|Sort-Object -Descending | select -First 1 
        $minDate = $worksheetSummary.Range("I2:I2000").value2|Sort-Object  | select -First 1 
         "$($SDNumberCounter)~~$($minDate.Substring(0,10) -replace('-',''))~~$($maxdate.Substring(0,10) -replace('-',''))" | Out-File -FilePath $outFile -Append -Encoding utf8
        $endRowSummary = $worksheetSummary.UsedRange.SpecialCells($usedCellType).Row
        $hangshu = 1 
        for($startRow=2;$startRow -le $endRowSummary;$startRow++)
        {
           [STRING]$SDNumber = $worksheetSummary.Columns.Item(4).Rows.Item($startRow).TEXT
            $getfromRows_shuilv=""
            
            
          if ($SDNumber.length -gt 10 -and ![string]::IsNullOrEmpty($SDNumber) -and $worksheetSummary.Columns.Item(15).Rows.Item($startRow).TEXT -eq "正常" -and $worksheetSummary.Columns.Item(16).Rows.Item($startRow).TEXT -eq "是" )
          {   
              $getfromRows_lines=1 
              $GTSHeader.金税编号="0~~0~~0~~$($SDNumber.SubString(0,10))"
              $GTSHeader.'金税(数电)号码'= $SDNumber.SubString($SDNumber.length - 8, 8)
              $GTSHeader.日期=$worksheetSummary.Columns.Item(9).Rows.Item($startRow).text.Substring(0,10) -replace('-','') 
              $fMDate = [datetime]$worksheetSummary.Columns.Item(9).Rows.Item($startRow).TEXT 
              $GTSHeader.月份=$fMDate.Month
              $sapinv = [STRING]$worksheetSummary.Columns.Item(19).Rows.Item($startRow).TEXT+'noSAPInvoce'
              $GTSHeader.SAP发票号码= $sapinv.SubString(0,10)
              $GTSHeader.未税金额=$worksheetSummary.Columns.Item(10).Rows.Item($startRow).text
              $GTSHeader.税额=$worksheetSummary.Columns.Item(11).Rows.Item($startRow).text
              $GTSHeader.客户名称=$worksheetSummary.Columns.Item(8).Rows.Item($startRow).text
              $GTSHeader.客户税号=$worksheetSummary.Columns.Item(7).Rows.Item($startRow).text
              $GTSHeader.客户的地址电话=""
              $GTSHeader.客户银行账号=""
              $GTSHeader.维苏威公司名称=$worksheetSummary.Columns.Item(6).Rows.Item($startRow).text
              $GTSHeader.维苏威公司税号=$worksheetSummary.Columns.Item(5).Rows.Item($startRow).text
              $GTSHeader.维苏威公司的地址电话=""
              $GTSHeader.维苏威公司银行账号=""
              $GTSHeader.备注=$worksheetSummary.Columns.Item(4).Rows.Item($startRow).text
              
              # find shu dian number in below sheet
              $WorkSheet= $workbook.WorkSheets.item("信息汇总表")
              $Found = $WorkSheet.Cells.Find($SDNumber) #What
              $GTSRows=  If ($Found) {
                      # Address Method https://msdn.microsoft.com/en-us/vba/excel-vba/articles/range-address-property-excel
                      $BeginAddress = $Found.Address(0,0,1,1)
                        #Initial Found Cell
                         [PSCustomObject]@{
                           prefix=0
                           品名=$Found.Offset(0,8).text
                           型号=$Found.Offset(0,9).text
                           单位=$Found.Offset(0,10).text
                           数量=$Found.Offset(0,11).text
                           含税金额=$Found.Offset(0,16).text
                           税率=[decimal]$Found.Offset(0,14).text.replace('%','')/100
                           未税金额=$Found.Offset(0,13).text
                           税额=$Found.Offset(0,15).text
                           suffix=0  
                          }
                     $getfromRows_shuilv=[decimal]$Found.Offset(0,14).text.replace('%','')/100
                   Do {
                    $Found = $WorkSheet.Cells.FindNext($Found)
                    $Address = $Found.Address(0,0,1,1)
                    If ($Address -eq $BeginAddress) {
                        BREAK
                    }
                  [PSCustomObject]@{
                       prefix=0
                       品名=$Found.Offset(0,8).text
                       型号=$Found.Offset(0,9).text
                       单位=$Found.Offset(0,10).text
                       数量=$Found.Offset(0,11).text
                       含税金额=$Found.Offset(0,16).text
                       税率=[decimal]$Found.Offset(0,14).text.replace('%','')/100
                       未税金额=$Found.Offset(0,13).text
                       税额=$Found.Offset(0,15).text
                       suffix=0  
                      }
                       $getfromRows_lines++  
                 
                } Until ($False)
             }
              "//发票$($hangshu)"| Out-File -FilePath $outFile -Append -Encoding utf8
              $GTSHeader.行数= $getfromRows_lines
              $GTSHeader.税率 =$getfromRows_shuilv
             $GTSHeader.PSObject.Properties.Value -join '~~' | Out-File -FilePath $outFile -Append -Encoding utf8
            foreach($c in $GTSRows)
            {$c.psobject.Properties.value -join "~~" | Out-File -FilePath $outFile -Append -Encoding utf8 }   #psobject.Properties.value -join "~~"  
              $hangshu++
            }
        } 
     
        $workbook.Close($false) 
    }
    finally {
        $excelApp.Quit()
    }
