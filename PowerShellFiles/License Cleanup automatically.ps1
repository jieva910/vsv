<#  2021.02.22
   Purpose : 自动清除SAPB1用户的Professional license 
   Method  : 1. 备份B1Upf.xml
             2. 读取B1Upf.xml 获得professional license user 列表，根据6个月不登录的规则 ,或者在AD里面查找5+3格式的USER id，如果没有AD账号则需要删除。
             3. 停止server tool service
             4. 从b1upf.xml里面删除符合以上第二点的用户
             5. 保存b1upf.xml
             6. 启动 server tool service ,并发送邮件给相关IT。
#>


[xml]$licenseUser = @"
<?xml version="1.0" encoding="UTF-16"?><Users>
  <User>
<UserName>AlertSvc</UserName>
<IsConnected>0</IsConnected>
    <Modules>
    </Modules>
  </User>
  <User>
<UserName>B1i</UserName>
<IsConnected>0</IsConnected>
    <Modules>
      <Module>
<KeyType>B1iINDIRECT</KeyType>
<KeyDesc>B1iINDIRECT_MSS</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
      <Module>
<KeyType>PROFESSIONAL</KeyType>
<KeyDesc>SAP Business One Professional User</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
      <Module>
<KeyType>SAP-ADDONS</KeyType>
<KeyDesc>SAP AddOns</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
      <Module>
<KeyType>SAP0000007051</KeyType>
<KeyDesc>B1i</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
    </Modules>
  </User>
  <User>
<UserName>CoultDar</UserName>
<IsConnected>0</IsConnected>
    <Modules>
      <Module>
<KeyType>PROFESSIONAL</KeyType>
<KeyDesc>SAP Business One Professional User</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
    </Modules>
  </User>
  <User>
<UserName>EDsUser</UserName>
<IsConnected>0</IsConnected>
    <Modules>
      <Module>
<KeyType>SAP-ADDONS</KeyType>
<KeyDesc>SAP AddOns</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
      <Module>
<KeyType>SAP0000007051</KeyType>
<KeyDesc>B1i</KeyDesc>
<DbType>MSS</DbType>
<BitmaskOfLicensedModules>-255</BitmaskOfLicensedModules>
<ReferingCount>0</ReferingCount>
<InstallNo>0020509653</InstallNo>
      </Module>
    </Modules>
  </User>
</Users>
"@

$username = 'EDsUser'
$licenseUser.SelectNodes("//UserName[.='$username']")| % {$_.ParentNode.ParentNode.removechild($_.ParentNode) } | Out-Null



$licenseUser.Users.User


