
<#
  Purpose : Add /update language translation 
  Date    : 2020.11
#>


cls

$cmp           = New-Object -ComObject 'SAPbobsCOM.Company'
$SourceSite    = "sztst"



# load sapb1 di connection lib
. C:\shared\PShell\PSLib\SAPB1_DI_Site_Connection.ps1

# SAPB1 DI connect to specific site
Fn_ConnectSAPB1 $cmp $SourceSite

$langTranslation = $cmp.GetBusinessObject(224)    #SAPbobsCOM.BoObjectTypes.oMultiLanguageTranslations

$langTranslation.PrimaryKeyofobject = "v3"

$langTranslation.FieldAlias = "U_Text"  #(userfieds)

$langTranslation.TableName = "@SWA_LD_TEXT" #(usertable)

# 1) Add language translation for new 

$langTranslation.TranslationsInUserLanguages.Add()

$langTranslation.TranslationsInUserLanguages.LanguageCodeOfUserLanguage = 15

$langTranslation.TranslationsInUserLanguages.Translationscontent = "中文这个是"

$langTranslation.TranslationsInUserLanguages.Add()

$langTranslation.TranslationsInUserLanguages.LanguageCodeOfUserLanguage = 8

$langTranslation.TranslationsInUserLanguages.Translationscontent = "this is for uk english"

 $langTranslation.add()
 $cmp.GetLastErrorDescription()


# 2) Update language translation for existing


if ($langTranslation.GetByKey(1842)) {
 $langTranslation.Numerator;$langTranslation.PrimaryKeyofobject
}
 
 if ($langTranslation.GetByKey(1842))  # this key is TranEntry of table OMLT
 {
    $langTranslation.TranslationsInUserLanguages.SetCurrentLine(0)
    $langTranslation.TranslationsInUserLanguages.LanguageCodeOfUserLanguage = 31 
  $langTranslation.TranslationsInUserLanguages.Translationscontent = "this is Turkey after 1st add translation"

  $langTranslation.Update()
 $cmp.GetLastErrorDescription()
 }

# 3) Add language translation for existing 

 if ($langTranslation.GetByKey(1842))
 {
    $langTranslation.TranslationsInUserLanguages.ADD()
    $langTranslation.TranslationsInUserLanguages.LanguageCodeOfUserLanguage = 31 
  $langTranslation.TranslationsInUserLanguages.Translationscontent = "this is Turkey after 1st add translation"

  $langTranslation.Update()
 $cmp.GetLastErrorDescription()
 }


 # add translation for authorization group name 2024.02.02

 $trans_csv = Import-Csv 'C:\Temp\Role Translation.csv'

 foreach($r in $trans_csv)
 {  
    $langTranslation = $cmp.GetBusinessObject(224)    #SAPbobsCOM.BoObjectTypes.oMultiLanguageTranslations
  $langTranslation.FieldAlias = "GroupName"  #(userfieds)
 $langTranslation.TableName = "OUGR" #(usertable)
   $langTranslation.PrimaryKeyofobject =  $r.groupid   # id or docentry or code
   $langTranslation.TranslationsInUserLanguages.Add()
   $langTranslation.TranslationsInUserLanguages.LanguageCodeOfUserLanguage = 15
    $langTranslation.TranslationsInUserLanguages.Translationscontent = $r.translation
     $langTranslation.add()
 $cmp.GetLastErrorDescription()
  }
 
