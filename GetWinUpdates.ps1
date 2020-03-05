#Отправляет по почте список серверов и их последнего обновления, если оно было более 3 месяцев назад
#
# ДАТА: 05 марта 2020 года										   
 
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

#Для успешного выполнения скрипта необходимо:
<# 

Powersheel версии 4
Powersheel ExecutionPolicy Unrestricted
Allow WinRM
ActiveDirectory module


#>
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Порядок выполнения скрипта: 
<# 

-Считывает конфиг с переменными, должен находиться в той-же папке, что и скрипт;
-Формирует список тех серверов, объекты которых включены, обращались к DC не более 14 дней назад, OU которых содержит Servers и не Test;
-Для каждого из полученных серверов по WinRM получить список обновлений;
-Если для сервера не получены значения  в течении  5 минут - продолжить с другим сервером;
-Если во время сбора данных получены ошибки - вывод в log файл;
-Если для сервера полученные обновления старше 3 месяцев, записать информацию о последнем обновлении в csv файл;
-Если для сервера полученные обновления не старше 3 месяцев - продолжить к другому серверу;
-Отправить письмо без авторизации с результатами выполнения.
#>
#-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

$confg = "$PSScriptRoot\GetWinUpdates"+('.txt')
$globallog = "$PSScriptRoot\GetWinUpdates"+('.log')

try

{
    $values = (Get-Content $confg).Replace( '\', '\\') | ConvertFrom-StringData 
    $To = $values.To
    $SmtpServer = $values.SmtpServer
    $SmtpPort = $values.SmtpPort
    $SmtpDomain = $values.SmtpDomain
    $Output = $values.Output
  
}

catch

{   
    Write-Host "No config file has been found" -ForegroundColor RED
    Write-Output $Error[0].Exception.Message
    Write-Host "Script terminated with error." -ForegroundColor Yellow
    Write-Output ("FATAL ERROR - Config file is accessible check not passed  "+(Get-Date)) | Out-File "$globallog" -Append
    Break 
}


$Date = Get-Date -Format "MM.dd.yyyy"
$Unic = Get-WmiObject -Class Win32_ComputerSystem |Select-Object -ExpandProperty "Domain"
$list = Get-ADComputer -Filter * -properties *|Where-Object {$_.enabled -eq $true} | Where-Object {(($_.distinguishedname -like "*Servers*") -and ($_.distinguishedname -notlike "*Test*")) -and ($_.LastLogonDate -ge ((Get-Date).AddDays(-14)))}| Select-Object -ExpandProperty "name"


foreach ($pc in $list) {
$error.Clear()
$Updates = Invoke-Command -ComputerName $pc -ScriptBlock {
    $VerbosePreference='Continue'
    $Session = New-Object -ComObject "Microsoft.Update.Session"
    $Searcher = $Session.CreateUpdateSearcher()

    $historyCount = $Searcher.GetTotalHistoryCount()

    $Searcher.QueryHistory(0, $historyCount) | Select-Object Title, Description, Date,

    @{name="Operation"; expression={switch($_.operation){

    1 {"Installation"}; 2 {"Uninstallation"}; 3 {"Other"}

}}}} -AsJob

Wait-Job $Updates -Timeout 300

if ($Updates.State -eq 'Completed') {
$Updates |select State, Location, PSBeginTime, PSEndTime| Out-File $Output\$Unic"_"$Date"_"Updates_SCRdebug.log -Append
} 
    else {
  $Updates |select State, Location, PSBeginTime, PSEndTime|  Out-File $Output\$Unic"_"$Date"_"Updates_SCRdebug.log -Append
  Stop-Job -Id $Updates.Id
} 
$Result = Receive-Job $Updates
if (($Result| Where-Object {$_.Date -le (Get-Date) -and $_.Date -ge ((Get-Date).AddMonths(-3))}).count -eq 0 ) {
    $Result 4>&1 | Select-Object -Property * -ExcludeProperty RunspaceID, PSShowComputerName | sort Date -desc | select PSComputerName, Title, Description, Date -First 1 |Export-Csv -Append -Delimiter ';' -Path $Output\$Unic"_"$Date"_"Updates_OLD.csv -Encoding UTF8 -NoTypeInformation
    }
    $error | Out-File $Output\$Unic"_"$Date"_"Updates_SCRerrors.log -Append
}

$RunState = @(Get-Content $Output\$Unic"_"$Date"_"Updates_SCRdebug.log |  Where-Object { $_.Contains("Running") } ).Count
$CompleteState = @(Get-Content $Output\$Unic"_"$Date"_"Updates_SCRdebug.log |  Where-Object { $_.Contains("Completed") } ).Count
$FailedState = @(Get-Content $Output\$Unic"_"$Date"_"Updates_SCRdebug.log |  Where-Object { $_.Contains("Failed") } ).Count

$From = (Get-WmiObject -Class Win32_ComputerSystem |Select-Object -ExpandProperty "name")+"@"+$SmtpDomain
$Attachments = (get-childitem $Output\$Unic"_"$Date"_Updates"*.*).fullname
$Subject = $Unic+" servers: whithout update last 3 month"
$Body = "Proceed "+($list.count)+" servers, Completed "+($CompleteState)+", Failed "+($RunState+$FailedState)


Send-MailMessage -To $To -From $From -Subject $Subject -Body $Body -Attachments $Attachments -Port $SmtpPort -SmtpServer $SmtpServer

Remove-Variable -Name * -Force -ErrorAction SilentlyContinue

