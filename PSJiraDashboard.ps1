cls 
[reflection.assembly]::LoadWithPartialName("system.windows.forms") | Out-Null #Подгрузка библиотек работы с формами
[reflection.assembly]::LoadWithPartialName("system.data")| Out-Null #Подгрузка библиотек работы с данными
$Global:Root = $PSScriptRoot #Каталог скрипта

Import-Module -Name "$PSScriptRoot\System\JiraPS"
Import-Module -Name "$PSScriptRoot\System\ShowUI"

Set-JiraConfigServer -server "http://servicedesk:8080" # Настройка на наш сервер Jira

#Get-JiraIssue -Key "pr-66"|Invoke-JiraIssueTransition -Transition 141 Нажатие на кнопку!!! В разных заявках поразному, я сделал по имени кнопки like "работ"
#Уточнить                  141 Уточнение                
#Отклонить                 191 Отклонено                
#Передать на Внешнюю линию 201 Передать на Внешнюю линию
#Выполнено                 231 Выполнено                
#Отправить на согласование 351 На согласовании          
#Начать работу             221 В работе
#В работу                  81  В работе 
#В работу                  11  В работе

#Задаем размер буфера (т.к. размер_окна <= размер_буфера)
$bsize = $Host.UI.RawUI.BufferSize
$bsize.Width = 200
$Host.UI.RawUI.BufferSize = $bsize
#Задаем размер окна
$wsize = $Host.UI.RawUI.WindowSize
$wsize.Width = 200
$wsize.Height = 35
$Host.UI.RawUI.WindowSize = $wsize
#Задаем позицию окна консоли (не сработает, если задано системно)
$wpos = $Host.UI.RawUI.WindowPosition
$wpos.X = 0
$wpos.Y = 0
$Host.UI.RawUI.WindowPosition = $wpos
#Задаем заголовок окна
$Host.UI.RawUI.WindowTitle = 'PSJiraDashboard Технический эксперт Кемерово'
#Очищаем консоль
Clear-Host

$DPCLogin = $env:USERNAME #пользователь Jira (7701-)
$DPCPass = Read-Host "Введите пароль от своей УЗ в домене DPC: " -AsSecureString # Пароль
$Global:DPCCred =  New-Object System.Management.Automation.PSCredential -ArgumentList $DPCLogin, $DPCPass # Переменая учётных данных
$DateTime = Get-Date -Format dd.MM.yyyy #Текущая дата
$ConfFile = "$root\Conf_$DateTime.csv"
$SmenaLog = if (!(test-path "$root\Smena_$DateTime.csv"))
{
    New-Item -Path "$root\" -Name "Smena_$DateTime.csv"|Out-Null; 
    "$root\Smena_$DateTime.csv"
}
else 
{
    "$root\Smena_$DateTime.csv"
} #Файлик с заявками принятыми в работу в течении смены

Function Take-NewJiraItem ($Credentials, $Issues,$who ) #Взятие переданных заявок в работу или на себя
{
    New-JiraSession -Credential $Credentials |Out-Null #Открываем сессию Jira
    $ETSpis = $Issues
    if ($ETSpis -ne $null) {
        $ETSpis|foreach {
            if (($_.status -eq "Зарегистрировано") -or ($_.status -eq "Согласовано") -or ($_.status -eq "На исполнение")) { #новая заявка
                if (($_|Select-Object -ExpandProperty transition).name -imatch "На исполнение") 
                {
                    $action =$_|Select-Object -ExpandProperty transition |where name -ieq "На исполнение"
                    Invoke-JiraIssueTransition -Issue $_  -Transition $action > $null #Взятие TR "на исполнение"
                }
                else
                {
                    write-host "Взята в работу заявка $_.key" -f Green
                    $action =$_|Select-Object -ExpandProperty transition |where resultstatus -Match "работ" # #Получение ID значение взятия в работу
                    Invoke-JiraIssueTransition -Issue $_  -Transition $action > $null #Взятие зарегистрированной заявки в работу
                    Set-JiraIssue -Issue $_ -Assignee $Credentials.UserName > $null #Назначение заявки на "Меня"
                    SaveToJLog -jiraIsues $_ -who $who
                }
            } 
            elseif (($_.status -ne "Зарегистрировано") -and ($_.status -ne "Согласовано"))  #Заявки в работе попавшие на ТЭ выбранного города
            {
                write-host "Назначена на меня $_.key" -f Green
                Set-JiraIssue -Issue $_ -Assignee $Credentials.UserName > $null #Назначение заявки на "Меня"
                SaveToJLog -jiraIsues $_ -who $who
            } 
            Check-SLA -JiraIssues $_ #Проверка заявки на истечение SLA
        }
    }
    Get-JiraSession | Remove-JiraSession # Закрываем сессию Jira
}

Function Show-JiraWin ($JiraIssues,$who) { #Отображение Win-окна в новыми заявками поверх всех окон, что бы точно увидеть!
    Show-UI -ScriptBlock {
        New-Grid -Columns ('Auto', '1*') -Rows ('1*','Auto') -Resource @{
            'Import-ModuleData' = {
                $modules = @($JiraIssues)       
                foreach ($m in $modules) {
                    New-TreeViewItem -Header ($m.Key+' '+$m.Summary)  -FontSize 16 -DataContext $m -ItemsSource @( 
                        $m.Status; $m.Description ; $m.Updated
                    ) 
                }                            
            }
            'JiraIss'=$JiraIssues
            'Who'=$who
        } {
            New-TreeView -FontSize 24 -Name NewIssues -On_loaded {        
                ${Import-ModuleData} | Add-ChildControl -parent $this -Clear
            } -On_SelectedItemChanged { }    

            New-UniformGrid -Row 1 -ColumnSpan 2 -Columns 2 { #-Columns 3 - количество колонок для кнопок
                New-Button -FontSize 18 -Row 1 -Name "TakeIssues" "_Взять заявки в работу" -On_Click {
                    Take-NewJiraItem -Credentials $Global:DPCCred -Issues ${JiraIss} -Who ${Who}
                    $window.Close() #Закрытие формы
                }
                New-Button -FontSize 18 -Row 1 -Column 2 -Name "CloseForm" "_Закрыть" -On_Click { #Кнопка закрыть
                    $window.Close() #Закрытие формы
                }
            }
            New-Border -Name HelpContainer -Column 1 #>
        }        
    } -Topmost
} 

Function Format-Color([hashtable] $Colors = @{}, [switch] $SimpleMatch) { #Арабская функция раскраски табличного вывода (взрывоопасно!)
	$lines = ($input | Out-String) -replace "`r", "" -split "`n"
	foreach($line in $lines) {
		$color = ''
		foreach($pattern in $Colors.Keys){
			if(!$SimpleMatch -and $line -match $pattern) { $color = $Colors[$pattern] }
			elseif ($SimpleMatch -and $line -like $pattern) { $color = $Colors[$pattern] }
		}
		if($color) {
			Write-Host -ForegroundColor $color $line
		} else {
			Write-Host $line
		}
	}
}

Function Take-IsMyLastComment ($JiraIssues) #true\false Проверка, что последний комментарий "мой", для проверки перед сдачей смены
{
    $tmp = Get-JiraIssueComment $JiraIssues
    if ($tmp -ne $null) {
        if(((($tmp[-1].Created).ToShortDateString()) -eq (Get-Date -Format dd.MM.yyyy).ToString()) -and (($tmp[-1].Author).ToString() -eq $DPCCred.UserName)){$true} else {$false}
    }
    else {$false}
}

Function Check-SLA ($JiraIssues) #Проверка на оставшийся SLA, если будет меньше 1 часа выдаёт сообщение.
{
    if (($JiraIssues.customfield_13402.ongoingCycle.remainingTime.millis/1000/60/60 -le 1) -and ($JiraIssues.customfield_13402.ongoingCycle.remainingTime.millis/1000/60/60 -ne 0))
    {
        [System.Reflection.assembly]::LoadWithPartialName("System.Windows.Forms")
        $result = [System.Windows.Forms.MessageBox]::Show("В заявке $($JiraIssues.key) истекает SLA","Внимание!", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        $false
    }    
    else {$true}
}

Function SaveToJLog ($jiraIsues,$who) #Ведение отчёта в течении "смены" отчёт лежит в каталоге со скриптом, вида Smena_05.06.2018.csv 
{
    $tlog = Get-Content -Path $SmenaLog
    if (!(Test-Path $ConfFile)) {$slog = $null} else {$slog = Get-Content -Path $ConfFile}
    $jiraIsues|ForEach-Object { 
        $ttime = (Get-Date -Format "dd.MM.yyyy HH:mm").ToString()
        $_|Add-Member -Name TakeTime -MemberType NoteProperty -Value $ttime
        $_|Add-Member -Name Who -MemberType NoteProperty -Value $who

#customfield_12401 #Контур
#customfield_12404 #Система
#customfield_12811 #Подсистема

        if ($tlog -eq $null)
        {
            $_|select key,summary,customfield_12401,customfield_12404,customfield_12811,TakeTime|ConvertTo-Csv -NoTypeInformation -Delimiter ';'|Add-Content -Path $SmenaLog -Force
        }
        elseif (($tlog|ConvertFrom-Csv -Delimiter ';').key -inotcontains $_.key)
        {
            $_|select key,summary,customfield_12401,customfield_12404,customfield_12811,TakeTime|ConvertTo-Csv -NoTypeInformation -delimiter ';'|select -Skip 1|Add-Content -Path $SmenaLog -Force
        }
        
        if ($slog -eq $null)
        {
            $_|select key,summary,customfield_12401,customfield_12404,customfield_12811,who|ConvertTo-Csv -NoTypeInformation -Delimiter ';'|Add-Content -Path $ConfFile -Force
        }
        elseif (($slog|ConvertFrom-Csv -Delimiter ';').key -inotcontains $_.key)
        {
            $_|select key,summary,customfield_12401,customfield_12404,customfield_12811,who|ConvertTo-Csv -NoTypeInformation -delimiter ';'|select -Skip 1|Add-Content -Path $ConfFile -Force
        }
    } 
}

Function RemFromLog ($jiraIsues) #Ведение отчёта в течении "смены" отчёт лежит в каталоге со скриптом, вида Smena_05.06.2018.csv 
{
    $slog = Get-Content -Path $ConfFile|ConvertFrom-Csv -Delimiter ';'
    $jiraIsues|ForEach-Object {
        $t=$_
        $slog = $slog|where {$_.key -ne $t}
    }
    Remove-Item -Path $ConfFile
    $slog|select key,summary,customfield_12401,customfield_12404,customfield_12811|ConvertTo-Csv -NoTypeInformation -Delimiter ';'|Add-Content -Path $ConfFile -Force
}

Function GetIssuesByKey ($Credentials,$IssueKeys)
{
    New-JiraSession -Credential $Credentials |Out-Null #Открываем сессию Jira
    $res = $IssueKeys|foreach {Get-JiraIssue -Key $_}
    Get-JiraSession | Remove-JiraSession # Закрываем сессию Jira
    $res
}

Function IssuesOnMe ($MyIssues, $Who) #Проверка заявок на мне на предмет новых заявок на мне или закрытых заявок
{
    if (Test-Path $ConfFile)
    {
        $TIssues = Get-Content -Path $ConfFile|ConvertFrom-Csv -Delimiter ';'
        $lostIssues = (Compare-Object -ReferenceObject $MyIssues.Key -DifferenceObject $TIssues.Key -IncludeEqual| where {$_.sideindicator -eq '=>'}).Inputobject #заявки которые на мне были а сейчас нет
        $newIssues = (Compare-Object -ReferenceObject $MyIssues.Key -DifferenceObject $TIssues.Key -IncludeEqual|where {$_.sideindicator -eq '<='}).Inputobject #новые заявки на мне которых не было
        if ($lostIssues -ne $null)
        {
            RemFromLog -jiraIsues $lostIssues
        }
    }
    else
    {
        $newIssues = $MyIssues.key
    }

    if ($newIssues -ne $null)
    {
        $tIs = GetIssuesByKey -Credentials $DPCCred -IssueKeys $newIssues
        Show-JiraWin -JiraIssues $tIs -Who $Who
        Write-Host "Новые заявки на мне!" -f Red
    }
}

#Начало работы
New-JiraSession -Credential $DPCCred |Out-Null
$MySpis = Get-JiraIssue -Query "assignee = currentUser() AND resolution = Unresolved order by updated DESC" #Проверяем заявки на себе
if (!(Test-Path $ConfFile))
{
    Write-host "Нет конфиг файла" -f Yellow
    IssuesOnMe -MyIssues $MySpis
}
else
{
    Write-host "Есть конфиг файл" -f Yellow
    $Cspis = Get-Content $ConfFile|ConvertFrom-Csv -Delimiter ';'
    if ($MySpis -ne $null) {$MySpis|where {$Cspis.key -inotcontains $_.key}|select key,summary,customfield_12401,customfield_12404,customfield_12811|ConvertTo-Csv -NoTypeInformation -delimiter ';'|select -Skip 1|Add-Content -Path $ConfFile} #Добавление строк без заголовков
}
Get-JiraSession | Remove-JiraSession
################

while ($true) # Цикл до покоса!
{
    New-JiraSession -Credential $DPCCred |Out-Null
    #Заявки на мне    
    $MySpis = Get-JiraIssue -Query "assignee = currentUser() AND status not in (Отказ) AND resolution = Unresolved order by updated DESC" #Проверяем заявки на себе
    $myspis | Add-Member -Name IsMyComLast -MemberType NoteProperty -Value '' #Добавляем переменную на проверку "моего" последнего комментария
    $myspis |foreach {$_.IsMyComLast = Take-IsMyLastComment -JiraIssues $_} #применяем

    cls
    if ($MySpis -ne $null) {IssuesOnMe -MyIssues $MySpis -Who "My"}

    #Заявки на ТЭ Кемерово
    $ETSpis =Get-JiraIssue -query "status not in (Отклонено, Выполнено, Закрыто, Отказ) AND assignee in (techexpert42)"  #Сюда пишется общий пользователь техэкспертов по региону
    if ($ETSpis -ne $null) # Если есть новые заявки на ТЭ 
    {
        Show-JiraWin -JiraIssues $ETSpis -who "techexpert42" #Сюда пишется общий пользователь техэкспертов по региону
        Write-Host "Новые заявки на ТЭ (Выбранного города)" -f Red #Выводим список новых заявок
        $ETSpis
    }
    Get-JiraSession | Remove-JiraSession

    Write-Host "Заявки на мне:" -f Yellow
    $myspis|select key, summary, status, IsMyComLast | format-color @{'True'='Green';'False'='Red'} # Если есть заявки в которых "Я" не сделал последний комментарий подсвечены красным, иначе зеленым
    Start-Sleep -Seconds 40 #Проверяем каждые 40 секунд
}