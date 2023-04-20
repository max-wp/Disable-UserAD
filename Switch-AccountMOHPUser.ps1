function Connect-MOHPmailServer{
    Param($adminUser)
    #Подключаемся к удаленному серверу Exchange
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'http://hd-mail.mohp.ru/PowerShell/' -Authentication Kerberos -Credential $adminUser

    $stateSes = $Session.State
    Import-PSSession $Session -DisableNameChecking
    Write-Host "Статус соединения с почтовым сервером: Session-state: $stateSes"
}
function Disconnect-MOHPmailServer {

    $Session = Get-PSSession
    Remove-PSSession $Session
    $stateSes = $Session.State
    Write-Host "Статус соединения с почтовым сервером: Session-state: $stateSes"
}

#ПРОВЕРКА на блокирующие ошибки
function Test-MOHPAccount{
    try {
        #Учетная запись
        $envUserName = $env:UserName
        $currentUser = "mohp.ru\$envUserName"
        #Если ввод данных был проигнорирован
        if($UserCredential -eq $null){
        $global:UserCredential = Get-Credential -Credential $currentUser -ErrorAction Stop
        Connect-mailServer $UserCredential
        }
        #Операции выполняемые на сервере Exchange
        #Set-SettingsMailBox $MUser.mailNickname
        #Send-MailMess 'TulpakovMS@hydroproject.com'
        #Send-MailMess  'Korneevvv@hydroproject.com'

    } catch {
    
        #OUTPUT
        Write-Host "Ошибка: Неверное имя пользователя или пароль" -ForegroundColor 'Red'#$_.Exception.Message
        Write-Host "Детали ошибки: $_.Exception.Message"
        exit
    }
}

#ОТКЛЮЧЕНИЕ УЧЕТНОЙ ЗАПИСИ УВОЛЕННОГО СОТРУДНИККА.
function Disable-MOHPUser([parameter (Mandatory=$true, HelpMessage='Введите Фамилию, логин или табельный номер пользователя')][string]$ADNameUser){
    #Область поиска
    $search_base='DC=mohp,DC=ru'
    $search_ADName = "*$ADNameUser*"
    $dateStr = Get-Date -Format "yyyy-MM-dd"
    
    #Группа для переноса
    $targetOU = 'OU=уволившиеся сотрудники,DC=MOHP,DC=RU'

    $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, extensionAttribute1, Department, title, SamAccountName, distinguishedName, mailNickname, description

    $nameMUser = $MUser.Name
    if (-not $MUser){
    
        Write-Host 'Пользователь не найден' -ForegroundColor 'Red'
        exit
    }
    
    $userCount = $MUser.count
    if ($userCount -ge 1){
        
        Write-Host "`nНайдено $userCount сотрудников:" -ForegroundColor 'Green'
        $Muser
        Write-Host "Для точной идентификации введите табельный номер или логин отключаемого пользователя, с повторным вызовом функции. `nПример: Disable-MOHPUser Пупкин" -ForegroundColor 'Red'
        exit
    }

    #Отключаем учетную запись
    Disable-ADAccount -Identity $MUser.SamAccountName# -Confirm
    #Переносим объект учетной записи в новый контейнер - Уволившиеся сотрудники
    Move-ADObject -Identity $MUser.distinguishedName -TargetPath $targetOU
    #Заменяем описание
    Set-ADUser $MUser.SamAccountName -Replace @{description="Уволен $dateStr"}
    
    #Почта
    function Send-MailMess {
        param($mailAdmin)
        Send-MailMessage -SmtpServer HD-MAIL -To "$mailAdmin" -From 'admin@hydroproject.com' -Subject "Увольнение" -Body "Пользователь $nameMUser уволен.`nПочтовый ящик скрыт из адрессной книги.`nУчетная запись перенесена в группу: Уволенные сотрудники.`n`nСообщение создано автоматически, отвечать на него не нужно!" -Encoding 'UTF8' -ErrorAction Stop
    }
    function Set-SettingsMailBox {
        param (
            $nikMailHide
        )

        $userMailbox = Get-Mailbox $nikMailHide
        #Скрываем ящик на сервере, если у учетной записи он существует
        if ($userMailbox -eq $false){
            Set-Mailbox $nikMailHide -HiddenFromAddressListsEnabled $True
        }
    }

    Test-MOHPAccount

    #Диагностическое сообщение об успешности операции
    Write-Host "`nУчетная запись отключена" -ForegroundColor Green
    $null = $MUser
    $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, description, distinguishedName
    $MUser

    #Обязательно закрываем сессию с почтовым сервером при запуске скрипта. С открытой сессией данная фукнция уже не видна
    Disconnect-MOHPmailServer

}

#ВКЛЮЧЕНИЕ УЧЕТНОЙ ЗАПИСИ УВОЛЕННОГО СОТРУДНИККА.

function Enable-MOHPUser([parameter (Mandatory=$true, HelpMessage='Введите Фамилию, логин или табельный номер пользователя')][string]$ADNameUser){
    #Область поиска
    $search_base='DC=mohp,DC=ru'
    $search_ADName = "*$ADNameUser*"
    $dateStr = Get-Date -Format "yyyy-MM-dd"
    
    #Группа для переноса
    #$targetOU = 'OU=уволившиеся сотрудники,DC=MOHP,DC=RU'

    $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, extensionAttribute1, Department, title, SamAccountName, distinguishedName, mailNickname, description

    $nameMUser = $MUser.Name
    if (-not $MUser){
    
        Write-Host 'Пользователь не найден' -ForegroundColor 'Red'
        exit
    }
    
    $userCount = $MUser.count
    if ($userCount -ge 1){
        
        Write-Host "`nНайдено $userCount сотрудников:" -ForegroundColor 'Green'
        $Muser
        Write-Host "Для точной идентификации пользователя введите табельный номер или логин, с повторным вызовом функции. `nПример: Disable-MOHPUser Пупкин" -ForegroundColor 'Red'
        exit
    }

    #Включаем учетную запись
    Enable-ADAccount -Identity $MUser.SamAccountName# -Confirm
    #Переносим объект учетной записи в новый контейнер - Уволившиеся сотрудники
    #Move-ADObject -Identity $MUser.distinguishedName -TargetPath $targetOU
    #Заменяем описание
    Set-ADUser $MUser.SamAccountName -Replace @{description="Принят $dateStr"}
    
    #Почта
    function Send-MailMess {
        param($mailAdmin)
        Send-MailMessage -SmtpServer HD-MAIL -To "$mailAdmin" -From 'admin@hydroproject.com' -Subject "Возвращение ранее уволенного" -Body "Пользователь $nameMUser восстановлен`nПочтовый ящик возвращен в адрессную книгу.`n`nСообщение создано автоматически, отвечать на него не нужно!" -Encoding 'UTF8' -ErrorAction Stop
    }
    function Set-SettingsMailBox {
        param (
            $nikMailHide
        )

        $userMailbox = Get-Mailbox $nikMailHide
        #Возвращаем ящик в адресную книгу на сервере, если у учетной записи он существует
        if ($userMailbox -eq $false){
            Set-Mailbox $nikMailHide -HiddenFromAddressListsEnabled $false
        }
    }

    Test-MOHPAccount

    #Диагностическое сообщение об успешности операции
    Write-Host "`nУчетная запись включена" -ForegroundColor Green
    $null = $MUser
    $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, description, distinguishedName
    $MUser

    #Обязательно закрываем сессию с почтовым сервером при запуске скрипта. С открытой сессией данная фукнция уже не видна
    Disconnect-MOHPmailServer
}