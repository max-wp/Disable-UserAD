function MOHPUser {
    param (
        $switch = 'disable',
        $user
    )
    


    function Connect-MOHPmailServer{

        try {

        #Если учетные данные еще не вводились, запрашиваем их
        if( $null -eq $UserCredential){
            #Учетная запись
            $envUserName = $env:UserName
            $currentUser = "mohp.ru\$envUserName"
            $global:UserCredential = Get-Credential -Credential $currentUser -ErrorAction Stop
        }

        #Подключаемся к серверу Exchange если нет открытых сессий
        if ($($Session.State) -eq 'Closed' -or $null -eq $Session){
            $Global:Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'http://hd-mail.mohp.ru/PowerShell/' -Authentication Kerberos -Credential $UserCredential -ErrorAction Stop
            $stateSes = $Session.State
        }
        
        if($stateSes -eq 'Opened'){
            Import-PSSession $Session -DisableNameChecking > $null
        }
        if ($null -eq $($Session.State)){
            Write-Host "Статус соединения с почтовым сервером: Сервер не найден!"
            exit
        }
        Write-Host "Статус соединения с почтовым сервером: Session-state: $($Session.State)"

        
    } catch {
        
            #OUTPUT
            $global:UserCredential = $null
            Write-Host "Ошибка: Неверное имя пользователя или пароль" -ForegroundColor 'Red'
            #Write-Host "Детали ошибки: $_.Exception.Message"
            exit
        }
    }

    #Закрывает все открытые сессии
    function Disconnect-MOHPmailServer {

        $Session = Get-PSSession
        if ($null -eq $Session){
            Write-Host "Открытые подключения не найдены!"
            exit
        }

        Remove-PSSession $Session
        $stateSes = $Session.State
        Write-Host "Статус соединения с почтовым сервером: Session-state: $stateSes"
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

        function Set-PropertiesAD{
            #Отключаем учетную запись
            Disable-ADAccount -Identity $MUser.SamAccountName# -Confirm
            #Переносим объект учетной записи в новый контейнер - Уволившиеся сотрудники
            Move-ADObject -Identity $MUser.distinguishedName -TargetPath $targetOU
            #Заменяем описание
            Set-ADUser $MUser.SamAccountName -Replace @{description="Перенесена в группу уволенных $dateStr"}
        }
        #Почта
        function Send-MailMess {
            param($mailAdmin)
            Send-MailMessage -SmtpServer HD-MAIL -To "$mailAdmin" -From 'admin@hydroproject.com' -Subject "Увольнение" -Body "Пользователь $nameMUser уволен.`nПочтовый ящик скрыт в адресной книге.`nУчетная запись перенесена в группу: Уволенные сотрудники.`n`nСообщение создано автоматически, отвечать на него не нужно!" -Encoding 'UTF8' -ErrorAction Stop
        }
        function Set-PropertiesMailBox {
            param (
                $nikMailHide
            )

            $userMailbox = Get-Mailbox $nikMailHide
            #Скрываем ящик на сервере, если у учетной записи он существует
            if ($userMailbox -eq $false){
                Set-Mailbox $nikMailHide -HiddenFromAddressListsEnabled $True
            }
        }

        Set-PropertiesAD
        Connect-MOHPmailServer
        #Операции выполняемые на сервере Exchange
        Set-PropertiesMailBox $MUser.mailNickname
        
        #Отправка сообщений
        switch ($env:UserName) {
            'oit_tulpakov' {Send-MailMess  'Korneevvv@hydroproject.com'}
            'oit_korneev' {Send-MailMess  'TulpakovMS@hydroproject.com'}
            Default {Send-MailMess  'Korneevvv@hydroproject.com'; Send-MailMess  'TulpakovMS@hydroproject.com'}
        }

        #Диагностическое сообщение об успешности операции
        Write-Host "`nУчетная запись отключена" -ForegroundColor Green
        $null = $MUser
        $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, description, distinguishedName
        $MUser

        #Обязательно закрываем сессию с почтовым сервером
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
        function Set-PropertiesAD{
            #Включаем учетную запись
            Enable-ADAccount -Identity $MUser.SamAccountName# -Confirm
            #Переносим объект учетной записи в новый контейнер - Уволившиеся сотрудники
            #Move-ADObject -Identity $MUser.distinguishedName -TargetPath $targetOU
            #Заменяем описание
            Set-ADUser $MUser.SamAccountName -Replace @{description="Учетка включена $dateStr"}
        }
        #Почта
        function Send-MailMess {
            param($mailAdmin)
            Send-MailMessage -SmtpServer HD-MAIL -To "$mailAdmin" -From 'admin@hydroproject.com' -Subject "Возвращение ранее уволенного" -Body "Пользователь $nameMUser восстановлен`nПочтовый ящик возвращен в адресную книгу.`n`nСообщение создано автоматически, отвечать на него не нужно!" -Encoding 'UTF8' -ErrorAction Stop
        }
        function Set-PropertiesMailBox {
            param (
                $nikMailHide
            )

            $userMailbox = Get-Mailbox $nikMailHide
            #Возвращаем ящик в адресную книгу на сервере, если у учетной записи он существует
            if ($userMailbox -eq $false){
                Set-Mailbox $nikMailHide -HiddenFromAddressListsEnabled $false
            }
        }
        Set-PropertiesAD
        Connect-MOHPmailServer
        #Операции выполняемые на сервере Exchange
        Set-PropertiesMailBox $MUser.mailNickname
        
        #Отправка сообщений
        switch ($env:UserName) {
            'oit_tulpakov' {Send-MailMess  'Korneevvv@hydroproject.com'}
            'oit_korneev' {Send-MailMess  'TulpakovMS@hydroproject.com'}
            Default {Send-MailMess  'Korneevvv@hydroproject.com'; Send-MailMess  'TulpakovMS@hydroproject.com'}
        }


        #Диагностическое сообщение об успешности операции
        Write-Host "`nУчетная запись включена" -ForegroundColor Green
        $null = $MUser
        $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, description, distinguishedName
        $MUser

        #Обязательно закрываем сессию с почтовым сервером
        Disconnect-MOHPmailServer
    }

    switch ($switch) {
        'disable' { Disable-MOHPUser $user}
        'enable' { Enable-MOHPUser $user}
        Default {}
    }

}