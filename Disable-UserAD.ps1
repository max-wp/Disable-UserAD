#Функция отключает учетную запись домена и переносит её в контейнер "уволившиеся сотрудники" с простановкой даты увольнения в описании (отключения)
#Поиск происходит по Фамилии, Имени, Отчеству, табельному номеру или логину. Если пользователь не один, нужно скопировать табельный номер или логин уволившегося сотруднника и вызвать функцию повторно вставив в область поиска новые данные.

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
        Write-Host "Для точной идентификации введите табельный номер или логин отключаемого пользователя" -ForegroundColor 'Red'
        exit
    }

    #Отключаем учетную запись
    Disable-ADAccount -Identity $MUser.SamAccountName# -Confirm
    #Переносим объект учетной записи в новый контейнер - Уволившиеся сотрудники
    Move-ADObject -Identity $MUser.distinguishedName -TargetPath $targetOU
    #Заменяем описание
    Set-ADUser $MUser.SamAccountName -Replace @{description="Уволен $dateStr"}
    
    #Почта
    function Connect-mailServer{
        param($nikMailHide)
        
        function Send-MailMess {
            param($mailAdmins)
            Send-MailMessage -SmtpServer HD-MAIL -To "$mailAdmins" -From 'admin@hydroproject.com' -Subject "Увольнение" -Body "Пользователь $nameMUser уволен.`nПочтовый ящик скрыт из адрессной книги.`nУчетная запись перенесена в группу: Уволенные сотрудники.`n`nСообщение создано автоматически, отвечать на него не нужно!" -Encoding 'UTF8'
        }

        #Подключиться к удаленному серверу Exchange
        $envUserName = $env:UserName
        $currentUser = "mohp.ru\$envUserName"
        if($UserCredential -eq $fulse){
            $global:UserCredential = Get-Credential -Credential $currentUser
        }
        
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'http://hd-mail.mohp.ru/PowerShell/' -Authentication Kerberos -Credential $UserCredential

        Import-PSSession $Session -DisableNameChecking
        
        #Скрываем ящик на сервере, если у учетной записи он существует
        $userMailbox = Get-Mailbox $nikMailHide

        if ($userMailbox -eq $false){
            Set-Mailbox $nikMailHide -HiddenFromAddressListsEnabled $True
        }
        #Обязательное завершение сессии
        Remove-PSSession $Session

        Send-MailMess 'TulpakovMS@hydroproject.com'#, Korneevvv@hydroproject.com

    }

    Connect-mailServer $MUser.mailNickname

    #Диагностическое сообщение об успешности операции
    Write-Host 'Учетная запись отключена' -ForegroundColor Green
    $null = $MUser
    $MUser = Get-ADUser -filter "(Name -like '$search_ADName') -or (extensionAttribute1 -like '$search_ADName') -or (SamAccountName -like '$search_ADName')" -SearchBase "$search_base" -Properties * | Select-Object Name, Enabled, description, distinguishedName
    $MUser
}