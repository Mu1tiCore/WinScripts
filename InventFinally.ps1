# запуск от имени администратора
# enable-psremoting -force     -для настройки WinRM на компьютерах
# Set-ExecutionPolicy RemoteSigned
# наладить выборку списка адресов!

#принимать 3 параметра (путь куда выводить,пароль,где брать айпи адреса)!!!!!!!!!! 


<# Comment

$file - адрес временной папки с IP адресами
$ip* - список адресов 
$ip - переменная цикла, коллекция адресов
$cred - учетные данные администратора
$q - имя комьютера в сети
$a - имя файла инвентаризации данного компьютера
$mas - массив данных инвентаризации
$i - переменная, с которой работает цикл foreach
$0..24 - данные о компьютере пользователя
$Hardware - аппаратная часть 
$Software - программная часть
$result - переменная цикла заполнения таблицы
$Row - строка книги excel
$Column - столбец книги excel
$BaseRow - переменная цикла, для перехода на строку до цикла заполнения
$path - переменная, определяющая директорию сохранения данных инвентаризации
$res - переменная Win Form
$form - Win Form
$textBox - поле Win Form
#>

#Автоматическое повышение до роли администратора!
<#
# Get the ID and security principal of the current user account
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent();
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID);

# Get the security principal for the administrator role
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator;

# Check to see if we are currently running as an administrator
if ($myWindowsPrincipal.IsInRole($adminRole))
{
    # We are running as an administrator, so change the title and background colour to indicate this
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)";
    $Host.UI.RawUI.BackgroundColor = "DarkBlue";
    Clear-Host;
}
else {
    # We are not running as an administrator, so relaunch as administrator

    # Create a new process object that starts PowerShell
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";

    # Specify the current script path and name as a parameter with added scope and support for scripts with spaces in it path
    $newProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"

    # Indicate that the process should be elevated
    $newProcess.Verb = "runas";

    # Start the new process
    [System.Diagnostics.Process]::Start($newProcess);

    # Exit from the current, unelevated, process
    Exit;
}
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form 
$form.Text = "Вводная информация"
$form.Size = New-Object System.Drawing.Size(400,200) 
$form.StartPosition = "CenterScreen"

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20) 
$label.Size = New-Object System.Drawing.Size(280,40) 
$label.Text = "Расположение файла с адресами (полный путь) Пример: C:\Users\*username*\Desktop\*.txt"
$form.Controls.Add($label) 

$textBox = New-Object System.Windows.Forms.TextBox 
$textBox.Location = New-Object System.Drawing.Point(10,60) 
$textBox.Size = New-Object System.Drawing.Size(260,20) 
$form.Controls.Add($textBox) 

$form.Topmost = $True

$form.Add_Shown({$textBox.Select()})
$res = $form.ShowDialog()

if ($res -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    $x
}
$file = $textBox.Text



Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form 
$form.Text = "Вводная информация"
$form.Size = New-Object System.Drawing.Size(400,200) 
$form.StartPosition = "CenterScreen"

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(75,120)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Location = New-Object System.Drawing.Point(150,120)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $CancelButton
$form.Controls.Add($CancelButton)

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20) 
$label.Size = New-Object System.Drawing.Size(280,40) 
$label.Text = "Директория сбора информации (полный путь) Пример: C:\Users\*username*\Desktop\new"
$form.Controls.Add($label) 

$textBox = New-Object System.Windows.Forms.TextBox 
$textBox.Location = New-Object System.Drawing.Point(10,60) 
$textBox.Size = New-Object System.Drawing.Size(260,20) 
$form.Controls.Add($textBox) 

$form.Topmost = $True

$form.Add_Shown({$textBox.Select()})
$res = $form.ShowDialog()

if ($res -eq [System.Windows.Forms.DialogResult]::OK)
{
    $x = $textBox.Text
    $x
}
$pat = $textBox.Text
#создание вложенной папки
$dat = Get-Date -Format d
$path = "$pat\$dat"
New-Item -Path $path -ItemType Directory

Write-Host "Пароль администратора"
$cred = (new-object -typename System.Management.Automation.PSCredential -argumentlist "office\wsadmin",(Read-Host -AsSecureString -asplaintext -force))




[string[]]$ip = get-content $file

foreach ($i in $ip)
{
Write-Host "Проверяется"$i
Set-Item  wsman:\localhost\client\trustedhosts -value $i -force
If ((Test-Connection $i -count 4 -Quiet) -eq "True")
{

function Generic
{
  Invoke-Command -computername $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_operatingsystem|select-object csname -ExpandProperty csname}
}
$q = Generic
$a = @($q,$i)
 
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $true 
$WorkBook = $Excel.Workbooks.Add() 
$Hardware = $WorkBook.Worksheets.Item(1)
$Hardware.name = 'Аппаратная часть' 
$Row = 1
$Column = 1



$Hardware.Cells.Item($Row,$Column) = 'Процессор'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Наименование'
$Hardware.Cells.Item($Row,2) = 'Сокет'
$Hardware.Cells.Item($Row,3) = 'Описание'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$0 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Processor|Select-Object Name -ExpandProperty Name}
$1 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Processor|Select-Object SocketDesignation -ExpandProperty SocketDesignation}
$2 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Processor|Select-Object Caption -ExpandProperty Caption}
Foreach ($result in $0)
{
    $Hardware.Cells.Item($Row, $column) = $result
    $Row++
}
$Row=$BaseRow
$Column++
Foreach ($result in $1)
{
    $Hardware.Cells.Item($Row, $column) = $result
    $Row++
}
$Row=$BaseRow
$Column++
Foreach ($result in $2)
{
    $Hardware.Cells.Item($Row, $column) = $result
    $Row++
}
$Column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Материнская плата'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Производитель'
$Hardware.Cells.Item($Row,2) = 'Модель'
$Hardware.Cells.Item($Row,3) = 'Серийный номер'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$3 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Baseboard|Select-Object Manufacturer -ExpandProperty Manufacturer}
$4 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Baseboard|Select-Object Product -ExpandProperty Product}
$5 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Baseboard|Select-Object SerialNumber -ExpandProperty SerialNumber}

$Hardware.Cells.Item($Row, $column) = $3
$column++
$Hardware.Cells.Item($Row, $column) = $4
$column++
$Hardware.Cells.Item($Row, $column) = $5
$Row++
$column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Видеокарта'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Наименование'
$Hardware.Cells.Item($Row,2) = 'Объем памяти (мб)'
$Hardware.Cells.Item($Row,3) = 'Видеопроцессор'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$6 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_VideoController|Select-Object Name -ExpandProperty Name}
$7 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_VideoController|Select-Object AdapterRam -ExpandProperty AdapterRam}
$8 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_VideoController|Select-Object VideoProcessor -ExpandProperty VideoProcessor}

$Hardware.Cells.Item($Row, $column) = $6
$column++
$Hardware.Cells.Item($Row, $column) = $7/1Mb
$column++
$Hardware.Cells.Item($Row, $column) = $8
$Row++
$column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Оперативная память'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Размер (мб)'
$Hardware.Cells.Item($Row,2) = 'Расположение'
$Hardware.Cells.Item($Row,3) = 'Производитель'
$Hardware.Cells.Item($Row,4) = 'Модель'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$9 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Physicalmemory|Select-Object Capacity -ExpandProperty Capacity}
$10 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Physicalmemory|Select-Object DeviceLocator -ExpandProperty DeviceLocator}
$11 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Physicalmemory|Select-Object Manufacturer -ExpandProperty Manufacturer}
$12 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Physicalmemory|Select-Object PartNumber -ExpandProperty PartNumber}
Foreach ($result in $9)
{
    $Hardware.Cells.Item($Row, $column) = $result/1Mb
    $Row++
}
$Row=$BaseRow
$Column++
Foreach ($result in $10)
{
    $Hardware.Cells.Item($Row, $column) = $result
    $Row++
}
$Row=$BaseRow
$Column++
Foreach ($result in $11)
{
    $Hardware.Cells.Item($Row, $column) = $result
    $Row++
}
$Row=$BaseRow
$Column++
Foreach ($result in $12)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Жесткие диски'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Модель'
$Hardware.Cells.Item($Row,2) = 'Колличество разделов'
$Hardware.Cells.Item($Row,3) = 'Интерфейс'
$Hardware.Cells.Item($Row,4) = 'Размер (гб)'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$13 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DiskDrive|Select-Object Model -ExpandProperty Model}
$14 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DiskDrive|Select-Object Partitions -ExpandProperty Partitions}
$15 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DiskDrive|Select-Object Interfacetype -ExpandProperty Interfacetype}
$16 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DiskDrive|Select-Object Size -ExpandProperty Size}
Foreach ($result in $13)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $14)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $15)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $16)
{
    $Hardware.Cells.Item($Row,$Column) = $result/1Gb
    $Row++
}
$Column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Монитор'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Наименование'
$Hardware.Cells.Item($Row,2) = 'Высота (точек)'
$Hardware.Cells.Item($Row,3) = 'Ширина (точек)'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$17 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DesktopMonitor|Select-Object Name -ExpandProperty Name}
$18 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DesktopMonitor|Select-Object ScreenHeight -ExpandProperty ScreenHeight}
$19 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_DesktopMonitor|Select-Object ScreenWidth -ExpandProperty ScreenWidth}
Foreach ($result in $17)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $18)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $19)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Клавиатура'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Наименование'
$Hardware.Cells.Item($Row,2) = 'Идентификатор'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$20 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Keyboard|Select-Object Name -ExpandProperty Name}
$21 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Keyboard|Select-Object DeviceID -ExpandProperty DeviceID}

$Hardware.Cells.Item($Row, $column) = $20
$column++
$Hardware.Cells.Item($Row, $column) = $21
$Row++
$column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Координатное устройство (мышь)'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Наименование'
$Hardware.Cells.Item($Row,2) = 'Идентификатор'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$22 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_PointingDevice|Select-Object Name -ExpandProperty Name}
$23 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_PointingDevice|Select-Object DeviceID -ExpandProperty DeviceID}

$Hardware.Cells.Item($Row, $column) = $22
$column++
$Hardware.Cells.Item($Row, $column) = $23
$Row++
$column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Сетевая карта'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Шлюз по умолчанию'
$Hardware.Cells.Item($Row,2) = 'Сетевой адрес'
$Hardware.Cells.Item($Row,3) = 'Домен'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$24 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" |Select-Object DefaultIPGateway -ExpandProperty DefaultIPGateway}
$25 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" |Select-Object IPAddress -ExpandProperty IPAddress}
$26 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_NetworkAdapterConfiguration -Filter "DHCPEnabled=True" |Select-Object DNSDomain -ExpandProperty DNSDomain}
Foreach ($result in $24)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $25)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $26)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}   
$Column = 1




$Row++
$Row++
$Hardware.Cells.Item($Row,$Column) = 'MAC-адрес устройства'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Наименование'
$Hardware.Cells.Item($Row,2) = 'Адаптер'
$Hardware.Cells.Item($Row,3) = 'MAC-адрес'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$27 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_NetworkAdapter -Filter "NetConnectionStatus>0"|Select-Object Name -ExpandProperty Name}
$28 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_NetworkAdapter -Filter "NetConnectionStatus>0"|Select-Object AdapterType -ExpandProperty AdapterType}
$29 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_NetworkAdapter -Filter "NetConnectionStatus>0"|Select-Object MACAddress -ExpandProperty MACAddress}
Foreach ($result in $27)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $28)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $29)
{
    $Hardware.Cells.Item($Row,$Column) = $result
    $Row++
}
$Column = 1




$Row++
$Hardware.Cells.Item($Row,$Column) = 'Компьютерная система (ноутбук или собранный)'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.ColorIndex = 55
$Hardware.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Hardware.Cells.Item($Row,1) = 'Производитель'
$Hardware.Cells.Item($Row,2) = 'Модель'
$Hardware.Cells.Item($Row,3) = 'Редакция'
$Hardware.Rows.Item($Row).Font.Bold = $true
$Hardware.Rows.Item($Row).Font.Size = 12
$Row++

$30 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_ComputerSystemProduct|Select-Object Vendor -ExpandProperty Vendor}
$31 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_ComputerSystemProduct|Select-Object Version -ExpandProperty Version}
$32 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_ComputerSystemProduct|Select-Object Name -ExpandProperty Name}

$Hardware.Cells.Item($Row, $column) = $30
$column++
$Hardware.Cells.Item($Row, $column) = $31
$column++
$Hardware.Cells.Item($Row, $column) = $32
$Row++
$column = 1

$Hardware.UsedRange.EntireColumn.AutoFit()|Out-Null








$WorkBook.Worksheets.Add()
$Software = $WorkBook.Worksheets.Item(1)
$Software.name = 'Программная часть' 
$Row = 1
$Column = 1


$Software.Cells.Item($Row,$Column) = 'Активная учетная запись'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.ColorIndex = 55
$Software.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$33 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_ComputerSystem|Select-Object Username -ExpandProperty Username}

$Software.Cells.Item($Row, $column) = $33
$Row++
$column = 1




$Row++
$Software.Cells.Item($Row,$Column) = 'Пользователи компьютера'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.ColorIndex = 55
$Software.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Software.Cells.Item($Row,1) = 'Логин'
$Software.Cells.Item($Row,2) = 'Активность'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$34 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Desktop|Select-Object Name -ExpandProperty Name}
$35 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Desktop|Select-Object ScreenSaverActive -ExpandProperty ScreenSaverActive}
Foreach ($result in $34)
{
    $Software.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $35)
{
    $Software.Cells.Item($Row,$Column) = $result
    $Row++
}
$Column = 1




$Row++
$Software.Cells.Item($Row,$Column) = 'Операционная система'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.ColorIndex = 55
$Software.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Software.Cells.Item($Row,1) = 'Сетевая имя'
$Software.Cells.Item($Row,2) = 'Наименование'
$Software.Cells.Item($Row,3) = 'Серийный номер'
$Software.Cells.Item($Row,4) = 'Пользователь'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.Size = 12
$Row++

$36 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_OperatingSystem|Select-Object csname -ExpandProperty csname}
$37 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_OperatingSystem|Select-Object Caption -ExpandProperty Caption}
$38 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_OperatingSystem|Select-Object SerialNumber -ExpandProperty SerialNumber}
$39 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_OperatingSystem|Select-Object RegisteredUser -ExpandProperty RegisteredUser}

$Software.Cells.Item($Row, $column) = $36
$column++
$Software.Cells.Item($Row, $column) = $37
$column++
$Software.Cells.Item($Row, $column) = $38
$column++
$Software.Cells.Item($Row, $column) = $39
$Row++
$column = 1




$Row++
$Software.Cells.Item($Row,$Column) = 'Установленные программы'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.ColorIndex = 55
$Software.Rows.Item($Row).Font.Size = 16
$Row++
$Column = 1

$Software.Cells.Item($Row,1) = 'Название'
$Software.Cells.Item($Row,2) = 'Версия'
$Software.Cells.Item($Row,3) = 'Производитель'
$Software.Rows.Item($Row).Font.Bold = $true
$Software.Rows.Item($Row).Font.Size = 12
$Row++

$BaseRow = $Row
$40 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Product|Select-Object Name -ExpandProperty Name}
$41 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Product|Select-Object Version -ExpandProperty Version}
$42 = Invoke-Command -ComputerName $i -Credential $cred -ScriptBlock {Get-WmiObject -class Win32_Product|Select-Object Vendor -ExpandProperty Vendor}
Foreach ($result in $40)
{
    $Software.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $41)
{
    $Software.Cells.Item($Row,$Column) = $result
    $Row++
}
$Row = $BaseRow
$Column++
Foreach ($result in $42)
{
    $Software.Cells.Item($Row,$Column) = $result
    $Row++
}
$Column = 1

$Software.UsedRange.EntireColumn.AutoFit()|Out-Null




$WorkBook.SaveAs("$path\$a.xlsx")
$Excel.close
$Excel.Quit()
                                                      
}
}
Remove-Item $file                                                           #удаление файла с адресами
Clear-Item wsman:\localhost\client\trustedhosts -Force                      #очистка доверенных хостов
Get-Item WSMan:\localhost\Client\TrustedHosts                               #вывод списка хостов
Write-Host "----------Выполнено----------"









