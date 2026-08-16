# 1. Принудительно закрываем старые зависшие графические окна в этой сессии
if ($Form) { $Form.Close(); $Form.Dispose() }

# 2. Подключаем базовые графические библиотеки Windows
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# ====================================================================
# ЧАСТЬ 1: СОЗДАНИЕ ПОЛНОГО ИНТЕРФЕЙСА FILEFINDER (ВСЕ КНОПКИ В СБОРЕ)
# ====================================================================

$Form = New-Object System.Windows.Forms.Form
$Form.Text = "FileFinder Explorer"
$Form.Size = New-Object System.Drawing.Size(660, 560) 
$Form.StartPosition = "CenterScreen"
$Form.FormBorderStyle = "FixedSingle"
$Form.MaximizeBox = $false

# Выбор диска
$LblDrive = New-Object System.Windows.Forms.Label
$LblDrive.Text = "Диск:"
$LblDrive.Location = New-Object System.Drawing.Point(20, 20)
$LblDrive.Size = New-Object System.Drawing.Size(45, 20)
$Form.Controls.Add($LblDrive)

$ComboDrive = New-Object System.Windows.Forms.ComboBox
$ComboDrive.Location = New-Object System.Drawing.Point(70, 17)
$ComboDrive.Size = New-Object System.Drawing.Size(70, 20)
$ComboDrive.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
foreach ($Drive in [System.IO.DriveInfo]::GetDrives()) {
    if ($Drive.IsReady) { [void]$ComboDrive.Items.Add($Drive.Name) }
}
if ($ComboDrive.Items.Count -gt 0) { $ComboDrive.SelectedIndex = 0 }
$Form.Controls.Add($ComboDrive)

# Что искать
$LblQuery = New-Object System.Windows.Forms.Label
$LblQuery.Text = "Что искать:"
$LblQuery.Location = New-Object System.Drawing.Point(155, 20)
$LblQuery.Size = New-Object System.Drawing.Size(80, 20)
$Form.Controls.Add($LblQuery)

$TxtQuery = New-Object System.Windows.Forms.TextBox
$TxtQuery.Location = New-Object System.Drawing.Point(240, 17)
$TxtQuery.Size = New-Object System.Drawing.Size(370, 20)
$TxtQuery.Text = "" 
$Form.Controls.Add($TxtQuery)

# Примерная папка
$LblSubPath = New-Object System.Windows.Forms.Label
$LblSubPath.Text = "Выбранный путь к папке:"
$LblSubPath.Location = New-Object System.Drawing.Point(20, 50)
$LblSubPath.Size = New-Object System.Drawing.Size(590, 15)
$Form.Controls.Add($LblSubPath)

$TxtSubPath = New-Object System.Windows.Forms.TextBox
$TxtSubPath.Location = New-Object System.Drawing.Point(20, 70)
$TxtSubPath.Size = New-Object System.Drawing.Size(590, 20)
$TxtSubPath.Text = "" 
$Form.Controls.Add($TxtSubPath)

# Интерактивное дерево TreeView
$TreeView = New-Object System.Windows.Forms.TreeView
$TreeView.Location = New-Object System.Drawing.Point(20, 105)
$TreeView.Size = New-Object System.Drawing.Size(590, 274)
$TreeView.CheckBoxes = $true 
$Form.Controls.Add($TreeView)

# Зеленая шкала процесса (нужна для режима глубокого сканирования)
$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(20, 395)
$ProgressBar.Size = New-Object System.Drawing.Size(590, 18)
$ProgressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
$Form.Controls.Add($ProgressBar)

# Метка статуса
$LblStatus = New-Object System.Windows.Forms.Label
$LblStatus.Text = "Введите запрос и выберите режим..."
$LblStatus.ForeColor = "Gray"
$LblStatus.Location = New-Object System.Drawing.Point(20, 422)
$LblStatus.Size = New-Object System.Drawing.Size(145, 35) 
$Form.Controls.Add($LblStatus)

# Кнопка «Открыть папку...» (Мгновенный проводник по центру)
$BtnOpenFolder = New-Object System.Windows.Forms.Button
$BtnOpenFolder.Text = "Открыть папку..."
$BtnOpenFolder.Location = New-Object System.Drawing.Point(175, 420) 
$BtnOpenFolder.Size = New-Object System.Drawing.Size(240, 35) 
$Form.Controls.Add($BtnOpenFolder)

# КНОПКА ВЕРНУЛАСЬ: «Начать поиск» (Глубокий сканер подсетей справа)
$BtnSearch = New-Object System.Windows.Forms.Button
$BtnSearch.Text = "Начать поиск"
$BtnSearch.Location = New-Object System.Drawing.Point(430, 420)
$BtnSearch.Size = New-Object System.Drawing.Size(180, 35)
$Form.Controls.Add($BtnSearch)

# Нижняя панель действий
$BtnW = 140; $BtnH = 35; $YPos = 470  

$BtnDelete = New-Object System.Windows.Forms.Button
$BtnDelete.Text = "Удалить"
$BtnDelete.Location = New-Object System.Drawing.Point(20, $YPos)
$BtnDelete.Size = New-Object System.Drawing.Size($BtnW, $BtnH)
$Form.Controls.Add($BtnDelete)

$BtnRename = New-Object System.Windows.Forms.Button
$BtnRename.Text = "Переименовать"
$BtnRename.Location = New-Object System.Drawing.Point(170, $YPos)
$BtnRename.Size = New-Object System.Drawing.Size($BtnW, $BtnH)
$Form.Controls.Add($BtnRename)

$BtnExplorer = New-Object System.Windows.Forms.Button
$BtnExplorer.Text = "Открыть в проводнике"
$BtnExplorer.Location = New-Object System.Drawing.Point(320, $YPos)
$BtnExplorer.Size = New-Object System.Drawing.Size($BtnW, $BtnH)
$Form.Controls.Add($BtnExplorer)

$BtnFormat = New-Object System.Windows.Forms.Button
$BtnFormat.Text = "Изменить формат"
$BtnFormat.Location = New-Object System.Drawing.Point(470, $YPos)
$BtnFormat.Size = New-Object System.Drawing.Size($BtnW, $BtnH)
$Form.Controls.Add($BtnFormat)
# ====================================================================
# ЧАСТЬ 2.1: СИСТЕМНЫЙ ДВИЖОК (ПРОСМОТР ПАПОК + ГЛУБОКИЙ ПОИСК ПО ИМЕНИ)
# ====================================================================

$MainTimer = New-Object System.Windows.Forms.Timer
$MainTimer.Interval = 5

$script:IsSearching = $false
$script:FoundCount = 0
$script:CheckedCount = 0
$script:FolderQueue = New-Object System.Collections.Queue
$script:NodeCache = @{}

$script:LastQuery = ""
$script:LastSubPath = ""
$script:LastDrive = ""

# Функция для построения дерева при глубоком поиске по ключевому слову
function Add-ToTree ($FullPath, $IsFolder) {
    $PathParts = $FullPath.Split('\')
    $CurrentNodes = $TreeView.Nodes
    $BuildPath = ""

    for ($i = 0; $i -lt $PathParts.Count; $i++) {
        $Part = $PathParts[$i]
        if ([string]::IsNullOrWhiteSpace($Part)) { continue }

        if ($i -eq 0) { $BuildPath = $Part }
        else { $BuildPath = [System.IO.Path]::Combine($BuildPath, $Part) }

        if ($script:NodeCache.ContainsKey($BuildPath.ToLower())) {
            $TargetNode = $script:NodeCache[$BuildPath.ToLower()]
            $CurrentNodes = $TargetNode.Nodes
        } else {
            if ($i -eq ($PathParts.Count - 1) -and -not $IsFolder) {
                $NewNode = $CurrentNodes.Add($BuildPath, $Part)
                $script:NodeCache[$BuildPath.ToLower()] = $NewNode
            } else {
                $NewNode = $CurrentNodes.Add($BuildPath, "[ПАПКА] $Part")
                $NewNode.Expand() # Разворачиваем найденные ветки на лету
                $script:NodeCache[$BuildPath.ToLower()] = $NewNode
                $CurrentNodes = $NewNode.Nodes
            }
        }
    }
}

# Функция для обычного мгновенного просмотра папки (на один уровень вглубь)
function Populate-Node ($ParentNode, $DirPath) {
    $ParentNode.Nodes.Clear()
    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($DirPath)
        foreach ($SubDir in $DirInfo.GetDirectories()) {
            if ($SubDir.Name.Contains("System Volume Information") -or $SubDir.Name.Contains('$Recycle.Bin')) { continue }
            $DirNode = $ParentNode.Nodes.Add($SubDir.FullName, "[ПАПКА] $($SubDir.Name)")
            [void]$DirNode.Nodes.Add("dummy") # Заглушка для плюсика
        }
        foreach ($File in $DirInfo.GetFiles()) {
            [void]$ParentNode.Nodes.Add($File.FullName, $File.Name)
        }
    } catch {}
}

# Раскрытие папок по плюсику в режиме Проводника
$TreeView.Add_BeforeExpand({
    param($sender, $e)
    $TargetNode = $e.Node
    if ($TargetNode.Nodes.Count -eq 1 -and $TargetNode.Nodes.Name -eq "dummy") {
        Populate-Node $TargetNode $TargetNode.Name
    }
})

# РЕЖИМ 1: Кнопка «Открыть папку...» (Мгновенный просмотр без сканирования диска)
$BtnOpenFolder.Add_Click({
    $FolderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderDialog.Description = "Выберите папку для мгновенного просмотра структуры:"
    $FolderDialog.ShowNewFolderButton = $false

    if ($FolderDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $FullPath = $FolderDialog.SelectedPath
        $TreeView.Nodes.Clear()
        $script:NodeCache.Clear()
        
        $TxtSubPath.Text = $FullPath
        $LblStatus.ForeColor = "Green"
        $LblStatus.Text = "Папка загружена!"

        $RootNode = $TreeView.Nodes.Add($FullPath, "[ПАПКА] " + [System.IO.Path]::GetFileName($FullPath))
        Populate-Node $RootNode $FullPath
        $RootNode.Expand()
    }
})

# Такт таймера для РЕЖИМА 2 (Глубокий поиск файлов по имени)
$MainTimer.Add_Tick({
    if (-not $script:IsSearching) { $MainTimer.Stop(); return }

    if ($script:FolderQueue.Count -eq 0) {
        $script:IsSearching = $false
        $MainTimer.Stop()
        $ProgressBar.Value = $ProgressBar.Maximum
        $BtnSearch.Text = "Начать поиск"
        $BtnSearch.BackColor = [System.Drawing.SystemColors]::Control
        $TxtSubPath.Enabled = $true
        $BtnOpenFolder.Enabled = $true
        $TreeView.ExpandAll() # Раскрываем всё в конце
        $LblStatus.ForeColor = "Green"
        $LblStatus.Text = "Поиск успешно завершен!`nНайдено объектов: $script:FoundCount"
        return
    }

    $CurrentDir = $script:FolderQueue.Dequeue()
    $script:CheckedCount++

    if ($script:CheckedCount -le $ProgressBar.Maximum) { $ProgressBar.Value = $script:CheckedCount }
    $Percent = [Math]::Round(($script:CheckedCount / $ProgressBar.Maximum) * 100, 1)
    $LblStatus.Text = "Прогресс: $Percent% (Найдено: $script:FoundCount)`nПроверка: $CurrentDir"
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $DirInfo = New-Object System.IO.DirectoryInfo($CurrentDir)
        $Files = $DirInfo.GetFiles()
        foreach ($File in $Files) {
            $script:CheckedCount++
            if ($script:CheckedCount -le $ProgressBar.Maximum) { $ProgressBar.Value = $script:CheckedCount }

            if ($File.Name.ToLower().Contains($script:CleanQuery)) {
                $script:FoundCount++
                Add-ToTree $File.FullName $false
            }
        }

        $SubDirs = $DirInfo.GetDirectories()
        foreach ($SubDir in $SubDirs) {
            if ($SubDir.Name.Contains("System Volume Information") -or $SubDir.Name.Contains('Recycle.Bin')) { continue }
            if ($SubDir.Name.ToLower().Contains($script:CleanQuery)) {
                $script:FoundCount++
                Add-ToTree $SubDir.FullName $true
            }
            $script:FolderQueue.Enqueue($SubDir.FullName)
        }
    } catch {}
})

# РЕЖИМ 2: КНОПКА «НАЧАТЬ ПОИСК» (Глубокое сканирование папок по имени файла)
$BtnSearch.Add_Click({
    if ($script:IsSearching) {
        $script:IsSearching = $false
        $MainTimer.Stop()
        $BtnSearch.Text = "Продолжить поиск"
        $BtnSearch.BackColor = [System.Drawing.SystemColors]::Control
        $LblStatus.Text = "Поиск приостановлен."
        $TxtSubPath.Enabled = $true
        $BtnOpenFolder.Enabled = $true
        return
    }

    if ([string]::IsNullOrWhiteSpace($TxtQuery.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Введите ключевое слово для поиска (например, MCMETA)!", "Внимание", "OK", "Warning")
        return
    }

    $SelectedDrive = $ComboDrive.SelectedItem
    $CurrentQueryText = $TxtQuery.Text.Trim()
    $CurrentSubPathText = $TxtSubPath.Text.Trim()

    $IsNewSearch = ($CurrentQueryText -ne $script:LastQuery) -or `
                   ($CurrentSubPathText -ne $script:LastSubPath) -or `
                   ($SelectedDrive -ne $script:LastDrive)

    if ($IsNewSearch -or $script:FolderQueue.Count -eq 0) {
        $script:LastQuery = $CurrentQueryText
        $script:LastSubPath = $CurrentSubPathText
        $script:LastDrive = $SelectedDrive

        $StartPath = $SelectedDrive
        $UserSubPath = $CurrentSubPathText.Trim(' \')
        if (-not [string]::IsNullOrWhiteSpace($UserSubPath)) {
            $StartPath = [System.IO.Path]::Combine($SelectedDrive, $UserSubPath)
        }

        $BtnSearch.Enabled = $false
        $BtnOpenFolder.Enabled = $false
        $LblStatus.ForeColor = "Blue"
        $LblStatus.Text = "Подсчет количества файлов..."
        $Form.Refresh()

        try {
            $TotalObjects = 0
            foreach ($SF in [System.IO.Directory]::EnumerateFiles($StartPath, "*", [System.IO.SearchOption]::TopDirectoryOnly)) { $TotalObjects++ }
            foreach ($SD in [System.IO.Directory]::EnumerateDirectories($StartPath, "*", [System.IO.SearchOption]::TopDirectoryOnly)) {
                try {
                    $TotalObjects += [System.IO.Directory]::GetFiles($SD).Length
                    $TotalObjects += [System.IO.Directory]::GetDirectories($SD).Length
                } catch {}
            }
            if ($TotalObjects -le 100) { $TotalObjects = 5000 }
        } catch { $TotalObjects = 5000 }

        $ProgressBar.Maximum = $TotalObjects
        $ProgressBar.Value = 0

        $script:FoundCount = 0
        $script:CheckedCount = 0
        $script:FolderQueue.Clear()
        $TreeView.Nodes.Clear()
        $script:NodeCache.Clear()
        
        $script:FolderQueue.Enqueue($StartPath)
        $script:CleanQuery = $CurrentQueryText.ToLower().Replace("*", "")
    }

    $script:IsSearching = $true
    $BtnSearch.Enabled = $true
    $BtnOpenFolder.Enabled = $false
    $BtnSearch.Text = "ОСТАНОВИТЬ"
    $BtnSearch.BackColor = [System.Drawing.Color]::LightCoral
    $TxtSubPath.Enabled = $false
    $MainTimer.Start()
})
# ====================================================================
# ЧАСТЬ 2.2: ЛОГИКА 4-Х КНОПОК УПРАВЛЕНИЯ ДЛЯ ДИНАМИЧЕСКОГО ДЕРЕВА
# ====================================================================

# Рекурсивная функция для сбора всех узлов дерева с галочками
function Get-CheckedNodes ($Nodes, $CheckedList) {
    foreach ($Node in $Nodes) {
        if ($Node.Checked -and $Node.Name -ne "dummy") { [void]$CheckedList.Add($Node) }
        if ($Node.Nodes.Count -gt 0) { Get-CheckedNodes $Node.Nodes $CheckedList }
    }
}

# ЛОГИКА 1: МАССОВОЕ УДАЛЕНИЕ ПО ГАЛОЧКАМ
$BtnDelete.Add_Click({
    $CheckedNodes = New-Object System.Collections.Generic.List[System.Windows.Forms.TreeNode]
    Get-CheckedNodes $TreeView.Nodes $CheckedNodes

    if ($CheckedNodes.Count -eq 0) { 
        [System.Windows.Forms.MessageBox]::Show("Выберите объекты галочками в дереве!", "Инфо", "OK", "Information")
        return 
    }

    $Confirm1 = [System.Windows.Forms.MessageBox]::Show( `
        "Удалить выбранные объекты в количестве: $($CheckedNodes.Count) шт.?", `
        "Удаление", "YesNo", "Warning", "Button2")
        
    if ($Confirm1 -eq [System.Windows.Forms.DialogResult]::Yes) {
        $Confirm2 = [System.Windows.Forms.MessageBox]::Show( `
            "Вы полностью согласны на безвозвратное удаление?", `
            "Подтверждение", "OKCancel", "Stop", "Button2")
            
        if ($Confirm2 -eq [System.Windows.Forms.DialogResult]::OK) {
            foreach ($Node in $CheckedNodes) {
                $Path = $Node.Name
                try {
                    if ($Node.Text.StartsWith("[ПАПКА] ")) {
                        [System.IO.Directory]::Delete($Path, $true)
                    } else { [System.IO.File]::Delete($Path) }
                    $Node.Remove() # Убираем узел из дерева на экране
                } catch { 
                    [System.Windows.Forms.MessageBox]::Show("Ошибка удаления: $Path", "Ошибка", "OK", "Error") 
                }
            }
        }
    }
})

# ЛОГИКА 2: МАССОВОЕ ПЕРЕИМЕНОВАНИЕ С АВТОНУМЕРАЦИЕЙ
$BtnRename.Add_Click({
    $CheckedNodes = New-Object System.Collections.Generic.List[System.Windows.Forms.TreeNode]
    Get-CheckedNodes $TreeView.Nodes $CheckedNodes
    $TotalChecked = $CheckedNodes.Count
    
    if ($TotalChecked -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Отметьте галочками папки или файлы!", "Ничего не выбрано", "OK", "Information")
        return
    }

    $DefaultInputText = "Новое_Имя"
    if ($TotalChecked -eq 1) {
        $DefaultInputText = [System.IO.Path]::GetFileNameWithoutExtension($CheckedNodes.Name)
    }

    $BaseNewName = [Microsoft.VisualBasic.Interaction]::InputBox( `
        "Введите новое базовое имя для выбранных объектов ($TotalChecked шт.):", `
        "Массовое переименование", $DefaultInputText)
    
    if ([string]::IsNullOrWhiteSpace($BaseNewName)) { return }

    for ($i = $TotalChecked - 1; $i -ge 0; $i--) {
        $Node = $CheckedNodes[$i]
        $OldPath = $Node.Name
        
        $IsFolder = $Node.Text.StartsWith("[ПАПКА] ")
        $Extension = [System.IO.Path]::GetExtension($OldPath)
        $ParentDir = [System.IO.Path]::GetDirectoryName($OldPath)
        
        if ($TotalChecked -gt 1) {
            $FileNumber = $i + 1
            $NewFullName = "${BaseNewName}_$FileNumber"
            if (-not $IsFolder) { $NewFullName += $Extension }
        } else {
            $NewFullName = $BaseNewName
            if (-not $IsFolder) { $NewFullName += $Extension }
        }
        
        $NewPath = [System.IO.Path]::Combine($ParentDir, $NewFullName)
        
        try {
            if ($IsFolder) { [System.IO.Directory]::Move($OldPath, $NewPath) }
            else { [System.IO.File]::Move($OldPath, $NewPath) }
            
            # Обновляем узел
            $Node.Name = $NewPath
            if ($IsFolder) { $Node.Text = "[ПАПКА] $NewFullName" }
            else { $Node.Text = $NewFullName }
            $Node.Checked = $true
        } catch {}
    }
})

# ЛОГИКА 3: КНОПКА «ОТКРЫТЬ В ПРОВОДНИКЕ»
$BtnExplorer.Add_Click({
    if (-not $TreeView.SelectedNode) { 
        [System.Windows.Forms.MessageBox]::Show("Выделите нужный элемент дерева мышкой (синим цветом)!", "Инфо", "OK", "Information")
        return 
    }
    $Path = $TreeView.SelectedNode.Name
    if ($Path -eq "dummy") { return }
    try { [System.Diagnostics.Process]::Start("explorer.exe", "/select,`"$Path`"") } catch {}
})

# ЛОГИКА 4: МАССОВОЕ ИЗМЕНЕНИЕ ФОРМАТА ПО ГАЛОЧКАМ
$BtnFormat.Add_Click({
    $CheckedNodes = New-Object System.Collections.Generic.List[System.Windows.Forms.TreeNode]
    Get-CheckedNodes $TreeView.Nodes $CheckedNodes
    $TotalChecked = $CheckedNodes.Count
    
    if ($TotalChecked -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Отметьте галочками файлы для смены формата!", "Ничего не выбрано", "OK", "Information")
        return
    }

    $NewExt = [Microsoft.VisualBasic.Interaction]::InputBox( `
        "Введите новое расширение для выбранных файлов ($TotalChecked шт.):`n(Например: .cfg, .txt или .log)", `
        "Массовая смена формата", ".txt")
    
    if ([string]::IsNullOrWhiteSpace($NewExt)) { return }

    for ($i = $TotalChecked - 1; $i -ge 0; $i--) {
        $Node = $CheckedNodes[$i]
        $OldPath = $Node.Name
        if ($Node.Text.StartsWith("[ПАПКА] ")) { continue }
        
        try {
            $NewPath = [System.IO.Path]::ChangeExtension($OldPath, $NewExt)
            [System.IO.File]::Move($OldPath, $NewPath)
            
            # Обновляем ветку дерева под новый формат
            $Node.Name = $NewPath
            $Node.Text = [System.IO.Path]::GetFileName($NewPath)
            $Node.Checked = $true
        } catch {}
    }
})

# Двойной клик по элементу дерева запускает файл или открывает папку
$TreeView.Add_DoubleClick({
    if ($TreeView.SelectedNode) {
        $Path = $TreeView.SelectedNode.Name
        if ($Path -eq "dummy") { return }
        try {
            if ($TreeView.SelectedNode.Text.StartsWith("[ПАПКА] ")) {
                [System.Diagnostics.Process]::Start("explorer.exe", "`"$Path`"")
            } else { [System.Diagnostics.Process]::Start($Path) }
        } catch {}
    }
})

# Завершение работы
$Form.Add_FormClosing({ $script:IsSearching = $false })
$Form.ShowDialog() | Out-Null
