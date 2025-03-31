Option Explicit

' Déclaration des variables globales
Dim objFSO, objFile, objWMIService, objShell, objNetwork
Dim strComputer, strOutputFile, strHTMLOutputFile, bHTMLReport

' Initialisation des objets
strComputer = "."
strOutputFile = "system_info_complete.log"
strHTMLOutputFile = "system_info_complete.html"
bHTMLReport = True   ' Mettre à False pour n'avoir que le rapport TXT

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strOutputFile, True, True)  ' Unicode support
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set objShell = CreateObject("WScript.Shell")
Set objNetwork = CreateObject("WScript.Network")

' Activation de la gestion d'erreurs
On Error Resume Next

' Variables pour le rapport HTML
Dim objHTMLFile, strHTMLContent
If bHTMLReport Then
    Set objHTMLFile = objFSO.CreateTextFile(strHTMLOutputFile, True, True)
    strHTMLContent = "<html><head><title>Rapport d'information système</title>" & _
                     "<style>" & _
                     "body { font-family: Arial, sans-serif; margin: 20px; }" & _
                     "h1 { color: #2c3e50; text-align: center; }" & _
                     "h2 { color: #3498db; margin-top: 30px; border-bottom: 1px solid #3498db; padding-bottom: 5px; }" & _
                     "table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }" & _
                     "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }" & _
                     "th { background-color: #f2f2f2; }" & _
                     "tr:nth-child(even) { background-color: #f9f9f9; }" & _
                     ".highlight { background-color: #e8f4f8; font-weight: bold; }" & _
                     "</style></head><body>" & _
                     "<h1>RAPPORT D'INFORMATION SYSTÈME</h1>"
End If

' Fonction pour écrire une section
Sub WriteSection(sectionName)
    objFile.WriteLine String(80, "=")
    objFile.WriteLine sectionName
    objFile.WriteLine String(80, "=")
    
    If bHTMLReport Then
        strHTMLContent = strHTMLContent & "<h2>" & sectionName & "</h2>"
    End If
End Sub

' Fonction pour démarrer une table HTML
Sub StartHTMLTable(headers)
    If bHTMLReport Then
        Dim header, headerHTML
        headerHTML = "<table><tr>"
        For Each header In headers
            headerHTML = headerHTML & "<th>" & header & "</th>"
        Next
        headerHTML = headerHTML & "</tr>"
        strHTMLContent = strHTMLContent & headerHTML
    End If
End Sub

' Fonction pour ajouter une ligne à la table HTML
Sub AddHTMLTableRow(values)
    If bHTMLReport Then
        Dim value, rowHTML
        rowHTML = "<tr>"
        For Each value In values
            rowHTML = rowHTML & "<td>" & value & "</td>"
        Next
        rowHTML = rowHTML & "</tr>"
        strHTMLContent = strHTMLContent & rowHTML
    End If
End Sub

' Fonction pour terminer une table HTML
Sub EndHTMLTable()
    If bHTMLReport Then
        strHTMLContent = strHTMLContent & "</table>"
    End If
End Sub

' Fonction pour vérifier et écrire une valeur
Function SafeValue(value)
    If IsNull(value) Or value = "" Then
        SafeValue = "Non disponible"
    Else
        SafeValue = CStr(value)
    End If
End Function

' Fonction pour formater les octets
Function FormatBytes(bytes)
    If IsNull(bytes) Then
        FormatBytes = "Non disponible"
    Else
        If bytes < 1024 Then
            FormatBytes = bytes & " octets"
        ElseIf bytes < 1024^2 Then
            FormatBytes = Round(bytes/1024, 2) & " Ko"
        ElseIf bytes < 1024^3 Then
            FormatBytes = Round(bytes/1024^2, 2) & " Mo"
        Else
            FormatBytes = Round(bytes/1024^3, 2) & " Go"
        End If
    End If
End Function

' Fonction pour vérifier l'état d'activation de Windows
Function GetWindowsActivationStatus()
    Dim objService, productKeyChannel
    
    Set objService = GetObject("winmgmts:\\.\root\cimv2\Security\MicrosoftVolumeActivation")
    
    On Error Resume Next
    For Each productKeyChannel in objService.ExecQuery("Select * from SoftwareLicensingProduct Where ApplicationID = '55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseIsAddon = False")
        If Err.Number = 0 Then
            If productKeyChannel.LicenseStatus = 1 Then
                GetWindowsActivationStatus = "Activé"
            Else
                GetWindowsActivationStatus = "Non activé (Code " & productKeyChannel.LicenseStatus & ")"
            End If
            Exit Function
        End If
    Next
    
    ' Méthode alternative si la première échoue
    Dim oExec, sOutput
    Set oExec = objShell.Exec("cscript //nologo %windir%\system32\slmgr.vbs /dli")
    
    sOutput = oExec.StdOut.ReadAll()
    If InStr(sOutput, "notification may be accurate") > 0 Or InStr(sOutput, "Licensed") > 0 Then
        GetWindowsActivationStatus = "Probablement activé"
    Else
        GetWindowsActivationStatus = "Probablement non activé"
    End If
    
    On Error GoTo 0
End Function

' En-tête
WriteSection("RAPPORT D'INFORMATION SYSTÈME")
objFile.WriteLine "Date du rapport : " & Now()
objFile.WriteLine "Nom de l'ordinateur : " & objNetwork.ComputerName
objFile.WriteLine "Utilisateur : " & objNetwork.UserName
objFile.WriteLine "Domaine : " & objNetwork.UserDomain
objFile.WriteLine

' Pour le rapport HTML
If bHTMLReport Then
    Dim headerData
    headerData = Array("Propriété", "Valeur")
    StartHTMLTable(headerData)
    AddHTMLTableRow(Array("Date du rapport", Now()))
    AddHTMLTableRow(Array("Nom de l'ordinateur", objNetwork.ComputerName))
    AddHTMLTableRow(Array("Utilisateur", objNetwork.UserName))
    AddHTMLTableRow(Array("Domaine", objNetwork.UserDomain))
    EndHTMLTable()
End If

' Système d'exploitation
WriteSection("SYSTÈME D'EXPLOITATION")
Dim objOS, activationStatus
activationStatus = GetWindowsActivationStatus()

If bHTMLReport Then
    StartHTMLTable(Array("Propriété", "Valeur"))
End If

For Each objOS in objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    If Err.Number = 0 Then
        objFile.WriteLine "OS : " & SafeValue(objOS.Caption)
        objFile.WriteLine "Version : " & SafeValue(objOS.Version)
        objFile.WriteLine "Build : " & SafeValue(objOS.BuildNumber)
        objFile.WriteLine "Architecture : " & SafeValue(objOS.OSArchitecture)
        objFile.WriteLine "Numéro de série : " & SafeValue(objOS.SerialNumber)
        objFile.WriteLine "Statut d'activation : " & activationStatus
        objFile.WriteLine "Chemin Windows : " & SafeValue(objOS.WindowsDirectory)
        objFile.WriteLine "Nom du système : " & SafeValue(objOS.CSName)
        objFile.WriteLine "Dernier démarrage : " & SafeValue(objOS.LastBootUpTime)
        objFile.WriteLine "Temps d'activité : " & FormatTimespan(DateDiff("s", CDate(objOS.LastBootUpTime), Now))
        objFile.WriteLine "Mémoire virtuelle totale : " & FormatBytes(objOS.TotalVirtualMemorySize * 1024)
        objFile.WriteLine "Mémoire virtuelle libre : " & FormatBytes(objOS.FreeVirtualMemory * 1024)
        objFile.WriteLine "Mémoire physique totale : " & FormatBytes(objOS.TotalVisibleMemorySize * 1024)
        objFile.WriteLine "Mémoire physique libre : " & FormatBytes(objOS.FreePhysicalMemory * 1024)
        
        If bHTMLReport Then
            AddHTMLTableRow(Array("OS", SafeValue(objOS.Caption)))
            AddHTMLTableRow(Array("Version", SafeValue(objOS.Version)))
            AddHTMLTableRow(Array("Build", SafeValue(objOS.BuildNumber)))
            AddHTMLTableRow(Array("Architecture", SafeValue(objOS.OSArchitecture)))
            AddHTMLTableRow(Array("Numéro de série", SafeValue(objOS.SerialNumber)))
            AddHTMLTableRow(Array("Statut d'activation", activationStatus))
            AddHTMLTableRow(Array("Chemin Windows", SafeValue(objOS.WindowsDirectory)))
            AddHTMLTableRow(Array("Nom du système", SafeValue(objOS.CSName)))
            AddHTMLTableRow(Array("Dernier démarrage", SafeValue(objOS.LastBootUpTime)))
            AddHTMLTableRow(Array("Temps d'activité", FormatTimespan(DateDiff("s", CDate(objOS.LastBootUpTime), Now))))
            AddHTMLTableRow(Array("Mémoire virtuelle totale", FormatBytes(objOS.TotalVirtualMemorySize * 1024)))
            AddHTMLTableRow(Array("Mémoire virtuelle libre", FormatBytes(objOS.FreeVirtualMemory * 1024)))
            AddHTMLTableRow(Array("Mémoire physique totale", FormatBytes(objOS.TotalVisibleMemorySize * 1024)))
            AddHTMLTableRow(Array("Mémoire physique libre", FormatBytes(objOS.FreePhysicalMemory * 1024)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations OS : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations OS : " & Err.Description))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Processeur
WriteSection("PROCESSEUR")
Dim objProcessor

If bHTMLReport Then
    StartHTMLTable(Array("Propriété", "Valeur"))
End If

For Each objProcessor in objWMIService.ExecQuery("Select * from Win32_Processor")
    If Err.Number = 0 Then
        objFile.WriteLine "Nom : " & SafeValue(objProcessor.Name)
        objFile.WriteLine "Fabricant : " & SafeValue(objProcessor.Manufacturer)
        objFile.WriteLine "ID : " & SafeValue(objProcessor.ProcessorId)
        objFile.WriteLine "Fréquence : " & SafeValue(objProcessor.MaxClockSpeed) & " MHz"
        objFile.WriteLine "Fréquence actuelle : " & SafeValue(objProcessor.CurrentClockSpeed) & " MHz"
        objFile.WriteLine "Nombre de cœurs : " & SafeValue(objProcessor.NumberOfCores)
        objFile.WriteLine "Nombre de threads : " & SafeValue(objProcessor.NumberOfLogicalProcessors)
        objFile.WriteLine "Cache L2 : " & FormatBytes(objProcessor.L2CacheSize * 1024)
        objFile.WriteLine "Cache L3 : " & FormatBytes(objProcessor.L3CacheSize * 1024)
        objFile.WriteLine "Socket : " & SafeValue(objProcessor.SocketDesignation)
        objFile.WriteLine "Virtualisation : " & SafeValue(objProcessor.VirtualizationFirmwareEnabled)
        objFile.WriteLine "Charge CPU actuelle : " & GetCPULoad() & "%"
        
        If bHTMLReport Then
            AddHTMLTableRow(Array("Nom", SafeValue(objProcessor.Name)))
            AddHTMLTableRow(Array("Fabricant", SafeValue(objProcessor.Manufacturer)))
            AddHTMLTableRow(Array("ID", SafeValue(objProcessor.ProcessorId)))
            AddHTMLTableRow(Array("Fréquence", SafeValue(objProcessor.MaxClockSpeed) & " MHz"))
            AddHTMLTableRow(Array("Fréquence actuelle", SafeValue(objProcessor.CurrentClockSpeed) & " MHz"))
            AddHTMLTableRow(Array("Nombre de cœurs", SafeValue(objProcessor.NumberOfCores)))
            AddHTMLTableRow(Array("Nombre de threads", SafeValue(objProcessor.NumberOfLogicalProcessors)))
            AddHTMLTableRow(Array("Cache L2", FormatBytes(objProcessor.L2CacheSize * 1024)))
            AddHTMLTableRow(Array("Cache L3", FormatBytes(objProcessor.L3CacheSize * 1024)))
            AddHTMLTableRow(Array("Socket", SafeValue(objProcessor.SocketDesignation)))
            AddHTMLTableRow(Array("Virtualisation", SafeValue(objProcessor.VirtualizationFirmwareEnabled)))
            AddHTMLTableRow(Array("Charge CPU actuelle", GetCPULoad() & "%"))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations CPU : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations CPU : " & Err.Description))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Fonction pour obtenir la charge CPU
Function GetCPULoad()
    Dim objCPULoad, cpuLoad
    cpuLoad = "Non disponible"
    
    For Each objCPULoad in objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Processor WHERE Name='_Total'")
        If Err.Number = 0 Then
            cpuLoad = 100 - objCPULoad.PercentIdleTime
        End If
    Next
    
    GetCPULoad = cpuLoad
End Function

' Fonction pour formater une durée en secondes
Function FormatTimespan(seconds)
    Dim days, hours, minutes
    
    days = Int(seconds / 86400)
    seconds = seconds Mod 86400
    
    hours = Int(seconds / 3600)
    seconds = seconds Mod 3600
    
    minutes = Int(seconds / 60)
    seconds = seconds Mod 60
    
    FormatTimespan = days & " jours, " & hours & " heures, " & minutes & " minutes, " & seconds & " secondes"
End Function

' Carte mère
WriteSection("CARTE MÈRE")
Dim objBoard

If bHTMLReport Then
    StartHTMLTable(Array("Propriété", "Valeur"))
End If

For Each objBoard in objWMIService.ExecQuery("Select * from Win32_BaseBoard")
    If Err.Number = 0 Then
        objFile.WriteLine "Fabricant : " & SafeValue(objBoard.Manufacturer)
        objFile.WriteLine "Modèle : " & SafeValue(objBoard.Product)
        objFile.WriteLine "Numéro de série : " & SafeValue(objBoard.SerialNumber)
        objFile.WriteLine "Version : " & SafeValue(objBoard.Version)
        
        If bHTMLReport Then
            AddHTMLTableRow(Array("Fabricant", SafeValue(objBoard.Manufacturer)))
            AddHTMLTableRow(Array("Modèle", SafeValue(objBoard.Product)))
            AddHTMLTableRow(Array("Numéro de série", SafeValue(objBoard.SerialNumber)))
            AddHTMLTableRow(Array("Version", SafeValue(objBoard.Version)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations carte mère : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations carte mère : " & Err.Description))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' BIOS
WriteSection("BIOS")
Dim objBIOS

If bHTMLReport Then
    StartHTMLTable(Array("Propriété", "Valeur"))
End If

For Each objBIOS in objWMIService.ExecQuery("Select * from Win32_BIOS")
    If Err.Number = 0 Then
        objFile.WriteLine "Version : " & SafeValue(objBIOS.SMBIOSBIOSVersion)
        objFile.WriteLine "Fabricant : " & SafeValue(objBIOS.Manufacturer)
        objFile.WriteLine "Date : " & SafeValue(objBIOS.ReleaseDate)
        objFile.WriteLine "Numéro de série : " & SafeValue(objBIOS.SerialNumber)
        
        If bHTMLReport Then
            AddHTMLTableRow(Array("Version", SafeValue(objBIOS.SMBIOSBIOSVersion)))
            AddHTMLTableRow(Array("Fabricant", SafeValue(objBIOS.Manufacturer)))
            AddHTMLTableRow(Array("Date", SafeValue(objBIOS.ReleaseDate)))
            AddHTMLTableRow(Array("Numéro de série", SafeValue(objBIOS.SerialNumber)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations BIOS : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations BIOS : " & Err.Description))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' RAM
WriteSection("MÉMOIRE RAM")
Dim objRAM, totalRAM

If bHTMLReport Then
    StartHTMLTable(Array("Banc", "Capacité", "Type", "Vitesse", "Fabricant", "Numéro de série"))
End If

totalRAM = 0
For Each objRAM in objWMIService.ExecQuery("Select * from Win32_PhysicalMemory")
    If Err.Number = 0 Then
        totalRAM = totalRAM + objRAM.Capacity
        objFile.WriteLine "--- Module de RAM ---"
        objFile.WriteLine "Capacité : " & FormatBytes(objRAM.Capacity)
        objFile.WriteLine "Type : " & GetRAMType(objRAM.MemoryType)
        objFile.WriteLine "Vitesse : " & SafeValue(objRAM.Speed) & " MHz"
        objFile.WriteLine "Fabricant : " & SafeValue(objRAM.Manufacturer)
        objFile.WriteLine "Numéro de série : " & SafeValue(objRAM.SerialNumber)
        objFile.WriteLine "Banc : " & SafeValue(objRAM.DeviceLocator)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objRAM.DeviceLocator), _
                                 FormatBytes(objRAM.Capacity), _
                                 GetRAMType(objRAM.MemoryType), _
                                 SafeValue(objRAM.Speed) & " MHz", _
                                 SafeValue(objRAM.Manufacturer), _
                                 SafeValue(objRAM.SerialNumber)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations RAM : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations RAM : " & Err.Description, "", "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
    strHTMLContent = strHTMLContent & "<p class='highlight'>Mémoire totale installée : " & FormatBytes(totalRAM) & "</p>"
End If

objFile.WriteLine "Mémoire totale installée : " & FormatBytes(totalRAM)
objFile.WriteLine
Err.Clear

' Fonction pour obtenir le type de RAM en texte
Function GetRAMType(typeCode)
    Select Case typeCode
        Case 0: GetRAMType = "Inconnu"
        Case 1: GetRAMType = "Autre"
        Case 2: GetRAMType = "DRAM"
        Case 3: GetRAMType = "SDRAM Synchrone"
        Case 4: GetRAMType = "Cache DRAM"
        Case 5: GetRAMType = "EDO"
        Case 6: GetRAMType = "EDRAM"
        Case 7: GetRAMType = "VRAM"
        Case 8: GetRAMType = "SRAM"
        Case 9: GetRAMType = "RAM"
        Case 10: GetRAMType = "ROM"
        Case 11: GetRAMType = "Flash"
        Case 12: GetRAMType = "EEPROM"
        Case 13: GetRAMType = "FEPROM"
        Case 14: GetRAMType = "EPROM"
        Case 15: GetRAMType = "CDRAM"
        Case 16: GetRAMType = "3DRAM"
        Case 17: GetRAMType = "SDRAM"
        Case 18: GetRAMType = "SGRAM"
        Case 19: GetRAMType = "RDRAM"
        Case 20: GetRAMType = "DDR"
        Case 21: GetRAMType = "DDR2"
        Case 22: GetRAMType = "DDR2 FB-DIMM"
        Case 24: GetRAMType = "DDR3"
        Case 25: GetRAMType = "FBD2"
        Case 26: GetRAMType = "DDR4"
        Case 27: GetRAMType = "LPDDR"
        Case 28: GetRAMType = "LPDDR2"
        Case 29: GetRAMType = "LPDDR3"
        Case 30: GetRAMType = "LPDDR4"
        Case Else: GetRAMType = "Type " & typeCode
    End Select
End Function

' Disques
WriteSection("DISQUES")
Dim objDisk

If bHTMLReport Then
    StartHTMLTable(Array("Modèle", "Type", "Taille", "Partitions", "Numéro de série", "État"))
End If

For Each objDisk in objWMIService.ExecQuery("Select * from Win32_DiskDrive")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Disque ---"
        objFile.WriteLine "Modèle : " & SafeValue(objDisk.Model)
        objFile.WriteLine "Type : " & SafeValue(objDisk.InterfaceType)
        objFile.WriteLine "Taille : " & FormatBytes(objDisk.Size)
        objFile.WriteLine "Partitions : " & SafeValue(objDisk.Partitions)
        objFile.WriteLine "Numéro de série : " & SafeValue(objDisk.SerialNumber)
        objFile.WriteLine "État : " & SafeValue(objDisk.Status)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objDisk.Model), _
                                 SafeValue(objDisk.InterfaceType), _
                                 FormatBytes(objDisk.Size), _
                                 SafeValue(objDisk.Partitions), _
                                 SafeValue(objDisk.SerialNumber), _
                                 SafeValue(objDisk.Status)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations disque : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations disque : " & Err.Description, "", "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

Err.Clear

' Volumes
WriteSection("VOLUMES")
Dim objVolume

If bHTMLReport Then
    StartHTMLTable(Array("Lettre", "Nom", "Système de fichiers", "Espace total", "Espace libre", "Pourcentage libre"))
End If

For Each objVolume in objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType=3")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Volume ---"
        objFile.WriteLine "Lettre : " & SafeValue(objVolume.DeviceID)
        objFile.WriteLine "Nom : " & SafeValue(objVolume.VolumeName)
        objFile.WriteLine "Système de fichiers : " & SafeValue(objVolume.FileSystem)
        objFile.WriteLine "Espace total : " & FormatBytes(objVolume.Size)
        objFile.WriteLine "Espace libre : " & FormatBytes(objVolume.FreeSpace)
        
        Dim percentFree
        percentFree = "N/A"
        If Not IsNull(objVolume.Size) And objVolume.Size > 0 Then
            percentFree = Round((objVolume.FreeSpace/objVolume.Size)*100, 2) & "%"
            objFile.WriteLine "Pourcentage libre : " & percentFree
        End If
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objVolume.DeviceID), _
                                 SafeValue(objVolume.VolumeName), _
                                 SafeValue(objVolume.FileSystem), _
                                 FormatBytes(objVolume.Size), _
                                 FormatBytes(objVolume.FreeSpace), _
                                 percentFree))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations volume : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations volume : " & Err.Description, "", "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

Err.Clear

' Cartes graphiques
WriteSection("CARTES GRAPHIQUES")
Dim objVideo

If bHTMLReport Then
    StartHTMLTable(Array("Nom", "Processeur", "RAM", "Résolution", "Bits par pixel", "Taux de rafraîchissement"))
End If

For Each objVideo in objWMIService.ExecQuery("Select * from Win32_VideoController")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Carte graphique ---"
        objFile.WriteLine "Nom : " & SafeValue(objVideo.Name)
        objFile.WriteLine "Processeur : " & SafeValue(objVideo.VideoProcessor)
        objFile.WriteLine "RAM : " & FormatBytes(objVideo.AdapterRAM)
        objFile.WriteLine "Résolution : " & SafeValue(objVideo.CurrentHorizontalResolution) & "x" & SafeValue(objVideo.CurrentVerticalResolution)
        objFile.WriteLine "Bits par pixel : " & SafeValue(objVideo.CurrentBitsPerPixel)
        objFile.WriteLine "Taux de rafraîchissement : " & SafeValue(objVideo.CurrentRefreshRate) & " Hz"
        objFile.WriteLine "Driver Version : " & SafeValue(objVideo.DriverVersion)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objVideo.Name), _
                                 SafeValue(objVideo.VideoProcessor), _
                                 FormatBytes(objVideo.AdapterRAM), _
                                 SafeValue(objVideo.CurrentHorizontalResolution) & "x" & SafeValue(objVideo.CurrentVerticalResolution), _
                                 SafeValue(objVideo.CurrentBitsPerPixel), _
                                 SafeValue(objVideo.CurrentRefreshRate) & " Hz"))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations carte graphique : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations carte graphique : " & Err.Description, "", "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

Err.Clear

' Réseau
WriteSection("RÉSEAU")
Dim objAdapter

If bHTMLReport Then
    StartHTMLTable(Array("Description", "Adresse IP", "Adresse MAC", "Passerelle", "Serveurs DNS", "DHCP"))
End If

For Each objAdapter in objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Adaptateur réseau ---"
        objFile.WriteLine "Description : " & SafeValue(objAdapter.Description)
        
        Dim ipAddresses, gateway, dnsServers, dhcpStatus
        
        ipAddresses = "Non disponible"
        If Not IsNull(objAdapter.IPAddress) Then
            ipAddresses = Join(objAdapter.IPAddress, ", ")
        End If
        objFile.WriteLine "Adresse IP : " & ipAddresses
        
        objFile.WriteLine "Adresse MAC : " & SafeValue(objAdapter.MACAddress)
        
        gateway = "Non disponible"
        If Not IsNull(objAdapter.DefaultIPGateway) Then
            gateway = Join(objAdapter.DefaultIPGateway, ", ")
        End If
        objFile.WriteLine "Passerelle : " & gateway
        
dnsServers = "Non disponible"
        If Not IsNull(objAdapter.DNSServerSearchOrder) Then
            dnsServers = Join(objAdapter.DNSServerSearchOrder, ", ")
        End If
        objFile.WriteLine "Serveurs DNS : " & dnsServers
        
        dhcpStatus = "Désactivé"
        If objAdapter.DHCPEnabled Then
            dhcpStatus = "Activé (Serveur : " & SafeValue(objAdapter.DHCPServer) & ")"
        End If
        objFile.WriteLine "DHCP : " & dhcpStatus
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objAdapter.Description), _
                                 ipAddresses, _
                                 SafeValue(objAdapter.MACAddress), _
                                 gateway, _
                                 dnsServers, _
                                 dhcpStatus))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations réseau : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations réseau : " & Err.Description, "", "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Adaptateurs sonores
WriteSection("ADAPTATEURS SONORES")
Dim objSound

If bHTMLReport Then
    StartHTMLTable(Array("Nom", "Fabricant", "Status"))
End If

For Each objSound in objWMIService.ExecQuery("Select * from Win32_SoundDevice")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Adaptateur sonore ---"
        objFile.WriteLine "Nom : " & SafeValue(objSound.Name)
        objFile.WriteLine "Fabricant : " & SafeValue(objSound.Manufacturer)
        objFile.WriteLine "Status : " & SafeValue(objSound.Status)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objSound.Name), _
                                 SafeValue(objSound.Manufacturer), _
                                 SafeValue(objSound.Status)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations audio : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations audio : " & Err.Description, "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Imprimantes
WriteSection("IMPRIMANTES")
Dim objPrinter

If bHTMLReport Then
    StartHTMLTable(Array("Nom", "Port", "Défaut", "Status"))
End If

For Each objPrinter in objWMIService.ExecQuery("Select * from Win32_Printer")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Imprimante ---"
        objFile.WriteLine "Nom : " & SafeValue(objPrinter.Name)
        objFile.WriteLine "Port : " & SafeValue(objPrinter.PortName)
        
        Dim defaultStatus
        defaultStatus = "Non"
        If objPrinter.Default Then
            defaultStatus = "Oui"
        End If
        objFile.WriteLine "Défaut : " & defaultStatus
        objFile.WriteLine "Status : " & SafeValue(objPrinter.Status)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objPrinter.Name), _
                                 SafeValue(objPrinter.PortName), _
                                 defaultStatus, _
                                 SafeValue(objPrinter.Status)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations imprimante : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations imprimante : " & Err.Description, "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Services
WriteSection("SERVICES PRINCIPAUX")
Dim objService, servicesCount

If bHTMLReport Then
    StartHTMLTable(Array("Nom", "Description", "État", "Type de démarrage"))
End If

servicesCount = 0
For Each objService in objWMIService.ExecQuery("Select * from Win32_Service Where (State='Running' AND StartMode='Auto') OR Name='wuauserv' OR Name='MpsSvc' OR Name='WinDefend' OR Name='BITS'")
    If Err.Number = 0 Then
        servicesCount = servicesCount + 1
        objFile.WriteLine "--- Service ---"
        objFile.WriteLine "Nom : " & SafeValue(objService.Name)
        objFile.WriteLine "Description : " & SafeValue(objService.Description)
        objFile.WriteLine "État : " & SafeValue(objService.State)
        objFile.WriteLine "Type de démarrage : " & GetStartupType(objService.StartMode)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objService.Name), _
                                 SafeValue(objService.Description), _
                                 SafeValue(objService.State), _
                                 GetStartupType(objService.StartMode)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations service : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations service : " & Err.Description, "", ""))
        End If
    End If
    
    ' Limiter à 20 services pour éviter un rapport trop long
    If servicesCount >= 20 Then
        objFile.WriteLine "(Liste limitée aux 20 premiers services pour éviter un rapport trop long)"
        If bHTMLReport Then
            AddHTMLTableRow(Array("Note", "(Liste limitée aux 20 premiers services pour éviter un rapport trop long)", "", ""))
        End If
        Exit For
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Fonction pour obtenir le type de démarrage en français
Function GetStartupType(startMode)
    Select Case startMode
        Case "Auto": GetStartupType = "Automatique"
        Case "Manual": GetStartupType = "Manuel"
        Case "Disabled": GetStartupType = "Désactivé"
        Case Else: GetStartupType = startMode
    End Select
End Function

' Applications installées
WriteSection("APPLICATIONS INSTALLÉES")
Dim objApp, appsCount

If bHTMLReport Then
    StartHTMLTable(Array("Nom", "Éditeur", "Version", "Date d'installation"))
End If

appsCount = 0
For Each objApp in objWMIService.ExecQuery("Select * from Win32_Product")
    If Err.Number = 0 Then
        appsCount = appsCount + 1
        objFile.WriteLine "--- Application ---"
        objFile.WriteLine "Nom : " & SafeValue(objApp.Name)
        objFile.WriteLine "Éditeur : " & SafeValue(objApp.Vendor)
        objFile.WriteLine "Version : " & SafeValue(objApp.Version)
        objFile.WriteLine "Date d'installation : " & SafeValue(objApp.InstallDate)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objApp.Name), _
                                 SafeValue(objApp.Vendor), _
                                 SafeValue(objApp.Version), _
                                 SafeValue(objApp.InstallDate)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations applications : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations applications : " & Err.Description, "", ""))
        End If
    End If
    
    ' Limiter à 30 applications pour éviter un rapport trop long
    If appsCount >= 30 Then
        objFile.WriteLine "(Liste limitée aux 30 premières applications pour éviter un rapport trop long)"
        If bHTMLReport Then
            AddHTMLTableRow(Array("Note", "(Liste limitée aux 30 premières applications pour éviter un rapport trop long)", "", ""))
        End If
        Exit For
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Mises à jour Windows
WriteSection("MISES À JOUR WINDOWS")
Dim objUpdate, objUpdateSession, objUpdateSearcher, searchResult, updates, updatesCount

If bHTMLReport Then
    StartHTMLTable(Array("ID", "Titre", "Date d'installation"))
End If

On Error Resume Next
Set objUpdateSession = CreateObject("Microsoft.Update.Session")
If Err.Number = 0 Then
    Set objUpdateSearcher = objUpdateSession.CreateUpdateSearcher()
    objUpdateSearcher.Online = False
    Set searchResult = objUpdateSearcher.Search("IsInstalled=1")
    updates = searchResult.Updates
    
    If Err.Number = 0 Then
        updatesCount = 0
        For Each objUpdate in updates
            updatesCount = updatesCount + 1
            objFile.WriteLine "--- Mise à jour ---"
            objFile.WriteLine "ID : " & SafeValue(objUpdate.Identity.UpdateID)
            objFile.WriteLine "Titre : " & SafeValue(objUpdate.Title)
            
            ' Récupérer la date d'installation si disponible
            Dim updateDate
            updateDate = "Non disponible"
            If objUpdate.InstallationBehavior.CanRequestUserInput = False Then
                updateDate = GetUpdateInstallDate(objUpdate.Identity.UpdateID)
            End If
            
            objFile.WriteLine "Date d'installation : " & updateDate
            objFile.WriteLine
            
            If bHTMLReport Then
                AddHTMLTableRow(Array(SafeValue(objUpdate.Identity.UpdateID), _
                                     SafeValue(objUpdate.Title), _
                                     updateDate))
            End If
            
            ' Limiter à 20 mises à jour pour éviter un rapport trop long
            If updatesCount >= 20 Then
                objFile.WriteLine "(Liste limitée aux 20 premières mises à jour pour éviter un rapport trop long)"
                If bHTMLReport Then
                    AddHTMLTableRow(Array("Note", "(Liste limitée aux 20 premières mises à jour pour éviter un rapport trop long)", ""))
                End If
                Exit For
            End If
        Next
    Else
        objFile.WriteLine "Erreur lors de la récupération des mises à jour : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des mises à jour : " & Err.Description, ""))
        End If
    End If
Else
    objFile.WriteLine "Erreur lors de la création de la session de mise à jour : " & Err.Description
    If bHTMLReport Then
        AddHTMLTableRow(Array("Erreur", "Erreur lors de la création de la session de mise à jour : " & Err.Description, ""))
    End If
End If

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Fonction pour obtenir la date d'installation d'une mise à jour
Function GetUpdateInstallDate(updateID)
    Dim objRegistry, strKeyPath, installedOn
    
    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
    strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
    
    ' Cette méthode est approximative car les mises à jour ne sont pas toutes enregistrées de la même façon
    ' Une alternative serait d'utiliser la commande WMIC
    installedOn = "Non disponible"
    
    ' Alternative avec commande WMIC
    Dim oExec, sResult
    
    Set oExec = objShell.Exec("wmic qfe where HotFixID='" & updateID & "' get InstalledOn /value")
    sResult = oExec.StdOut.ReadAll()
    
    If InStr(sResult, "InstalledOn=") > 0 Then
        installedOn = Mid(sResult, InStr(sResult, "InstalledOn=") + 12)
        installedOn = Trim(installedOn)
    End If
    
    GetUpdateInstallDate = installedOn
End Function

' Informations de sécurité
WriteSection("INFORMATIONS DE SÉCURITÉ")

If bHTMLReport Then
    StartHTMLTable(Array("Élément", "Statut"))
End If

' Vérifier l'état de Windows Defender
Dim defenderStatus
defenderStatus = CheckDefenderStatus()
objFile.WriteLine "Windows Defender : " & defenderStatus
If bHTMLReport Then
    AddHTMLTableRow(Array("Windows Defender", defenderStatus))
End If

' Vérifier l'état du pare-feu Windows
Dim firewallStatus
firewallStatus = CheckFirewallStatus()
objFile.WriteLine "Pare-feu Windows : " & firewallStatus
If bHTMLReport Then
    AddHTMLTableRow(Array("Pare-feu Windows", firewallStatus))
End If

' Vérifier les mises à jour automatiques
Dim autoUpdatesStatus
autoUpdatesStatus = CheckAutoUpdateStatus()
objFile.WriteLine "Mises à jour automatiques : " & autoUpdatesStatus
If bHTMLReport Then
    AddHTMLTableRow(Array("Mises à jour automatiques", autoUpdatesStatus))
End If

' Vérifier l'UAC
Dim uacStatus
uacStatus = CheckUACStatus()
objFile.WriteLine "Contrôle des comptes d'utilisateurs (UAC) : " & uacStatus

If bHTMLReport Then
    AddHTMLTableRow(Array("Contrôle des comptes d'utilisateurs (UAC)", uacStatus))
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Vérifier l'état de Windows Defender
Function CheckDefenderStatus()
    Dim defenderKey, defenderValue, defenderStatus
    defenderStatus = "Impossible de déterminer"
    
    ' Méthode 1: Vérifier via WMI
    Dim objDefender
    For Each objDefender in objWMIService.ExecQuery("Select * from Win32_Service Where Name='WinDefend'")
        If objDefender.State = "Running" Then
            defenderStatus = "Actif"
        Else
            defenderStatus = "Désactivé"
        End If
        Exit For
    Next
    
    ' Méthode 2: Vérifier via la commande PowerShell
    If defenderStatus = "Impossible de déterminer" Then
        Dim oExec, sResult
        Set oExec = objShell.Exec("powershell -command ""Get-MpComputerStatus | Select -ExpandProperty AntivirusEnabled""")
        sResult = oExec.StdOut.ReadAll()
        
        If InStr(sResult, "True") > 0 Then
            defenderStatus = "Actif"
        ElseIf InStr(sResult, "False") > 0 Then
            defenderStatus = "Désactivé"
        End If
    End If
    
    CheckDefenderStatus = defenderStatus
End Function

' Vérifier l'état du pare-feu Windows
Function CheckFirewallStatus()
    Dim firewallStatus
    firewallStatus = "Impossible de déterminer"
    
    ' Méthode 1: Vérifier via WMI
    Dim objFirewall
    For Each objFirewall in objWMIService.ExecQuery("Select * from Win32_Service Where Name='MpsSvc'")
        If objFirewall.State = "Running" Then
            firewallStatus = "Actif"
        Else
            firewallStatus = "Désactivé"
        End If
        Exit For
    Next
    
    ' Méthode 2: Vérifier via la commande PowerShell
    If firewallStatus = "Impossible de déterminer" Then
        Dim oExec, sResult
        Set oExec = objShell.Exec("powershell -command ""Get-NetFirewallProfile -Profile Domain,Public,Private | Select -ExpandProperty Enabled""")
        sResult = oExec.StdOut.ReadAll()
        
        If InStr(sResult, "True") > 0 And InStr(sResult, "False") = 0 Then
            firewallStatus = "Actif sur tous les profils"
        ElseIf InStr(sResult, "True") > 0 Then
            firewallStatus = "Actif sur certains profils"
        ElseIf InStr(sResult, "False") > 0 And InStr(sResult, "True") = 0 Then
            firewallStatus = "Désactivé sur tous les profils"
        End If
    End If
    
    CheckFirewallStatus = firewallStatus
End Function

' Vérifier l'état des mises à jour automatiques
Function CheckAutoUpdateStatus()
    Dim autoUpdateStatus
    autoUpdateStatus = "Impossible de déterminer"
    
    ' Méthode 1: Vérifier via la clé de registre
    Dim objRegistry, value
    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
    objRegistry.GetDWORDValue &H80000002, "SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU", "NoAutoUpdate", value
    
    If IsNull(value) Then
        ' Clé de stratégie de groupe non définie, vérifier la clé normale
        objRegistry.GetDWORDValue &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "AUOptions", value
        
        If Not IsNull(value) Then
            Select Case value
                Case 1: autoUpdateStatus = "Désactivé"
                Case 2: autoUpdateStatus = "Notification avant téléchargement"
                Case 3: autoUpdateStatus = "Notification avant installation"
                Case 4: autoUpdateStatus = "Installation automatique"
                Case 5: autoUpdateStatus = "Installation automatique ou notification"
                Case Else: autoUpdateStatus = "Configuration inconnue (" & value & ")"
            End Select
        End If
    ElseIf value = 1 Then
        autoUpdateStatus = "Désactivé par stratégie"
    Else
        autoUpdateStatus = "Probablement activé"
    End If
    
    ' Méthode 2: Vérifier si le service est actif
    If autoUpdateStatus = "Impossible de déterminer" Then
        Dim objService
        For Each objService in objWMIService.ExecQuery("Select * from Win32_Service Where Name='wuauserv'")
            If objService.State = "Running" And objService.StartMode = "Auto" Then
                autoUpdateStatus = "Actif"
            ElseIf objService.State = "Running" Then
                autoUpdateStatus = "Actif (manuel)"
            ElseIf objService.StartMode = "Disabled" Then
                autoUpdateStatus = "Désactivé"
            Else
                autoUpdateStatus = "Inactif"
            End If
            Exit For
        Next
    End If
    
    CheckAutoUpdateStatus = autoUpdateStatus
End Function

' Vérifier l'état de l'UAC
Function CheckUACStatus()
    Dim uacStatus, objRegistry, consentPromptBehaviorAdmin, enableLUA
    
    uacStatus = "Impossible de déterminer"
    Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
    
    objRegistry.GetDWORDValue &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "ConsentPromptBehaviorAdmin", consentPromptBehaviorAdmin
    objRegistry.GetDWORDValue &H80000002, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", enableLUA
    
    If Not IsNull(enableLUA) Then
        If enableLUA = 0 Then
            uacStatus = "Désactivé"
        ElseIf Not IsNull(consentPromptBehaviorAdmin) Then
            Select Case consentPromptBehaviorAdmin
                Case 0: uacStatus = "Actif (Ne jamais notifier)"
                Case 1: uacStatus = "Actif (Notification sans assombrissement)"
                Case 2: uacStatus = "Actif (Notification avec assombrissement)"
                Case 3: uacStatus = "Actif (Demander les informations d'identification)"
                Case 4: uacStatus = "Actif (Notification avec assombrissement - Sécurité maximale)"
                Case 5: uacStatus = "Actif (Demander les informations d'identification - Sécurité maximale)"
                Case Else: uacStatus = "Actif (Configuration inconnue)"
            End Select
        Else
            uacStatus = "Actif"
        End If
    End If
    
    CheckUACStatus = uacStatus
End Function

' Informations utilisateur
WriteSection("UTILISATEURS")
Dim objUser

If bHTMLReport Then
    StartHTMLTable(Array("Nom", "Nom complet", "Description", "Statut", "Désactivé", "Type de compte"))
End If

For Each objUser in objWMIService.ExecQuery("Select * from Win32_UserAccount Where LocalAccount = True")
    If Err.Number = 0 Then
        objFile.WriteLine "--- Utilisateur ---"
        objFile.WriteLine "Nom : " & SafeValue(objUser.Name)
        objFile.WriteLine "Nom complet : " & SafeValue(objUser.FullName)
        objFile.WriteLine "Description : " & SafeValue(objUser.Description)
        objFile.WriteLine "Statut : " & IIf(objUser.Status = "OK", "OK", SafeValue(objUser.Status))
        objFile.WriteLine "Désactivé : " & IIf(objUser.Disabled, "Oui", "Non")
        objFile.WriteLine "Type de compte : " & GetAccountType(objUser.AccountType)
        objFile.WriteLine
        
        If bHTMLReport Then
            AddHTMLTableRow(Array(SafeValue(objUser.Name), _
                                 SafeValue(objUser.FullName), _
                                 SafeValue(objUser.Description), _
                                 IIf(objUser.Status = "OK", "OK", SafeValue(objUser.Status)), _
                                 IIf(objUser.Disabled, "Oui", "Non"), _
                                 GetAccountType(objUser.AccountType)))
        End If
    Else
        objFile.WriteLine "Erreur lors de la récupération des informations utilisateur : " & Err.Description
        If bHTMLReport Then
            AddHTMLTableRow(Array("Erreur", "Erreur lors de la récupération des informations utilisateur : " & Err.Description, "", "", "", ""))
        End If
    End If
Next

If bHTMLReport Then
    EndHTMLTable()
End If

objFile.WriteLine
Err.Clear

' Fonction pour obtenir le type de compte en texte
Function GetAccountType(typeCode)
    Select Case typeCode
        Case 256: GetAccountType = "Utilisateur normal"
        Case 512: GetAccountType = "Administrateur"
        Case Else: GetAccountType = "Type " & typeCode
    End Select
End Function

' Fonction IIf alternative pour VBScript
Function IIf(condition, trueValue, falseValue)
    If condition Then
        IIf = trueValue
    Else
        IIf = falseValue
    End If
End Function

' Résumé
WriteSection("RÉSUMÉ ET RECOMMANDATIONS")
Dim recommendations(), recCount
recCount = 0
ReDim recommendations(20)

' Vérifier l'espace disque
Dim lowDiskSpaceDetected
lowDiskSpaceDetected = False
For Each objVolume in objWMIService.ExecQuery("Select * from Win32_LogicalDisk Where DriveType=3")
    If Not IsNull(objVolume.Size) And objVolume.Size > 0 Then
        If (objVolume.FreeSpace / objVolume.Size) < 0.1 Then
            lowDiskSpaceDetected = True
            recCount = recCount + 1
            recommendations(recCount) = "Le volume " & objVolume.DeviceID & " a moins de 10% d'espace libre. Nettoyage recommandé."
        End If
    End If
Next

' Vérifier l'état de Windows Defender
If InStr(defenderStatus, "Désactivé") > 0 Then
    recCount = recCount + 1
    recommendations(recCount) = "Windows Defender est désactivé. Activation recommandée."
End If

' Vérifier l'état du pare-feu Windows
If InStr(firewallStatus, "Désactivé") > 0 Then
    recCount = recCount + 1
    recommendations(recCount) = "Le pare-feu Windows est désactivé. Activation recommandée."
End If

' Vérifier l'état des mises à jour automatiques
If InStr(autoUpdatesStatus, "Désactivé") > 0 Then
    recCount = recCount + 1
    recommendations(recCount) = "Les mises à jour automatiques sont désactivées. Activation recommandée."
End If

' Vérifier l'état de l'UAC
If InStr(uacStatus, "Désactivé") > 0 Then
    recCount = recCount + 1
    recommendations(recCount) = "Le contrôle des comptes d'utilisateurs (UAC) est désactivé. Activation recommandée."
End If

' Vérifier la mémoire système
Dim memoryWarning
memoryWarning = False
For Each objOS in objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    If Not IsNull(objOS.FreePhysicalMemory) And Not IsNull(objOS.TotalVisibleMemorySize) Then
        If (objOS.FreePhysicalMemory / objOS.TotalVisibleMemorySize) < 0.1 Then
            memoryWarning = True
            recCount = recCount + 1
            recommendations(recCount) = "Moins de 10% de mémoire physique disponible. Vérifier les applications en cours d'exécution."
        End If
    End If
Next

' Écrire les recommandations
objFile.WriteLine "Recommandations :"
If recCount > 0 Then
    For i = 1 To recCount
        objFile.WriteLine "- " & recommendations(i)
    Next
Else
    objFile.WriteLine "- Aucune recommandation particulière. Le système semble en bon état."
End If

If bHTMLReport Then
    strHTMLContent = strHTMLContent & "<h2>RÉSUMÉ ET RECOMMANDATIONS</h2>"
    
    If recCount > 0 Then
        strHTMLContent = strHTMLContent & "<ul>"
        For i = 1 To recCount
            strHTMLContent = strHTMLContent & "<li>" & recommendations(i) & "</li>"
        Next
        strHTMLContent = strHTMLContent & "</ul>"
    Else
        strHTMLContent = strHTMLContent & "<p>Aucune recommandation particulière. Le système semble en bon état.</p>"
    End If
    
    ' Fermer le document HTML
    strHTMLContent = strHTMLContent & "</body></html>"
    objHTMLFile.WriteLine strHTMLContent
    objHTMLFile.Close
End If

' Conclusion
objFile.WriteLine
objFile.WriteLine String(80, "=")
objFile.WriteLine "FIN DU RAPPORT"
objFile.WriteLine "Généré le " & Now()
objFile.WriteLine "Pour plus d'informations, contactez votre administrateur système."
objFile.WriteLine String(80, "=")

' Fermer le fichier
objFile.Close

' Message final
WScript.Echo "Rapport d'information système généré avec succès !" & vbCrLf & _
             "Fichier texte : " & strOutputFile & vbCrLf & _
             IIf(bHTMLReport, "Fichier HTML : " & strHTMLOutputFile, "")

' Libérer les objets
Set objFSO = Nothing
Set objFile = Nothing
Set objHTMLFile = Nothing
Set objWMIService = Nothing
Set objShell = Nothing
Set objNetwork = Nothing