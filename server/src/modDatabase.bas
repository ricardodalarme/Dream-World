Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

' For clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
    Dim filename As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
    Print #1, "The following error occured at ' " & procName & "' in ' " & contName & "' ."
    Print #1, "Run-time error ' " & erNumber & "' : " & erDesc & "."
    Print #1, vbNullString
    Close #1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleError", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LCase$(Dir$(tDir & tName, vbDirectory)) <> tName Then Call MkDir$(tDir & tName)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChkDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddLog", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call WritePrivateProfileString(Header, Var, Value, File)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return function value
    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If
    Else
        If LenB(Dir$(filename)) > 0 Then
            FileExist = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FileExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SaveOptions()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", Str$(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Debug", Str$(Options.Debug)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    Options.Debug = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Debug")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next i

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & GetPlayerName(BannedByIndex)
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BanIndex", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ServerBanIndex(ByVal BanPlayerIndex As Long)
    Dim filename As String
    Dim IP As String
    Dim F As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If

    Next i

    IP = Mid$(IP, 1, i)
    F = FreeFile
    Open filename For Append As #F
    Print #F, IP & "," & "Server"
    Close #F
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & " by " & "the Server" & "!", White)
    Call AddLog("The Server" & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & "The Server" & "!")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ServerBanIndex", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' File path
    filename = "data\accounts\" & Trim$(Name) & ".bin"

    ' Return function value
    If FileExist(filename) Then
        AccountExist = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "AccountExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return function value
    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, ACCOUNT_LENGTH, RightPassword
        Close #nFileNum

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "PasswordOK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ClearPlayer(index)
    Call SetPlayerLogin(index, Name)
    Call SetPlayerPassword(index, Password)
    Call SavePlayer(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddAccount", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call FileCopy(App.Path & "\data\accounts\charlist.txt", App.Path & "\data\accounts\chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DeleteName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Return function value
    If LenB(GetPlayerName(index)) > 0 Then
        CharExist = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CharExist", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim F As Long
    Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If LenB(Trim$(Player(index).Char(TempPlayer(index).Char).Name)) = 0 Then
        Player(index).Char(TempPlayer(index).Char).Name = Name
        Player(index).Char(TempPlayer(index).Char).Sex = Sex
        Player(index).Char(TempPlayer(index).Char).Class = ClassNum
        
        If Player(index).Char(TempPlayer(index).Char).Sex = SEX_MALE Then
            Player(index).Char(TempPlayer(index).Char).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Char(TempPlayer(index).Char).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Char(TempPlayer(index).Char).Level = 1
        
        For n = 1 To Stats.Stat_Count - 1
            Player(index).Char(TempPlayer(index).Char).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).Char(TempPlayer(index).Char).Dir = DIR_DOWN
        Player(index).Char(TempPlayer(index).Char).Map = Class(ClassNum).StartMap
        Player(index).Char(TempPlayer(index).Char).x = Class(ClassNum).StartX
        Player(index).Char(TempPlayer(index).Char).y = Class(ClassNum).StartY
        Player(index).Char(TempPlayer(index).Char).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Char(TempPlayer(index).Char).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).StartItemCount > 0 Then
            For n = 1 To Class(ClassNum).StartItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Char(TempPlayer(index).Char).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Char(TempPlayer(index).Char).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start Spells
        If Class(ClassNum).StarSpellCount > 0 Then
            For n = 1 To Class(ClassNum).StarSpellCount
                If Class(ClassNum).StarSpell(n) > 0 Then
                    ' Spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StarSpell(n)).Name)) > 0 Then
                        Player(index).Char(TempPlayer(index).Char).Spell(n) = Class(ClassNum).StarSpell(n)
                    End If
                End If
            Next
        End If
        
        ' set start titles
        If Class(ClassNum).StartTitleCount > 0 Then
            For n = 1 To Class(ClassNum).StartTitleCount
                If Class(ClassNum).StartTitle(n) > 0 Then
                    ' title exist?
                    If Len(Trim$(Item(Class(ClassNum).StartTitle(n)).Name)) > 0 Then
                        Player(index).Char(TempPlayer(index).Char).Title.Title(n) = Class(ClassNum).StartTitle(n)
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(index)
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddChar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    F = FreeFile
    Open App.Path & "\data\accounts\charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindChar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveAllPlayersOnline", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SavePlayer(ByVal index As Long)
    Dim filename As String
    Dim F As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\accounts\" & Trim$(Player(index).Login) & ".bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Player(index)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SavePlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ClearPlayer(index)
    filename = App.Path & "\data\accounts\" & Trim$(Name) & ".bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Player(index)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeTempPlayer(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim TempPlayer(index).DoT(MAX_DOTS)
    ReDim TempPlayer(index).HoT(MAX_DOTS)
    ReDim TempPlayer(index).TradeOffer(MAX_INV)
    ReDim TempPlayer(index).SpellCD(MAX_PLAYER_SPELLS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeTempPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizePlayer(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Player(index).Char(MAX_PLAYER_CHARS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizePlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeChar(ByVal index As Long, ByVal CharNum As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Player(index).Char(CharNum).Stat(Stats.Stat_Count - 1)
    ReDim Player(index).Char(CharNum).Vital(Vitals.Vital_Count - 1)
    ReDim Player(index).Char(CharNum).Spell(MAX_PLAYER_SPELLS)
    ReDim Player(index).Char(CharNum).Hotbar(MAX_HOTBAR)
    ReDim Player(index).Char(CharNum).Inv(MAX_INV)
    ReDim Player(index).Char(CharNum).Equipment(Equipment.Equipment_Count - 1)
    ReDim Player(index).Char(CharNum).Title.Title(MAX_PLAYER_TITLES)
    ReDim Player(index).Char(CharNum).KillNpcs(MAX_NPCS)
    ReDim Player(index).Char(CharNum).Quests(MAX_PLAYER_QUESTS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizePlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear temp player
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Call ResizeTempPlayer(index)
    Set TempPlayer(index).Buffer = New clsBuffer
    
    ' Clear player
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Call ResizePlayer(index)
    Call SetPlayerLogin(index, vbNullString)
    
    ' Clear char
    For i = 1 To MAX_PLAYER_CHARS
        Call ClearChar(index, i)
    Next

    ' Clear player of list
    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(index).Char(CharNum)), LenB(Player(index).Char(CharNum)))
    Call ResizeChar(index, CharNum)
    Player(index).Char(CharNum).Name = vbNullString
    Player(index).Char(CharNum).Class = 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearChar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateClassesINI", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim StartItemCount As Long, StarSpellCount As Long, StartTitleCount As Long
    Dim x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If CheckClasses Then
        ReDim Class(Max_Classes)
        Call SaveClasses
    Else
        filename = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next n
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next n
        
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))

        ' Localization
        Class(i).StartMap = Val(GetVar(filename, "CLASS" & i, "StartMap"))
        Class(i).StartX = Val(GetVar(filename, "CLASS" & i, "StartX"))
        Class(i).StartY = Val(GetVar(filename, "CLASS" & i, "StartY"))

        ' how many starting items?
        StartItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If StartItemCount > 0 Then ReDim Class(i).StartItem(StartItemCount)
        If StartItemCount > 0 Then ReDim Class(i).StartValue(StartItemCount)
        
        ' loop for items & values
        Class(i).StartItemCount = StartItemCount
        If StartItemCount >= 1 And StartItemCount <= MAX_INV Then
            For x = 1 To StartItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting Spells?
        StarSpellCount = Val(GetVar(filename, "CLASS" & i, "StarSpellCount"))
        If StarSpellCount > 0 Then ReDim Class(i).StarSpell(StarSpellCount)
        
        ' loop for Spells
        Class(i).StarSpellCount = StarSpellCount
        If StarSpellCount >= 1 And StarSpellCount <= MAX_PLAYER_SPELLS Then
            For x = 1 To StarSpellCount
                Class(i).StarSpell(x) = Val(GetVar(filename, "CLASS" & i, "StarSpell" & x))
            Next
        End If
        
        ' how many starting titles?
        StartTitleCount = Val(GetVar(filename, "CLASS" & i, "StartTitleCount"))
        If StartTitleCount > 0 Then ReDim Class(i).StartTitle(StartTitleCount)
        
        ' loop for titles
        Class(i).StartTitleCount = StartTitleCount
        If StartTitleCount >= 1 And StartTitleCount <= MAX_PLAYER_TITLES Then
            For x = 1 To StartTitleCount
                Class(i).StartTitle(x) = Val(GetVar(filename, "CLASS" & i, "StartTitle" & x))
            Next
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadClasses", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", Str$(Class(i).Stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", Str$(Class(i).Stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", Str$(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", Str$(Class(i).Stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", Str$(Class(i).Stat(Stats.Willpower)))
        Call PutVar(filename, "CLASS" & i, "StartMap", Str$(Class(i).StartMap))
        Call PutVar(filename, "CLASS" & i, "StartX", Str$(Class(i).StartX))
        Call PutVar(filename, "CLASS" & i, "StartY", Str$(Class(i).StartY))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, Str$(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, Str$(Class(i).StartValue(x)))
        Next
        ' loop for Spells
        For x = 1 To UBound(Class(i).StarSpell)
            Call PutVar(filename, "CLASS" & i, "StarSpell" & x, Str$(Class(i).StarSpell(x)))
        Next
        ' loop for titles
        For x = 1 To UBound(Class(i).StartTitle)
            Call PutVar(filename, "CLASS" & i, "StartTitle" & x, Str$(Class(i).StartTitle(x)))
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveClasses", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' File path
    filename = App.Path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckClasses", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub ResizeClass(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Class(index).Stat(Stats.Stat_Count - 1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearClasses()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Call ResizeClass(i)
        Class(i).Name = vbNullString
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearClasses", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***********
' ** Items **
' ***********
Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim F  As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\items\item" & ItemNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(ItemNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Item(index).Add_Stat(Stats.Stat_Count - 1)
    ReDim Item(index).Stat_Req(Stats.Stat_Count - 1)
    ReDim Item(index).BagItem(MAX_BAG)
    ReDim Item(index).BagValue(MAX_BAG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearItem(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Call ResizeItem(index)
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShop(ByVal ShopNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\shops\shop" & ShopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(ShopNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

    DoEvents
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckShops()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
            Call SaveShop(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ResizeShop(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ReDim Shop(index).TradeItem(MAX_TRADES)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShop(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Call ResizeShop(index)
    Shop(index).Name = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShop", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearShops()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearShops", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\Spells\Spells" & SpellNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(SpellNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.Path & "\data\Spells\Spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckSpells()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\Spells\Spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpell(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1 ' Needs to be 1 for the Spell editor
    Spell(index).Desc = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpell", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearSpells()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** NPCs **
' **********

Sub SaveNpc(ByVal NpcNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Npc(NpcNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Npc(i)
        Close #F
    Next
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeNpc(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Npc(index).DropChance(MAX_NPC_DROPS)
    ReDim Npc(index).DropItem(MAX_NPC_DROPS)
    ReDim Npc(index).DropItemValue(MAX_NPC_DROPS)
    ReDim Npc(index).Stat(Stats.Stat_Count - 1)
    ReDim Npc(index).Quest(MAX_NPC_QUESTS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearNpc(ByVal index As Long)
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Call ResizeNpc(index)
    Npc(index).Name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).HP = 1
    
    For i = 1 To Stats.Stat_Count - 1
        Npc(index).Stat(i) = 1
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** Resources **
' **********

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Resource(ResourceNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Resource(i)
        Close #F
    Next
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckResources()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResource(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResource", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearResources()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Animation(AnimationNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Animation(i)
        Close #F
    Next
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckAnimations()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeAnimation(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Animation(index).Frames(1)
    ReDim Animation(index).LoopCount(1)
    ReDim Animation(index).LoopTime(1)
    ReDim Animation(index).Sprite(1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearAnimation(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Call ResizeAnimation(index)
    Animation(index).Name = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearAnimations()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\maps\map" & MapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(MapNum).Name
    Put #F, , Map(MapNum).Music
    Put #F, , Map(MapNum).Revision
    Put #F, , Map(MapNum).Moral
    Put #F, , Map(MapNum).Panorama
    Put #F, , Map(MapNum).Red
    Put #F, , Map(MapNum).Green
    Put #F, , Map(MapNum).Blue
    Put #F, , Map(MapNum).Alpha
    Put #F, , Map(MapNum).Fog
    Put #F, , Map(MapNum).FogSpeed
    Put #F, , Map(MapNum).FogOpacity
    Put #F, , Map(MapNum).Up
    Put #F, , Map(MapNum).Down
    Put #F, , Map(MapNum).Left
    Put #F, , Map(MapNum).Right
    Put #F, , Map(MapNum).UpLeft
    Put #F, , Map(MapNum).UpRight
    Put #F, , Map(MapNum).DownLeft
    Put #F, , Map(MapNum).DownRight
    Put #F, , Map(MapNum).BootMap
    Put #F, , Map(MapNum).BootX
    Put #F, , Map(MapNum).BootY
    Put #F, , Map(MapNum).MaxX
    Put #F, , Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            Put #F, , Map(MapNum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(MapNum).Npc(x)
    Next
    Close #F
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Panorama
        Get #F, , Map(i).Red
        Get #F, , Map(i).Green
        Get #F, , Map(i).Blue
        Get #F, , Map(i).Alpha
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).FogOpacity
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).UpLeft
        Get #F, , Map(i).UpRight
        Get #F, , Map(i).DownLeft
        Get #F, , Map(i).DownRight
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(Map(i).MaxX, Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).Npc(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next

        Close #F

        CacheDoors i
        CacheResources i
        DoEvents
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMaps", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckMaps()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckMaps", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, index)), LenB(MapItem(MapNum, index)))

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal MapNum As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim MapNpc(MapNum).Npc(MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(MapNum).Npc(index)), LenB(MapNpc(MapNum).Npc(index)))
    
    For i = 1 To MAX_MAP_NPCS
        ReDim MapNpc(MapNum).Npc(i).DoT(MAX_DOTS)
        ReDim MapNpc(MapNum).Npc(i).HoT(MAX_DOTS)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap(ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    ReDim Map(MapNum).Npc(MAX_MAP_NPCS)
    Map(MapNum).Name = vbNullString
    Map(MapNum).MaxX = MAX_MAPX
    Map(MapNum).MaxY = MAX_MAPY
    Map(MapNum).FogOpacity = 255
    ReDim Map(MapNum).Tile(Map(MapNum).MaxX, Map(MapNum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMaps()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMaps", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetClassName = Trim$(Class(ClassNum).Name)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetClassName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case Vital
        Case Vitals.HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case Vitals.MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Stats.Intelligence) * 10) + 2
            End With
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetClassMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetClassStat = Class(ClassNum).Stat(Stat)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetClassStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SaveBank(ByVal index As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\banks\" & Trim$(Player(index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(index)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveBank", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadBank(ByVal index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ClearBank(index)

    filename = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Bank(index)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadBank", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearBank(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Bank(index)), LenB(Bank(index)))
    ReDim Bank(index).Item(MAX_BANK)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearBank", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearParty(ByVal partyNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Party(partyNum)), LenB(Party(partyNum)))
    ReDim Party(partyNum).Member(MAX_PARTY_MEMBERS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearParty", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***********
' ** doors **
' ***********
Sub SaveDoor(ByVal DoorNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\doors\door" & DoorNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Door(DoorNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadDoors()
    Dim filename As String
    Dim i As Integer
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckDoors

    For i = 1 To MAX_DOORS
        filename = App.Path & "\data\doors\door" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Door(i)
        Close #F
    Next
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckDoors()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOORS
        If Not FileExist("\data\doors\door" & i & ".dat") Then
            Call SaveDoor(i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeDoor(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Door(index).Stat_Req(Stats.Stat_Count - 1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearDoor(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Door(index)), LenB(Door(index)))
    Call ResizeDoor(index)
    Door(index).Name = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearDoors()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOORS
        Call ClearDoor(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    filename = App.Path & "\data\quests\quest" & QuestNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Quest(QuestNum)
    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadQuests()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckQuests

    For i = 1 To MAX_QUESTS
        filename = App.Path & "\data\quests\quest" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Quest(i)
        Close #F
    Next
    
    DoEvents

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadQuests", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckQuests()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS

        If Not FileExist("\Data\quests\quest" & i & ".dat") Then
            Call SaveQuest(i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckQuests", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeQuest(ByVal index As Long)
    Dim i As Byte, x As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Quest(index).VitalRew(Vitals.Vital_Count - 1)
    ReDim Quest(index).StatRew(Stats.Stat_Count - 1)
    ReDim Quest(index).StatReq(Stats.Stat_Count - 1)
    ReDim Quest(index).Task(MAX_QUEST_TASKS)

    For i = 1 To MAX_QUEST_TASKS
        ReDim Quest(index).Task(i).Message(3)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearQuest(ByVal index As Long)
    Dim i As Byte, x As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Quest(index)), LenB(Quest(index)))
    Call ResizeQuest(index)
    Quest(index).Name = vbNullString
    Quest(index).Description = vbNullString
    
    For i = 1 To MAX_QUEST_TASKS
        For x = 1 To 3
            Quest(index).Task(i).Message(x) = vbNullString
        Next
    
        Quest(index).Task(i).Num = 1
        Quest(index).Task(i).Value = 1
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearQuests()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        Call ClearQuest(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearQuests", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ************
' ** Titles **
' ************
Sub SaveTitle(ByVal TitleNum As Long)
    Dim filename As String
    Dim F As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    filename = App.Path & "\data\titles\title" & TitleNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Title(TitleNum)
    Close #F
    
    Exit Sub
errorhandler:
    HandleError "SaveTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadTitles()
    Dim filename As String
    Dim i As Long
    Dim F As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CheckTitles

    For i = 1 To MAX_TITLES
        filename = App.Path & "\data\titles\title" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Title(i)
        Close #F
    Next
    
    DoEvents

    Exit Sub
errorhandler:
    HandleError "LoadTitles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckTitles()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_TITLES
        If Not FileExist("\Data\titles\title" & i & ".dat") Then
            Call SaveTitle(i)
        End If
    Next

    Exit Sub
errorhandler:
    HandleError "CheckTitles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeTitle(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Title(index).VitalRew(Vitals.Vital_Count - 1)
    ReDim Title(index).StatRew(Stats.Stat_Count - 1)
    ReDim Title(index).StatReq(Stats.Stat_Count - 1)
    
    Exit Sub
errorhandler:
    HandleError "ResizeTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearTitle(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Title(index)), LenB(Title(index)))
    Call ResizeTitle(index)
    Title(index).Name = vbNullString
    
    Exit Sub
errorhandler:
    HandleError "ClearTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearTitles()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_TITLES
        Call ClearTitle(i)
    Next
    
    Exit Sub
errorhandler:
    HandleError "ClearTitles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
