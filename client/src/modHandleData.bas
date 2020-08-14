Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub HandleDataPackets(PacketNum As Long, Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Checks which is the command to run
    If PacketNum = ServerPackets.SAlertMsg Then HandleAlertMsg Data()
    If PacketNum = ServerPackets.SLoginOk Then HandleLoginOk Data()
    If PacketNum = ServerPackets.SNewCharClasses Then HandleNewCharClasses Data()
    If PacketNum = ServerPackets.SClassesData Then HandleClassesData Data()
    If PacketNum = ServerPackets.SInGame Then HandleInGame Data()
    If PacketNum = ServerPackets.SPlayerInv Then HandlePlayerInv Data()
    If PacketNum = ServerPackets.SPlayerInvUpdate Then HandlePlayerInvUpdate Data()
    If PacketNum = ServerPackets.SPlayerWornEq Then HandlePlayerWornEq Data()
    If PacketNum = ServerPackets.SPlayerHp Then HandlePlayerHp Data()
    If PacketNum = ServerPackets.SPlayerMp Then HandlePlayerMp Data()
    If PacketNum = ServerPackets.SPlayerStats Then HandlePlayerStats Data()
    If PacketNum = ServerPackets.SPlayerData Then HandlePlayerData Data()
    If PacketNum = ServerPackets.SPlayerMove Then HandlePlayerMove Data()
    If PacketNum = ServerPackets.SNpcMove Then HandleNpcMove Data()
    If PacketNum = ServerPackets.SPlayerDir Then HandlePlayerDir Data()
    If PacketNum = ServerPackets.SNpcDir Then HandleNpcDir Data()
    If PacketNum = ServerPackets.SPlayerXY Then HandlePlayerXY Data()
    If PacketNum = ServerPackets.SPlayerXYMap Then HandlePlayerXYMap Data()
    If PacketNum = ServerPackets.SAttack Then HandleAttack Data()
    If PacketNum = ServerPackets.SNpcAttack Then HandleNpcAttack Data()
    If PacketNum = ServerPackets.SCheckForMap Then HandleCheckForMap Data()
    If PacketNum = ServerPackets.SMapData Then HandleMapData Data()
    If PacketNum = ServerPackets.SMapItemData Then HandleMapItemData Data()
    If PacketNum = ServerPackets.SMapNpcData Then HandleMapNpcData Data()
    If PacketNum = ServerPackets.SMapDone Then HandleMapDone
    If PacketNum = ServerPackets.SGlobalMsg Then HandleGlobalMsg Data()
    If PacketNum = ServerPackets.SAdminMsg Then HandleAdminMsg Data()
    If PacketNum = ServerPackets.SPlayerMsg Then HandlePlayerMsg Data()
    If PacketNum = ServerPackets.SMapMsg Then HandleMapMsg Data()
    If PacketNum = ServerPackets.SSpawnItem Then HandleSpawnItem Data()
    If PacketNum = ServerPackets.SUpdateItem Then HandleUpdateItem Data()
    If PacketNum = ServerPackets.SSpawnNpc Then HandleSpawnNpc Data()
    If PacketNum = ServerPackets.SNpcDead Then HandleNpcDead Data()
    If PacketNum = ServerPackets.SUpdateNpc Then HandleUpdateNpc Data()
    If PacketNum = ServerPackets.SEditMap Then HandleEditMap
    If PacketNum = ServerPackets.SUpdateShop Then HandleUpdateShop Data()
    If PacketNum = ServerPackets.SUpdateSpell Then HandleUpdateSpell Data()
    If PacketNum = ServerPackets.SSpells Then HandleSpells Data()
    If PacketNum = ServerPackets.SLeft Then HandleLeft Data()
    If PacketNum = ServerPackets.SResourceCache Then HandleResourceCache Data()
    If PacketNum = ServerPackets.SUpdateResource Then HandleUpdateResource Data()
    If PacketNum = ServerPackets.SSendPing Then HandleSendPing Data()
    If PacketNum = ServerPackets.SActionMsg Then HandleActionMsg Data()
    If PacketNum = ServerPackets.SPlayerEXP Then HandlePlayerExp Data()
    If PacketNum = ServerPackets.SBlood Then HandleBlood Data()
    If PacketNum = ServerPackets.SUpdateAnimation Then HandleUpdateAnimation Data()
    If PacketNum = ServerPackets.SAnimation Then HandleAnimation Data()
    If PacketNum = ServerPackets.SMapNpcVitals Then HandleMapNpcVitals Data()
    If PacketNum = ServerPackets.SCooldown Then HandleCooldown Data()
    If PacketNum = ServerPackets.SClearSpellBuffer Then HandleClearSpellBuffer Data()
    If PacketNum = ServerPackets.SSayMsg Then HandleSayMsg Data()
    If PacketNum = ServerPackets.SOpenShop Then HandleOpenShop Data()
    If PacketNum = ServerPackets.SResetShopAction Then HandleResetShopAction Data()
    If PacketNum = ServerPackets.SStunned Then HandleStunned Data()
    If PacketNum = ServerPackets.SMapWornEq Then HandleMapWornEq Data()
    If PacketNum = ServerPackets.SBank Then HandleBank Data()
    If PacketNum = ServerPackets.STrade Then HandleTrade Data()
    If PacketNum = ServerPackets.SCloseTrade Then HandleCloseTrade Data()
    If PacketNum = ServerPackets.STradeUpdate Then HandleTradeUpdate Data()
    If PacketNum = ServerPackets.STradeStatus Then HandleTradeStatus Data()
    If PacketNum = ServerPackets.STarget Then HandleTarget Data()
    If PacketNum = ServerPackets.SHotbar Then HandleHotbar Data()
    If PacketNum = ServerPackets.SHighIndex Then HandleHighIndex Data()
    If PacketNum = ServerPackets.SSound Then HandleSound Data()
    If PacketNum = ServerPackets.SPartyUpdate Then HandlePartyUpdate Data()
    If PacketNum = ServerPackets.SPartyVitals Then HandlePartyVitals Data()
    If PacketNum = ServerPackets.SDoorCache Then HandleDoorCache Data()
    If PacketNum = ServerPackets.SUpdateDoor Then HandleUpdateDoor Data()
    If PacketNum = ServerPackets.SDialogue Then HandleDialogue Data()
    If PacketNum = ServerPackets.SUpdateQuest Then HandleUpdateQuest Data()
    If PacketNum = ServerPackets.SQuestCommand Then HandleQuestCommand Data()
    If PacketNum = ServerPackets.SCloseShop Then HandleCloseShop Data()
    If PacketNum = ServerPackets.SCloseTrade Then HandleCloseTrade Data()
    If PacketNum = ServerPackets.SCloseBank Then HandleCloseBank Data()
    If PacketNum = ServerPackets.SUpdateTitle Then HandleUpdateTitle Data()
    If PacketNum = ServerPackets.STitles Then HandleTitles Data()
    If PacketNum = ServerPackets.SMapReport Then HandleMapReport Data()
    If PacketNum = ServerPackets.SCharData Then HandleCharData Data()
    If PacketNum = ServerPackets.SItemEditor Then HandleItemEditor
    If PacketNum = ServerPackets.SNpcEditor Then HandleNpcEditor
    If PacketNum = ServerPackets.SShopEditor Then HandleShopEditor
    If PacketNum = ServerPackets.SSpellEditor Then HandleSpellEditor
    If PacketNum = ServerPackets.SResourceEditor Then HandleResourceEditor
    If PacketNum = ServerPackets.SAnimationEditor Then HandleAnimationEditor
    If PacketNum = ServerPackets.SDoorEditor Then HandleDoorEditor
    If PacketNum = ServerPackets.SQuestEditor Then HandleQuestEditor
    If PacketNum = ServerPackets.STitleEditor Then HandleTitleEditor
    If PacketNum = ServerPackets.SChatMsg Then HandleChatMsg Data()

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDataPackets", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MsgType = Buffer.ReadLong
    
    If MsgType < 0 Or MsgType >= SMSG_COUNT Then
        DestroyGame
        Exit Sub
    End If
    
    HandleDataPackets MsgType, Buffer.ReadBytes(Buffer.Length)
    Set Buffer = Nothing
    
    ' Add packet info to debugger
    If MyIndex > 0 Then
        If GetPlayerAccess(MyIndex) >= ADMIN_MONITOR Then
            Call DebugAdd(Time & " Received - " & UBound(Data()) & " bytes" & " - " & MsgType, 0)
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAlertMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmLoad.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = True
    frmMenu.picRegister.Visible = False
    frmMenu.picCharacters.Visible = False
    frmMenu.Visible = True

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call MsgBox(Buffer.ReadString, vbOKOnly, Options.Game_Name)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleLoginOk(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' save options
    Options.SavePass = frmMenu.chkPass.Value
    Options.Username = Trim$(frmMenu.txtLUser.text)

    If frmMenu.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    
    SaveOptions
    
    ' Now we can receive game data
    MyIndex = Buffer.ReadLong
    
    ' player high myindex
    Player_HighIndex = Buffer.ReadLong
    
    Set Buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("Receiving game data...")
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByRef Data() As Byte)
    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong
    ReDim Class(Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next
            
            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = Buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmMenu.picCharacter.Visible = True
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picCharacters.Visible = False
    frmLoad.Visible = False

    frmMenu.cmbClass.Clear
    For i = 1 To Max_Classes
        frmMenu.cmbClass.AddItem Trim$(Class(i).Name)
    Next

    frmMenu.cmbClass.ListIndex = 0

    newCharSex = SEX_MALE
    newCharSprite = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleClassesData(ByRef Data() As Byte)
    Dim n As Long
    Dim i As Long
    Dim z As Long, x As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1
    ' Max classes
    Max_Classes = Buffer.ReadLong 'CByte(Parse(n))
    ReDim Class(Max_Classes)
    n = n + 1

    For i = 1 To Max_Classes

        With Class(i)
            .Name = Buffer.ReadString 'Trim$(Parse(n))

            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .MaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .MaleSprite(x) = Buffer.ReadLong
            Next
            
            ' get array size
            z = Buffer.ReadLong
            ' redim array
            ReDim .FemaleSprite(0 To z)
            ' loop-receive data
            For x = 0 To z
                .FemaleSprite(x) = Buffer.ReadLong
            Next
                            
            For x = 1 To Stats.Stat_Count - 1
                .Stat(x) = Buffer.ReadLong
            Next
        End With

        n = n + 10
    Next

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClassesData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleInGame(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InGame = True
    InMenu = False
    Call GameInit
    Call GameLoop
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInv(ByRef Data() As Byte)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, Buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, Buffer.ReadLong)
        n = n + 2
    Next
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerInvUpdate(ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong 'CLng(Parse(1))
    Call SetPlayerInvItemNum(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(2)))
    Call SetPlayerInvItemValue(MyIndex, n, Buffer.ReadLong) 'CLng(Parse(3)))
    ' changes, clear drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePlayerWornEq(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(MyIndex, Buffer.ReadLong, Shield)
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapWornEq(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim playerNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    playerNum = Buffer.ReadLong
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Armor)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Weapon)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Helmet)
    Call SetPlayerEquipment(playerNum, Buffer.ReadLong, Shield)
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerHp(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.HP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.HP, Buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
        frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP) & " (" & Int(GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP) * 100) & "%)"
        ' hp bar
        frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
    End If
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMp(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Player(MyIndex).MaxVital(Vitals.MP) = Buffer.ReadLong
    Call SetPlayerVital(MyIndex, Vitals.MP, Buffer.ReadLong)

    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP) & " (" & Int(GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP) * 100) & "%)"
        ' mp bar
        frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / SPRBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / SPRBar_Width)) * SPRBar_Width
    End If
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerStats(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat MyIndex, i, Buffer.ReadLong
        frmMain.lblCharStat(i).Caption = GetPlayerStat(MyIndex, i)
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerExp(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim TNL As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SetPlayerExp MyIndex, Buffer.ReadLong
    TNL = Buffer.ReadLong
    frmMain.lblEXP.Caption = GetPlayerExp(MyIndex) & "/" & TNL & " (" & Int(GetPlayerExp(MyIndex) / TNL * 100) & "%)"
    ' mp bar
    frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerExp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerData(ByRef Data() As Byte)
    Dim i As Long, x As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong

    Call SetPlayerName(i, Buffer.ReadString)
    Call SetPlayerLevel(i, Buffer.ReadLong)
    Call SetPlayerPOINTS(i, Buffer.ReadLong)
    Call SetPlayerSprite(i, Buffer.ReadLong)
    Call SetPlayerMap(i, Buffer.ReadLong)
    Call SetPlayerX(i, Buffer.ReadLong)
    Call SetPlayerY(i, Buffer.ReadLong)
    Call SetPlayerDir(i, Buffer.ReadLong)
    Call SetPlayerAccess(i, Buffer.ReadLong)
    Call SetPlayerPK(i, Buffer.ReadLong)
    Call SetPlayerKillPlayers(i, Buffer.ReadLong)
    Call SetPlayerTUsing(i, Buffer.ReadLong)

    For x = 1 To MAX_PLAYER_TITLES
        Call SetPlayerTitle(i, x, Buffer.ReadLong)
    Next

    For x = 1 To MAX_NPCS
        Call SetPlayerKillNpcs(i, x, Buffer.ReadInteger)
    Next

    For x = 1 To MAX_PLAYER_QUESTS
        Call SetPlayerQuestNum(i, x, Buffer.ReadInteger)
        Call SetPlayerQuestStatus(i, x, Buffer.ReadByte)
        Call SetPlayerQuestPart(i, x, Buffer.ReadByte)
    Next

    For x = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, x, Buffer.ReadLong
    Next
    Set Buffer = Nothing

    ' Check if the player is the client player
    If i = MyIndex Then
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        DirUpLeft = False
        DirUpRight = False
        DirDownLeft = False
        DirDownRight = False
        
        ' Set the character windows
        frmMain.lblCharName = GetPlayerName(MyIndex) & " - Level " & GetPlayerLevel(MyIndex)
        
        For x = 1 To Stats.Stat_Count - 1
            frmMain.lblCharStat(x).Caption = GetPlayerStat(MyIndex, x)
        Next

        ' Set training label visiblity depending on points
        frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
        If GetPlayerPOINTS(MyIndex) > 0 Then
            For x = 1 To Stats.Stat_Count - 1
                If GetPlayerStat(MyIndex, x) < 255 Then
                    frmMain.lblTrainStat(x).Visible = True
                Else
                    frmMain.lblTrainStat(x).Visible = False
                End If
            Next
        Else
            For x = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(x).Visible = False
            Next
        End If

        PlayerQuests
    End If

    ' Make sure they aren't walking
    Player(i).Moving = 0
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMove(ByRef Data() As Byte)
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Dir As Long
    Dim n As Byte
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    n = Buffer.ReadLong
    Set Buffer = Nothing
    Call SetPlayerX(i, x)
    Call SetPlayerY(i, y)
    Call SetPlayerDir(i, Dir)
    Player(i).XOffset = 0
    Player(i).YOffset = 0
    Player(i).Moving = n

    Select Case GetPlayerDir(i)
        Case DIR_UP
            Player(i).YOffset = PIC_Y
        Case DIR_DOWN
            Player(i).YOffset = PIC_Y * -1
        Case DIR_LEFT
            Player(i).XOffset = PIC_X
        Case DIR_RIGHT
            Player(i).XOffset = PIC_X * -1
        Case DIR_UP_LEFT
            Player(i).YOffset = PIC_Y
            Player(i).XOffset = PIC_X
        Case DIR_UP_RIGHT
            Player(i).YOffset = PIC_Y
            Player(i).XOffset = PIC_X * -1
        Case DIR_DOWN_LEFT
            Player(i).YOffset = PIC_Y * -1
            Player(i).XOffset = PIC_X
        Case DIR_DOWN_RIGHT
            Player(i).YOffset = PIC_Y * -1
            Player(i).XOffset = PIC_X * -1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcMove(ByRef Data() As Byte)
    Dim MapNpcNum As Long
    Dim x As Long
    Dim y As Long
    Dim Dir As Long
    Dim Movement As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MapNpcNum = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Movement = Buffer.ReadLong
    Set Buffer = Nothing

    With MapNpc(MapNpcNum)
        .x = x
        .y = y
        .Dir = Dir
        .XOffset = 0
        .YOffset = 0
        .Moving = Movement

        Select Case .Dir
            Case DIR_UP
                .YOffset = PIC_Y
            Case DIR_DOWN
                .YOffset = PIC_Y * -1
            Case DIR_LEFT
                .XOffset = PIC_X
            Case DIR_RIGHT
                .XOffset = PIC_X * -1
            Case DIR_UP_LEFT
                .YOffset = PIC_Y
                .XOffset = PIC_X
            Case DIR_UP_RIGHT
                .YOffset = PIC_Y
                .XOffset = PIC_X * -1
            Case DIR_DOWN_LEFT
                .YOffset = PIC_Y * -1
                .XOffset = PIC_X
            Case DIR_DOWN_RIGHT
                .YOffset = PIC_Y * -1
                .XOffset = PIC_X * -1
        End Select

    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerDir(ByRef Data() As Byte)
    Dim i As Long
    Dim Dir As Byte
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerDir(i, Dir)
    Set Buffer = Nothing

    With Player(i)
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDir(ByRef Data() As Byte)
    Dim i As Long
    Dim Dir As Byte
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Set Buffer = Nothing

    With MapNpc(i)
        .Dir = Dir
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXY(ByRef Data() As Byte)
    Dim x As Long
    Dim y As Long
    Dim Dir As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(MyIndex, x)
    Call SetPlayerY(MyIndex, y)
    Call SetPlayerDir(MyIndex, Dir)
    Set Buffer = Nothing
    
    ' Make sure they aren't walking
    Player(MyIndex).Moving = 0
    Player(MyIndex).XOffset = 0
    Player(MyIndex).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXY", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerXYMap(ByRef Data() As Byte)
    Dim x As Long
    Dim y As Long
    Dim Dir As Long
    Dim Buffer As clsBuffer
    Dim thePlayer As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    thePlayer = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    Dir = Buffer.ReadLong
    Call SetPlayerX(thePlayer, x)
    Call SetPlayerY(thePlayer, y)
    Call SetPlayerDir(thePlayer, Dir)
    Set Buffer = Nothing
    
    ' Make sure they aren't walking
    Player(thePlayer).Moving = 0
    Player(thePlayer).XOffset = 0
    Player(thePlayer).YOffset = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerXYMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAttack(ByRef Data() As Byte)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    Player(i).Attacking = 1
    Player(i).AttackTimer = GetTickCount
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcAttack(ByRef Data() As Byte)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    i = Buffer.ReadLong
    ' Set player to attacking
    MapNpc(i).Attacking = 1
    MapNpc(i).AttackTimer = GetTickCount
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCheckForMap(ByRef Data() As Byte)
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim NeedMap As Byte
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' clear the blood
    For i = 1 To MAX_BYTE
        Blood(i).x = 0
        Blood(i).y = 0
        Blood(i).Sprite = 0
        Blood(i).Timer = 0
    Next
    
    ' Get map num
    x = Buffer.ReadLong
    ' Get revision
    y = Buffer.ReadLong

    If FileExist(MAP_PATH & "map" & x & MAP_EXT) Then
        Call LoadMap(x)
        ' Check to see if the revisions match
        NeedMap = 1

        If Map.Revision = y Then
            ' We do so we dont need the map
            'Call SendData(CNeedMap & SEP_CHAR & "n" & END_CHAR)
            NeedMap = 0
        End If

    Else
        NeedMap = 1
    End If

    ' Either the revisions didn't match or we dont have the map, so we need it
    Set Buffer = New clsBuffer
    Buffer.WriteLong ClientPackets.CNeedMap
    Buffer.WriteLong NeedMap
    SendData Buffer.ToArray()
    Set Buffer = Nothing

    ' Check if we get a map from Soundeone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleMapData(ByRef Data() As Byte)
    Dim n As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()

    MapNum = Buffer.ReadLong
    Map.Name = Buffer.ReadString
    Map.Music = Buffer.ReadByte
    Map.Revision = Buffer.ReadLong
    Map.Moral = Buffer.ReadByte
    Map.Panorama = Buffer.ReadLong
    Map.Red = Buffer.ReadByte
    Map.Green = Buffer.ReadByte
    Map.Blue = Buffer.ReadByte
    Map.Alpha = Buffer.ReadByte
    Map.Fog = Buffer.ReadByte
    Map.FogSpeed = Buffer.ReadByte
    Map.FogOpacity = Buffer.ReadByte
    Map.Up = Buffer.ReadLong
    Map.Down = Buffer.ReadLong
    Map.Left = Buffer.ReadLong
    Map.Right = Buffer.ReadLong
    Map.UpLeft = Buffer.ReadLong
    Map.UpRight = Buffer.ReadLong
    Map.DownLeft = Buffer.ReadLong
    Map.DownRight = Buffer.ReadLong
    Map.BootMap = Buffer.ReadLong
    Map.BootX = Buffer.ReadByte
    Map.BootY = Buffer.ReadByte
    Map.MaxX = Buffer.ReadByte
    Map.MaxY = Buffer.ReadByte

    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                For n = 1 To MAX_MAP_LAYERS
                    Map.Tile(x, y).Layer(i, n).x = Buffer.ReadLong
                    Map.Tile(x, y).Layer(i, n).y = Buffer.ReadLong
                    Map.Tile(x, y).Layer(i, n).Tileset = Buffer.ReadLong
                Next
            Next
            Map.Tile(x, y).Type = Buffer.ReadByte
            Map.Tile(x, y).Data1 = Buffer.ReadLong
            Map.Tile(x, y).Data2 = Buffer.ReadLong
            Map.Tile(x, y).Data3 = Buffer.ReadLong
            Map.Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map.Npc(x) = Buffer.ReadLong
        n = n + 1
    Next

    Set Buffer = Nothing
    
    ' Save the map
    Call SaveMap(MapNum)

    ' Check if we get a map from Soundeone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeDialogue

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapItemData(ByRef Data() As Byte)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .Num = Buffer.ReadLong
            .Value = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
        End With
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByRef Data() As Byte)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(i)
            .Num = Buffer.ReadLong
            .x = Buffer.ReadLong
            .y = Buffer.ReadLong
            .Dir = Buffer.ReadLong
            .Vital(HP) = Buffer.ReadLong
        End With
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapDone()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' clear the action msgs
    For i = 1 To MAX_BYTE
        ClearActionMsg (i)
    Next i
    Action_HighIndex = 1
    
    If Map.Music > 0 Then
        Play_Music Map.Music
    Else
        Stop_Music
    End If
    
    ' re-position the map name
    Call UpdateDrawMapName
    
    ' get the npc high myindex
    For i = MAX_MAP_NPCS To 1 Step -1
        If MapNpc(i).Num > 0 Then
            Npc_HighIndex = i + 1
            Exit For
        End If
    Next
    ' make sure we're not overflowing
    If Npc_HighIndex > MAX_MAP_NPCS Then Npc_HighIndex = MAX_MAP_NPCS

    GettingMap = False
    CanMoveNow = True
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleGlobalMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAdminMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnItem(ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapItem(n)
        .Num = Buffer.ReadLong
        .Value = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
    End With
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateItem(ByRef Data() As Byte)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Item(n).Name = Buffer.ReadString
    Item(n).Desc = Buffer.ReadString
    Item(n).Sound = Buffer.ReadByte
    Item(n).Pic = Buffer.ReadLong
    Item(n).Type = Buffer.ReadByte
    Item(n).Data1 = Buffer.ReadLong
    Item(n).Data2 = Buffer.ReadLong
    Item(n).Data3 = Buffer.ReadLong
    Item(n).ClassReq = Buffer.ReadLong
    Item(n).AccessReq = Buffer.ReadLong
    Item(n).LevelReq = Buffer.ReadLong
    Item(n).Price = Buffer.ReadLong
    Item(n).Rarity = Buffer.ReadByte
    Item(n).Speed = Buffer.ReadLong
    Item(n).Animation = Buffer.ReadLong
    Item(n).Paperdoll = Buffer.ReadLong
    Item(n).AddHP = Buffer.ReadLong
    Item(n).AddMP = Buffer.ReadLong
    Item(n).AddEXP = Buffer.ReadLong
    Item(n).Damage = Buffer.ReadLong
    Item(n).Protection = Buffer.ReadLong
    
    For i = 1 To Stats.Stat_Count - 1
        Item(n).Add_Stat(i) = Buffer.ReadByte
        Item(n).Stat_Req(i) = Buffer.ReadByte
    Next
    
    For i = 1 To MAX_BAG
        Item(n).BagItem(i) = Buffer.ReadLong
        Item(n).BagValue(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    tmpCurrencyItem = 0
    CurrencyMenu = 0 ' clear

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateAnimation(ByRef Data() As Byte)
    Dim n As Long, i As Byte
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Animation(n).Name = Buffer.ReadString
    Animation(n).Sound = Buffer.ReadByte
    
    For i = 0 To 1
        Animation(n).Sprite(i) = Buffer.ReadLong
        Animation(n).Frames(i) = Buffer.ReadLong
        Animation(n).LoopCount(i) = Buffer.ReadLong
        Animation(n).looptime(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpawnNpc(ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong

    With MapNpc(n)
        .Num = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .Dir = Buffer.ReadLong
        ' Client use only
        .XOffset = 0
        .YOffset = 0
        .Moving = 0
    End With
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcDead(ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Call ClearMapNpc(n)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateNpc(ByRef Data() As Byte)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    n = Buffer.ReadLong
    Npc(n).Name = Buffer.ReadString
    Npc(n).AttackSay = Buffer.ReadString
    Npc(n).Sound = Buffer.ReadByte
    Npc(n).Sprite = Buffer.ReadLong
    Npc(n).SpawnSecs = Buffer.ReadLong
    Npc(n).Behaviour = Buffer.ReadByte
    Npc(n).Range = Buffer.ReadByte
    Npc(n).HP = Buffer.ReadLong
    Npc(n).EXP = Buffer.ReadLong
    Npc(n).Animation = Buffer.ReadLong
    Npc(n).Damage = Buffer.ReadLong
    Npc(n).Level = Buffer.ReadLong
    Npc(n).ShopNum = Buffer.ReadLong
    
    For i = 1 To MAX_NPC_DROPS
        Npc(n).DropChance(i) = Buffer.ReadByte
        Npc(n).DropItem(i) = Buffer.ReadInteger
        Npc(n).DropItemValue(i) = Buffer.ReadLong
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Npc(n).Stat(i) = Buffer.ReadByte
    Next
    
    For i = 1 To MAX_NPC_QUESTS
        Npc(n).Quest(i) = Buffer.ReadInteger
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateResource(ByRef Data() As Byte)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong
    Resource(ResourceNum).Name = Buffer.ReadString
    Resource(ResourceNum).SuccessMessage = Buffer.ReadString
    Resource(ResourceNum).EmptyMessage = Buffer.ReadString
    Resource(ResourceNum).Sound = Buffer.ReadByte
    Resource(ResourceNum).Type = Buffer.ReadByte
    Resource(ResourceNum).ResourceImage = Buffer.ReadLong
    Resource(ResourceNum).ItemReward = Buffer.ReadLong
    Resource(ResourceNum).ToolRequired = Buffer.ReadLong
    Resource(ResourceNum).Health = Buffer.ReadLong
    Resource(ResourceNum).RespawnTime = Buffer.ReadLong
    Resource(ResourceNum).ToolRequired = Buffer.ReadLong
    Resource(ResourceNum).Animation = Buffer.ReadLong
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateShop(ByRef Data() As Byte)
    Dim Buffer As clsBuffer, ShopNum As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ShopNum = Buffer.ReadLong
    Shop(ShopNum).Name = Buffer.ReadString
    Shop(ShopNum).BuyRate = Buffer.ReadLong
    
    For i = 1 To MAX_TRADES
        Shop(ShopNum).TradeItem(i).Item = Buffer.ReadLong
        Shop(ShopNum).TradeItem(i).ItemValue = Buffer.ReadLong
        Shop(ShopNum).TradeItem(i).CostItem = Buffer.ReadLong
        Shop(ShopNum).TradeItem(i).CostValue = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateSpell(ByRef Data() As Byte)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SpellNum = Buffer.ReadLong
    Spell(SpellNum).Name = Buffer.ReadString
    Spell(SpellNum).Desc = Buffer.ReadString
    Spell(SpellNum).Sound = Buffer.ReadByte
    Spell(SpellNum).Type = Buffer.ReadByte
    Spell(SpellNum).MPCost = Buffer.ReadLong
    Spell(SpellNum).LevelReq = Buffer.ReadLong
    Spell(SpellNum).AccessReq = Buffer.ReadLong
    Spell(SpellNum).ClassReq = Buffer.ReadLong
    Spell(SpellNum).CastTime = Buffer.ReadLong
    Spell(SpellNum).CDTime = Buffer.ReadLong
    Spell(SpellNum).Icon = Buffer.ReadLong
    Spell(SpellNum).Map = Buffer.ReadLong
    Spell(SpellNum).x = Buffer.ReadLong
    Spell(SpellNum).y = Buffer.ReadLong
    Spell(SpellNum).Dir = Buffer.ReadByte
    Spell(SpellNum).Vital = Buffer.ReadLong
    Spell(SpellNum).Duration = Buffer.ReadLong
    Spell(SpellNum).Interval = Buffer.ReadLong
    Spell(SpellNum).Range = Buffer.ReadLong
    Spell(SpellNum).IsAoE = Buffer.ReadLong
    Spell(SpellNum).CastAnim = Buffer.ReadLong
    Spell(SpellNum).SpellAnim = Buffer.ReadLong
    Spell(SpellNum).StunDuration = Buffer.ReadLong
    Spell(SpellNum).BaseStat = Buffer.ReadLong
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpells(ByRef Data() As Byte)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = Buffer.ReadLong
    Next

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleLeft(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Call ClearPlayer(Buffer.ReadLong)
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByRef Data() As Byte)
    Dim Buffer As clsBuffer, Resource_Num As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Resource_Index = Buffer.ReadLong
    Resource_Num = Buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        MapResource(Resource_Num).ResourceState = Buffer.ReadByte
        MapResource(Resource_Num).x = Buffer.ReadLong
        MapResource(Resource_Num).y = Buffer.ReadLong

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSendPing(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    PingEnd = GetTickCount
    Ping = PingEnd - PingStart

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleActionMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, message As String, Color As Long, tmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    message = Buffer.ReadString
    Color = Buffer.ReadLong
    tmpType = Buffer.ReadLong
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    Set Buffer = Nothing
    
    CreateActionMsg message, Color, tmpType, x, y
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBlood(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, Sprite As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong

    Set Buffer = Nothing
    
    ' randomise sprite
    Sprite = Rand(1, BloodCount)
    
    ' make sure tile doesn't already have blood
    For i = 1 To MAX_BYTE
        If Blood(i).x = x And Blood(i).y = y Then
            ' already have blood :(
            Exit Sub
        End If
    Next
    
    ' carry on with the set
    BloodIndex = BloodIndex + 1
    If BloodIndex >= MAX_BYTE Then BloodIndex = 1
    
    With Blood(BloodIndex)
        .x = x
        .y = y
        .Sprite = Sprite
        .Timer = GetTickCount
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimation(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_BYTE Then AnimationIndex = 1
    
    With AnimInstance(AnimationIndex)
        .Animation = Buffer.ReadLong
        .x = Buffer.ReadLong
        .y = Buffer.ReadLong
        .LockType = Buffer.ReadByte
        .lockindex = Buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
    End With
    
    ' play the sound if we've got one
    PlayMapSound SoundEntity.seAnimation, AnimInstance(AnimationIndex).Animation

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcVitals(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim MapNpcNum As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    MapNpcNum = Buffer.ReadLong
    For i = 1 To Vitals.Vital_Count - 1
        MapNpc(MapNpcNum).Vital(i) = Buffer.ReadLong
    Next
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapNpcVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCooldown(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    SpellCD(Slot) = GetTickCount
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleClearSpellBuffer(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSayMsg(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Access As Long
    Dim Name As String
    Dim message As String
    Dim Colour As Long
    Dim Header As String
    Dim PK As Long
    Dim saycolour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Name = Buffer.ReadString
    Access = Buffer.ReadLong
    PK = Buffer.ReadLong
    message = Buffer.ReadString
    Header = Buffer.ReadString
    saycolour = Buffer.ReadLong
    
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                Colour = RGB(255, 96, 0)
            Case 1
                Colour = QBColor(DarkGrey)
            Case 2
                Colour = QBColor(Cyan)
            Case 3
                Colour = QBColor(BrightGreen)
            Case 4
                Colour = QBColor(Yellow)
        End Select
    Else
        Colour = DX8Color(BrightRed)
    End If
    
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = Colour
    frmMain.txtChat.SelText = vbNewLine & Header & Name & ": "
    frmMain.txtChat.SelColor = saycolour
    frmMain.txtChat.SelText = message
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
   
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleOpenShop(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim ShopNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ShopNum = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    OpenShop ShopNum
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResetShopAction(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleStunned(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    StunDuration = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_BANK
        Bank.Item(i).Num = Buffer.ReadLong
        Bank.Item(i).Value = Buffer.ReadLong
    Next
    
    InBank = True
    frmMain.picBank.Visible = True

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTrade(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    InTrade = Buffer.ReadLong
    frmMain.picTrade.Visible = True
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCloseTrade(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InTrade = 0
    frmMain.picTrade.Visible = False
    frmMain.lblTradeStatus.Caption = vbNullString

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeUpdate(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim dataType As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    dataType = Buffer.ReadByte
    
    If dataType = 0 Then ' ours!
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = Buffer.ReadLong
            TradeYourOffer(i).Value = Buffer.ReadLong
        Next
        frmMain.lblYourWorth.Caption = Buffer.ReadLong & "g"
    ElseIf dataType = 1 Then 'theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = Buffer.ReadLong
            TradeTheirOffer(i).Value = Buffer.ReadLong
        Next
        frmMain.lblTheirWorth.Caption = Buffer.ReadLong & "g"
    End If
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim tradeStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeStatus = Buffer.ReadByte
    
    Set Buffer = Nothing
    
    Select Case tradeStatus
        Case 0 ' clear
            frmMain.lblTradeStatus.Caption = vbNullString
        Case 1 ' they've accepted
            frmMain.lblTradeStatus.Caption = "Other player has accepted."
        Case 2 ' you've accepted
            frmMain.lblTradeStatus.Caption = "Waiting for other player to accept."
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTarget(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    myTarget = Buffer.ReadLong
    myTargetType = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTarget", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHotbar(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = Buffer.ReadLong
        Hotbar(i).sType = Buffer.ReadByte
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleHighIndex(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Player_HighIndex = Buffer.ReadLong
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSound(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim x As Long, y As Long, entityType As Long, entityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    entityType = Buffer.ReadLong
    entityNum = Buffer.ReadLong
    
    PlayMapSound entityType, entityNum
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyUpdate(ByRef Data() As Byte)
    Dim Buffer As clsBuffer, i As Long, inParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    inParty = Buffer.ReadByte
    
    ' exit out if we're not in a party
    If inParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' reset the labels
        For i = 1 To MAX_PARTY_MEMBERS
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        Next
        ' exit out early
        Exit Sub
    End If
    
    ' carry on otherwise
    Party.Leader = Buffer.ReadLong
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = Buffer.ReadLong
        If Party.Member(i) > 0 Then
            frmMain.lblPartyMember(i).Caption = GetPlayerName(Party.Member(i))
            frmMain.imgPartyHealth(i).Visible = True
            frmMain.imgPartySpirit(i).Visible = True
        Else
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        End If
    Next
    Party.MemberCount = Buffer.ReadLong
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePartyVitals(ByRef Data() As Byte)
    Dim playerNum As Long, partyIndex As Long
    Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' which player?
    playerNum = Buffer.ReadLong
    ' set vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(playerNum).MaxVital(i) = Buffer.ReadLong
        Player(playerNum).Vital(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = playerNum Then
            partyIndex = i
        End If
    Next
    
    ' exit out if wrong data
    If partyIndex <= 0 Or partyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' hp bar
    frmMain.imgPartyHealth(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
    ' spr bar
    frmMain.imgPartySpirit(partyIndex).Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateDoor(ByRef Data() As Byte)
    Dim n As Long, i As Byte
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Door(n).Name = Buffer.ReadString
    Door(n).OpeningImage = Buffer.ReadInteger
    Door(n).OpenWith = Buffer.ReadInteger
    Door(n).Respawn = Buffer.ReadLong
    Door(n).Animation = Buffer.ReadLong
    Door(n).Sound = Buffer.ReadByte
    Door(n).Map = Buffer.ReadInteger
    Door(n).x = Buffer.ReadByte
    Door(n).y = Buffer.ReadByte
    Door(n).Dir = Buffer.ReadByte
    Door(n).LevelReq = Buffer.ReadInteger
    
    For i = 1 To Stats.Stat_Count - 1
        Door(n).Stat_Req(i) = Buffer.ReadByte
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateDoor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDoorCache(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Door_Num As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' if in map editor, we cache shit ourselves
    If InMapEditor Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Door_Index = Buffer.ReadLong
    Door_Num = Buffer.ReadLong
    Doors_Init = False

    If Door_Index > 0 Then
        ReDim Preserve MapDoor(0 To Door_Index)

        ' Localization
        MapDoor(Door_Num).x = Buffer.ReadLong
        MapDoor(Door_Num).y = Buffer.ReadLong
        ' State
        MapDoor(Door_Num).State = Buffer.ReadLong
        Doors_Init = True
    Else
        ReDim MapDoor(0 To 1)
    End If

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoorCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDialogue(ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dialogue Buffer.ReadString, Buffer.ReadString, Buffer.ReadByte, Buffer.ReadLong
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDialogue", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleUpdateQuest(ByRef Data() As Byte)
    Dim n As Long, i As Long, x As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Quest(n).Name = Buffer.ReadString
    Quest(n).Description = Buffer.ReadString
    Quest(n).Retry = Buffer.ReadLong
    Quest(n).LevelReq = Buffer.ReadInteger
    Quest(n).QuestReq = Buffer.ReadInteger
    Quest(n).ClassReq = Buffer.ReadByte
    Quest(n).SpriteReq = Buffer.ReadInteger
    Quest(n).LevelRew = Buffer.ReadInteger
    Quest(n).ExpRew = Buffer.ReadLong
    Quest(n).ClassRew = Buffer.ReadByte
    Quest(n).SpriteRew = Buffer.ReadInteger
    
    For i = 1 To Stats.Stat_Count - 1
        Quest(n).StatReq(i) = Buffer.ReadLong
        Quest(n).StatRew(i) = Buffer.ReadLong
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Quest(n).VitalRew(i) = Buffer.ReadLong
    Next
    
    For i = 1 To MAX_QUEST_TASKS
        Quest(n).Task(i).Type = Buffer.ReadByte
        Quest(n).Task(i).Instant = Buffer.ReadLong
        Quest(n).Task(i).Num = Buffer.ReadInteger
        Quest(n).Task(i).Value = Buffer.ReadLong
    
        For x = 1 To 3
            Quest(n).Task(i).message(x) = Buffer.ReadString
        Next
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateQuest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleQuestCommand(ByRef Data() As Byte)
    Dim Buffer As clsBuffer, Command As Byte, Value As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadByte
    Value = Buffer.ReadLong
    Set Buffer = Nothing
    
    Select Case Command
            ' Select npc quest
        Case 1
            frmMain.picSelectQuest.Visible = True
            UpdateSelectQuest Value
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestCommand", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleCloseShop(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Close pictures
    frmMain.picShop.Visible = False
    
    ' Reset variables
    InShop = 0
    ShopAction = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleCloseBank(ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Close pictures
    frmMain.picBank.Visible = False
    
    ' Reset variables
    InBank = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleUpdateTitle(ByRef Data() As Byte)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    n = Buffer.ReadLong
    Title(n).Name = Buffer.ReadString
    Title(n).Description = Buffer.ReadString
    Title(n).Icon = Buffer.ReadInteger
    Title(n).Type = Buffer.ReadByte
    Title(n).Color = Buffer.ReadByte
    Title(n).Sound = Buffer.ReadByte
    Title(n).UseAnimation = Buffer.ReadLong
    Title(n).RemoveAnimation = Buffer.ReadLong
    Title(n).Passive = Buffer.ReadLong
    Title(n).LevelReq = Buffer.ReadLong
    
    For i = 1 To Stats.Stat_Count - 1
        Title(n).StatReq(i) = Buffer.ReadLong
        Title(n).StatRew(i) = Buffer.ReadLong
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Title(n).VitalRew(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUpdateTitle", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleTitles(ByRef Data() As Byte)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_TITLES
        Call SetPlayerTitle(MyIndex, i, Buffer.ReadLong)
    Next

    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTitles", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call MapEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapReport(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MapNum As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    frmMapReport.lstMaps.Clear
    
    For MapNum = 1 To MAX_MAPS
        frmMapReport.lstMaps.AddItem MapNum & ": " & Buffer.ReadString
    Next MapNum
    
    frmMapReport.Show
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapReport", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleCharData(ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear the character list.
    frmMenu.lstChars.Clear
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_CHARS
        ' Sets the value of the chars
        CharData(i).Name = Buffer.ReadString
        CharData(i).Level = Buffer.ReadLong
        CharData(i).Class = Buffer.ReadLong
        
        ' Display the character information to the user.
        If Len(Trim$(CharData(i).Name)) <= 0 Then
            frmMenu.lstChars.AddItem "Free character slot"
        Else
            frmMenu.lstChars.AddItem Trim$(CharData(i).Name) & " a level " & CharData(i).Level & " " & Trim$(Class(CharData(i).Class).Name)
        End If
    Next
    Set Buffer = Nothing

    ' Reset values
    frmMenu.lstChars.ListIndex = 0

    ' Show the window
    frmMenu.picCredits.Visible = False
    frmMenu.picLogin.Visible = False
    frmMenu.picCharacter.Visible = False
    frmMenu.picRegister.Visible = False
    frmMenu.picCharacters.Visible = True
    frmLoad.Hide
    frmMenu.Show
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapReport", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleItemEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Item
        CurrentEditor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleAnimationEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Animation
        CurrentEditor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNpcEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_NPC
        CurrentEditor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNpcEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Resource
        CurrentEditor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShopEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Shop
        CurrentEditor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleSpellEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Spell
        CurrentEditor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleQuestEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Quest
        CurrentEditor = EDITOR_QUEST
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_QUESTS
            .lstIndex.AddItem i & ": " & Trim$(Quest(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        QuestEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTitleEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    With frmEditor_Title
        CurrentEditor = EDITOR_TITLE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_TITLES
            .lstIndex.AddItem i & ": " & Trim$(Title(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        TitleEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTitleEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleDoorEditor()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With frmEditor_Door
        CurrentEditor = EDITOR_DOOR
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_DOORS
            .lstIndex.AddItem i & ": " & Trim$(Door(i).Name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        DoorEditorInit
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoorEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleChatMsg(ByRef Data() As Byte)
    Dim msg As String, Color As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    msg = Buffer.ReadString
    Color = Buffer.ReadLong
    Call AddText(msg, Color)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapmsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
