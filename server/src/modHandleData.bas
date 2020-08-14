Attribute VB_Name = "modHandleData"
Option Explicit

Public Sub HandleDataPackets(PacketNum As Long, index As Long, Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Checks which is the command to run
    If PacketNum = ClientPackets.CNewAccount Then HandleNewAccount index, Data()
    If PacketNum = ClientPackets.CDelAccount Then HandleDelAccount index, Data()
    If PacketNum = ClientPackets.CLogin Then HandleLogin index, Data()
    If PacketNum = ClientPackets.CAddChar Then HandleAddChar index, Data()
    If PacketNum = ClientPackets.CSayMsg Then HandleSayMsg index, Data()
    If PacketNum = ClientPackets.CEmoteMsg Then HandleEmoteMsg index, Data()
    If PacketNum = ClientPackets.CBroadcastMsg Then HandleBroadcastMsg index, Data()
    If PacketNum = ClientPackets.CPlayerMsg Then HandlePlayerMsg index, Data()
    If PacketNum = ClientPackets.CPlayerMove Then HandlePlayerMove index, Data()
    If PacketNum = ClientPackets.CPlayerDir Then HandlePlayerDir index, Data()
    If PacketNum = ClientPackets.CUseItem Then HandleUseItem index, Data()
    If PacketNum = ClientPackets.CAttack Then HandleAttack index, Data()
    If PacketNum = ClientPackets.CUseStatPoint Then HandleUseStatPoint index, Data()
    If PacketNum = ClientPackets.CWarpMeTo Then HandleWarpMeTo index, Data()
    If PacketNum = ClientPackets.CWarpToMe Then HandleWarpToMe index, Data()
    If PacketNum = ClientPackets.CWarpTo Then HandleWarpTo index, Data()
    If PacketNum = ClientPackets.CSetSprite Then HandleSetSprite index, Data()
    If PacketNum = ClientPackets.CRequestNewMap Then HandleRequestNewMap index, Data()
    If PacketNum = ClientPackets.CMapData Then HandleMapData index, Data()
    If PacketNum = ClientPackets.CNeedMap Then HandleNeedMap index, Data()
    If PacketNum = ClientPackets.CMapGetItem Then HandleMapGetItem index, Data()
    If PacketNum = ClientPackets.CMapDropItem Then HandleMapDropItem index, Data()
    If PacketNum = ClientPackets.CMapRespawn Then HandleMapRespawn index, Data()
    If PacketNum = ClientPackets.CMapReport Then HandleMapReport index, Data()
    If PacketNum = ClientPackets.CKickPlayer Then HandleKickPlayer index, Data()
    If PacketNum = ClientPackets.CBanList Then HandleBanList index, Data()
    If PacketNum = ClientPackets.CBanDestroy Then HandleBanDestroy index, Data()
    If PacketNum = ClientPackets.CBanPlayer Then HandleBanPlayer index, Data()
    If PacketNum = ClientPackets.CRequestEditMap Then HandleRequestEditMap index, Data()
    If PacketNum = ClientPackets.CRequestEditItem Then HandleRequestEditItem index, Data()
    If PacketNum = ClientPackets.CSaveItem Then HandleSaveItem index, Data()
    If PacketNum = ClientPackets.CRequestEditNpc Then HandleRequestEditNpc index, Data()
    If PacketNum = ClientPackets.CSaveNpc Then HandleSaveNpc index, Data()
    If PacketNum = ClientPackets.CRequestEditShop Then HandleRequestEditShop index, Data()
    If PacketNum = ClientPackets.CSaveShop Then HandleSaveShop index, Data()
    If PacketNum = ClientPackets.CRequestEditSpell Then HandleRequestEditSpell index, Data()
    If PacketNum = ClientPackets.CSaveSpell Then HandleSaveSpell index, Data()
    If PacketNum = ClientPackets.CSetAccess Then HandleSetAccess index, Data()
    If PacketNum = ClientPackets.CWhosOnline Then HandleWhosOnline index, Data()
    If PacketNum = ClientPackets.CSetMotd Then HandleSetMotd index, Data()
    If PacketNum = ClientPackets.CSearch Then HandleSearch index, Data()
    If PacketNum = ClientPackets.CSpells Then HandleSpells index, Data()
    If PacketNum = ClientPackets.CCast Then HandleCast index, Data()
    If PacketNum = ClientPackets.CQuit Then HandleQuit index, Data()
    If PacketNum = ClientPackets.CSwapInvSlots Then HandleSwapInvSlots index, Data()
    If PacketNum = ClientPackets.CRequestEditResource Then HandleRequestEditResource index, Data()
    If PacketNum = ClientPackets.CSaveResource Then HandleSaveResource index, Data()
    If PacketNum = ClientPackets.CCheckPing Then HandleCheckPing index, Data()
    If PacketNum = ClientPackets.CUnequip Then HandleUnequip index, Data()
    If PacketNum = ClientPackets.CRequestPlayerData Then HandleRequestPlayerData index, Data()
    If PacketNum = ClientPackets.CRequestItems Then HandleRequestItems index, Data()
    If PacketNum = ClientPackets.CRequestNPCS Then HandleRequestNPCS index, Data()
    If PacketNum = ClientPackets.CRequestResources Then HandleRequestResources index, Data()
    If PacketNum = ClientPackets.CSpawnItem Then HandleSpawnItem index, Data()
    If PacketNum = ClientPackets.CRequestEditAnimation Then HandleRequestEditAnimation index, Data()
    If PacketNum = ClientPackets.CSaveAnimation Then HandleSaveAnimation index, Data()
    If PacketNum = ClientPackets.CRequestAnimations Then HandleRequestAnimations index, Data()
    If PacketNum = ClientPackets.CRequesSpells Then HandleRequesSpells index, Data()
    If PacketNum = ClientPackets.CRequestShops Then HandleRequestShops index, Data()
    If PacketNum = ClientPackets.CRequestLevelUp Then HandleRequestLevelUp index, Data()
    If PacketNum = ClientPackets.CForgeSpell Then HandleForgeSpell index, Data()
    If PacketNum = ClientPackets.CCloseShop Then HandleCloseShop index, Data()
    If PacketNum = ClientPackets.CBuyItem Then HandleBuyItem index, Data()
    If PacketNum = ClientPackets.CSellItem Then HandleSellItem index, Data()
    If PacketNum = ClientPackets.CChangeBankSlots Then HandleChangeBankSlots index, Data()
    If PacketNum = ClientPackets.CDepositItem Then HandleDepositItem index, Data()
    If PacketNum = ClientPackets.CWithdrawItem Then HandleWithdrawItem index, Data()
    If PacketNum = ClientPackets.CCloseBank Then HandleCloseBank index, Data()
    If PacketNum = ClientPackets.CAdminWarp Then HandleAdminWarp index, Data()
    If PacketNum = ClientPackets.CTradeRequest Then HandleTradeRequest index, Data()
    If PacketNum = ClientPackets.CAcceptTrade Then HandleAcceptTrade index, Data()
    If PacketNum = ClientPackets.CDeclineTrade Then HandleDeclineTrade index, Data()
    If PacketNum = ClientPackets.CTradeItem Then HandleTradeItem index, Data()
    If PacketNum = ClientPackets.CUntradeItem Then HandleUntradeItem index, Data()
    If PacketNum = ClientPackets.CHotbarChange Then HandleHotbarChange index, Data()
    If PacketNum = ClientPackets.CHotbarUse Then HandleHotbarUse index, Data()
    If PacketNum = ClientPackets.CSwapSpellSlots Then HandleSwapSpellSlots index, Data()
    If PacketNum = ClientPackets.CAcceptTradeRequest Then HandleAcceptTradeRequest index, Data()
    If PacketNum = ClientPackets.CDeclineTradeRequest Then HandleDeclineTradeRequest index, Data()
    If PacketNum = ClientPackets.CPartyRequest Then HandlePartyRequest index, Data()
    If PacketNum = ClientPackets.CAcceptParty Then HandleAcceptParty index, Data()
    If PacketNum = ClientPackets.CDeclineParty Then HandleDeclineParty index, Data()
    If PacketNum = ClientPackets.CPartyLeave Then HandlePartyLeave index, Data()
    If PacketNum = ClientPackets.CRequestEditDoor Then HandleRequestEditDoor index, Data()
    If PacketNum = ClientPackets.CSaveDoor Then HandleSaveDoor index, Data()
    If PacketNum = ClientPackets.CRequestDoors Then HandleRequestDoors index, Data()
    If PacketNum = ClientPackets.CCheckDoor Then HandleCheckDoor index, Data()
    If PacketNum = ClientPackets.CRequestEditQuest Then HandleRequestEditQuest index, Data()
    If PacketNum = ClientPackets.CSaveQuest Then HandleSaveQuest index, Data()
    If PacketNum = ClientPackets.CRequestQuests Then HandleRequestQuests index, Data()
    If PacketNum = ClientPackets.CQuestCommand Then HandleQuestCommand index, Data()
    If PacketNum = ClientPackets.CRequestEditTitle Then HandleRequestEditTitle index, Data()
    If PacketNum = ClientPackets.CSaveTitle Then HandleSaveTitle index, Data()
    If PacketNum = ClientPackets.CRequestTitles Then HandleRequestTitles index, Data()
    If PacketNum = ClientPackets.CSwapTitleSlots Then HandleSwapTitleSlots index, Data()
    If PacketNum = ClientPackets.CTitleCommand Then HandleTitleCommand index, Data()
    If PacketNum = ClientPackets.CRequestNewChar Then HandleRequestNewChar index, Data()
    If PacketNum = ClientPackets.CRequestDelChar Then HandleRequestDelChar index, Data()
    If PacketNum = ClientPackets.CRequestUseChar Then HandleRequestUseChar index, Data()
    If PacketNum = ClientPackets.CCreateChat Then HandleChatCreate index, Data()
    If PacketNum = ClientPackets.CJoinChat Then HandleChatJoin index, Data()
    If PacketNum = ClientPackets.CWhoChat Then HandleChatWho index, Data()
    If PacketNum = ClientPackets.CChatMsg Then HandleChatMsg index, Data()
    If PacketNum = ClientPackets.CLeaveChat Then HandleChatLeave index, Data()

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDataPackets", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim MsgType As Long
        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgType = Buffer.ReadLong

    If MsgType < 0 Or MsgType >= CMSG_COUNT Then
        Exit Sub
    End If
    
    HandleDataPackets MsgType, index, Buffer.ReadBytes(Buffer.Length)

    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString
            
            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Len(Trim$(Name)) > ACCOUNT_LENGTH Or Len(Trim$(Password)) > NAME_LENGTH Then
                Call AlertMsg(index, "Your account name must be between 3 and 12 characters long. Your password must be between 3 and 20 characters long.")
                Exit Sub
            End If

            ' Prevent hacking
            For i = 1 To Len(Name)
                n = AscW(Mid$(Name, i, 1))

                If Not isNameLegal(n) Then
                    Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If

            Next i

            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(index, Name, Password)
                Call TextAdd("Account " & Name & " has been created.")
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                
                ' Load the player
                Call LoadPlayer(index, Name)

                ' send new char shit
                Call SendClasses(index)
                Call SendCharactersData(index)
                        
                ' Show the player up on the socket status
                Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
                Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            Else
                Call AlertMsg(index, "Sorry, that account name is already taken!")
            End If
            
            Set Buffer = Nothing
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNewAccount", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Prevent hacking
            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "The name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            ' Delete names from master name file
            Call LoadPlayer(index, Name)

            If LenB(Trim$(Player(index).Char(TempPlayer(index).Char).Name)) > 0 Then
                Call DeleteName(Player(index).Char(TempPlayer(index).Char).Name)
            End If

            Call ClearPlayer(index)
            ' Everything went ok
            Call Kill(App.Path & "\data\Accounts\" & Trim$(Name) & ".bin")
            Call AddLog("Account " & Trim$(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(index, "Your account has been deleted.")
            
            Set Buffer = Nothing
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDelAccount", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Password As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsPlaying(index) Then
        If Not IsLoggedIn(index) Then
            Set Buffer = New clsBuffer
            Buffer.WriteBytes Data()
            ' Get the data
            Name = Buffer.ReadString
            Password = Buffer.ReadString

            ' Check versions
            If Buffer.ReadLong < CLIENT_MAJOR Or Buffer.ReadLong < CLIENT_MINOR Or Buffer.ReadLong < CLIENT_REVISION Then
                Call AlertMsg(index, "Version outdated, please visit " & Options.Website)
                Exit Sub
            End If

            If isShuttingDown Then
                Call AlertMsg(index, "Server is either rebooting or being shutdown.")
                Exit Sub
            End If

            If Len(Trim$(Name)) < 3 Or Len(Trim$(Password)) < 3 Then
                Call AlertMsg(index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If

            If Not AccountExist(Name) Then
                Call AlertMsg(index, "That account name does not exist.")
                Exit Sub
            End If

            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(index, "Incorrect password.")
                Exit Sub
            End If

            If IsMultiAccounts(Name) Then
                Call AlertMsg(index, "Multiple account logins is not authorized.")
                Exit Sub
            End If

            ' Load the player
            Call LoadPlayer(index, Name)
            ClearBank index
            LoadBank index, Name
            
            ' send new char shit
            Call SendClasses(index)
            Call SendCharactersData(index)
                    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", PLAYER_LOG)
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
            Set Buffer = Nothing
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleLogin", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim Sex As Long
    Dim Class As Long
    Dim Sprite As Long
    Dim i As Long
    Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not IsPlaying(index) Then
        Set Buffer = New clsBuffer
        Buffer.WriteBytes Data()
        Name = Buffer.ReadString
        Sex = Buffer.ReadLong
        Class = Buffer.ReadLong
        Sprite = Buffer.ReadLong

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Then
            Call AlertMsg(index, "Character name must be at least three characters in length.")
            Exit Sub
        End If

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))

            If Not isNameLegal(n) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                Exit Sub
            End If
        Next i

        ' Prevent hacking
        If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
            Exit Sub
        End If

        ' Prevent hacking
        If Class < 1 Or Class > Max_Classes Then
            Exit Sub
        End If

        ' Check if char already exists in slot
        If CharExist(index) Then
            Call AlertMsg(index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "Sorry, but that name is in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Sex, Class, Sprite)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "' s account.", PLAYER_LOG)
        
        ' log them in!!
        JoinGame index
        Set Buffer = Nothing
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAddChar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)
        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = vbNullString
                End If
            End If
        End If
    Next i

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, ' " & Msg & "' ", PLAYER_LOG)
    Call SayMsg_Map(GetPlayerMap(index), index, Msg, QBColor(White))
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte)
    Dim Msg As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next i

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, PLAYER_LOG)
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Right$(Msg, Len(Msg) - 1), EmoteColor)
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleEmoteMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBroadcastMsg(ByVal index As Long, ByRef Data() As Byte)
    Dim Msg As String
    Dim s As String
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next i

    s = "[Global]" & GetPlayerName(index) & ": " & Msg
    Call SayMsg_Global(index, Msg, QBColor(White))
    Call AddLog(s, PLAYER_LOG)
    Call TextAdd(s)
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBroadcastMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandlePlayerMsg(ByVal index As Long, ByRef Data() As Byte)
    Dim Msg As String
    Dim i As Long
    Dim MsgTo As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    MsgTo = FindPlayer(Buffer.ReadString)
    Msg = Buffer.ReadString

    ' Prevent hacking
    For i = 1 To Len(Msg)

        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If

    Next i

    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        If MsgTo > 0 Then
            Call AddLog(GetPlayerName(index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "' ", PLAYER_LOG)
            Call PlayerMsg(MsgTo, GetPlayerName(index) & " tells you, ' " & Msg & "' ", TellColor)
            Call PlayerMsg(index, "You tell " & GetPlayerName(MsgTo) & ", ' " & Msg & "' ", TellColor)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Cannot message yourself.", BrightRed)
    End If
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte)
    Dim Dir As Long
    Dim Movement As Long
    Dim Buffer As clsBuffer
    Dim tmpX As Long, tmpY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong ' CLng(Parse(1))
    Movement = Buffer.ReadLong ' CLng(Parse(2))
    tmpX = Buffer.ReadLong
    tmpY = Buffer.ReadLong
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    ' Prevent hacking
    If Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    ' Prevent player from moving if they have casted a Spell
    If TempPlayer(index).SpellBuffer.Spell > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Cant move if in the bank!
    If TempPlayer(index).InBank Then
        TempPlayer(index).InBank = False
    End If

    ' if stunned, stop them moving
    If TempPlayer(index).StunDuration > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If
    
    ' Prever player from moving if in shop
    If TempPlayer(index).InShop > 0 Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    ' Desynced
    If GetPlayerX(index) <> tmpX Then
        SendPlayerXY (index)
        Exit Sub
    End If

    If GetPlayerY(index) <> tmpY Then
        SendPlayerXY (index)
        Exit Sub
    End If

    Call PlayerMove(index, Dir, Movement)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte)
    Dim Dir As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then
        Exit Sub
    End If

    Dir = Buffer.ReadLong ' CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerDir
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte)
    Dim invNum As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' get inventory slot number
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong
    Set Buffer = Nothing

    UseItem index, invNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUseItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long
    Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' can' t attack whilst casting
    If TempPlayer(index).SpellBuffer.Spell > 0 Then Exit Sub
    
    ' can' t attack whilst stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    SendAttack index

    ' Try to attack a player
    For i = 1 To Player_HighIndex
        ' Make sure we dont try to attack ourselves
        If i <> index Then
            TryPlayerAttackPlayer index, i
        End If
    Next

    ' Try to attack a npc
    For i = 1 To MAX_MAP_NPCS
        TryPlayerAttackNpc index, i
    Next
    
    ' Check tradeskills
    Call GetPlayerNextTile(index, GetPlayerDir(index), x, y)

    ' Prevent subscript out range
    If GetPlayerY(index) = 0 Or GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
    If GetPlayerX(index) = 0 Or GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub

    ' Verifies that has some resource in front of the attacker
    CheckResource index, x, y
    
    ' Attack sound
    If GetPlayerEquipment(index, Weapon) > 0 Then SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, Weapon)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte)
    Dim PointType As Byte
    Dim Buffer As clsBuffer
    Dim sMes As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PointType = Buffer.ReadByte ' CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_Count) Then Exit Sub

    ' Make sure they have points
    If GetPlayerPOINTS(index) <= 0 Then Exit Sub
    
    ' make sure they' re not maxed#
    If GetPlayerRawStat(index, PointType) >= 255 Then
        PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
        Exit Sub
    End If
        
    ' Take away a stat point
    Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)

    ' Everything is ok
    Select Case PointType
        Case Stats.Strength
            Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
            sMes = "Strength"
        Case Stats.Endurance
            Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
            sMes = "Endurance"
        Case Stats.Intelligence
            Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
            sMes = "Intelligence"
        Case Stats.Agility
            Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
            sMes = "Agility"
        Case Stats.Willpower
            Call SetPlayerStat(index, Stats.Willpower, GetPlayerRawStat(index, Stats.Willpower) + 1)
            sMes = "Willpower"
    End Select
        
    SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    ' Send the update
    ' Call SendStats(Index)
    CheckTitle index
    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUseStatPoint", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) ' Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", White)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleWarpMeTo", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player
    n = FindPlayer(Buffer.ReadString) ' Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp yourself to yourself!", White)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleWarpToMe", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The map
    n = Buffer.ReadLong ' CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If n < 0 Or n > MAX_MAPS Then
        Exit Sub
    End If

    Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleWarpTo", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The sprite
    n = Buffer.ReadLong ' CLng(Parse(1))
    Set Buffer = Nothing
    Call SetPlayerSprite(index, n)
    Call SendPlayerData(index)
    Exit Sub

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSetSprite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte)
    Dim Dir As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Dir = Buffer.ReadLong ' CLng(Parse(1))
    Set Buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Sub
    End If

    Call PlayerMove(index, Dir, 1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestNewMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long, n As Long
    Dim MapNum As Long
    Dim x As Long
    Dim y As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    MapNum = GetPlayerMap(index)
    i = Map(MapNum).Revision + 1
    Map(MapNum).Name = Buffer.ReadString
    Map(MapNum).Music = Buffer.ReadByte
    Map(MapNum).Revision = i
    Map(MapNum).Moral = Buffer.ReadByte
    Map(MapNum).Panorama = Buffer.ReadLong
    Map(MapNum).Red = Buffer.ReadByte
    Map(MapNum).Green = Buffer.ReadByte
    Map(MapNum).Blue = Buffer.ReadByte
    Map(MapNum).Alpha = Buffer.ReadByte
    Map(MapNum).Fog = Buffer.ReadByte
    Map(MapNum).FogSpeed = Buffer.ReadByte
    Map(MapNum).FogOpacity = Buffer.ReadByte
    Map(MapNum).Up = Buffer.ReadLong
    Map(MapNum).Down = Buffer.ReadLong
    Map(MapNum).Left = Buffer.ReadLong
    Map(MapNum).Right = Buffer.ReadLong
    Map(MapNum).UpLeft = Buffer.ReadLong
    Map(MapNum).UpRight = Buffer.ReadLong
    Map(MapNum).DownLeft = Buffer.ReadLong
    Map(MapNum).DownRight = Buffer.ReadLong
    Map(MapNum).BootMap = Buffer.ReadLong
    Map(MapNum).BootX = Buffer.ReadByte
    Map(MapNum).BootY = Buffer.ReadByte
    Map(MapNum).MaxX = Buffer.ReadByte
    Map(MapNum).MaxY = Buffer.ReadByte
    ReDim Map(MapNum).Tile(Map(MapNum).MaxX, Map(MapNum).MaxY)

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                For n = 1 To MAX_MAP_LAYERS
                    Map(MapNum).Tile(x, y).Layer(i, n).x = Buffer.ReadLong
                    Map(MapNum).Tile(x, y).Layer(i, n).y = Buffer.ReadLong
                    Map(MapNum).Tile(x, y).Layer(i, n).Tileset = Buffer.ReadLong
                Next
            Next
            Map(MapNum).Tile(x, y).Type = Buffer.ReadByte
            Map(MapNum).Tile(x, y).Data1 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data2 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).Data3 = Buffer.ReadLong
            Map(MapNum).Tile(x, y).DirBlock = Buffer.ReadByte
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Map(MapNum).Npc(x) = Buffer.ReadLong
        Call ClearMapNpc(x, MapNum)
    Next

    ' Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i, MapNum)
    Next

    ' Respawn
    Call SpawnMapItems(MapNum)
    
    ' Map caches
    Call CacheResources(MapNum)
    Call CacheDoors(MapNum)
    Call MapCache_Create(MapNum)

    ' Save the map
    Call SaveMap(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i))
        End If
    Next i

    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte)
    Dim s As String
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Get yes/no value
    s = Buffer.ReadLong ' Parse(1)
    Set Buffer = Nothing

    ' Check if map data is needed to be sent
    If s = 1 Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    ' send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next

    ' send door cache
    For i = 0 To DoorCache(GetPlayerMap(index)).Count
        SendDoorCacheTo index, i
    Next
    
    TempPlayer(index).GettingMap = NO
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SMapDone
    SendDataTo index, Buffer.ToArray()

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleNeedMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up Soundething packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call PlayerMapGetItem(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapGetItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop Soundething packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte)
    Dim invNum As Long
    Dim amount As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteBytes Data()
    invNum = Buffer.ReadLong ' CLng(Parse(1))
    amount = Buffer.ReadLong ' CLng(Parse(2))
    Set Buffer = Nothing
    
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_INV Then Exit Sub
    
    If GetPlayerInvItemNum(index, invNum) < 1 Or GetPlayerInvItemNum(index, invNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, invNum)).Type = iCurrency Then
        If amount < 1 Or amount > GetPlayerInvItemValue(index, invNum) Then Exit Sub
    End If
    
    ' everything worked out fine
    Call PlayerMapDropItem(index, invNum, amount)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapDropItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next

    Call PlayerMsg(index, "Map respawned.", Blue)
    Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapRespawn", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Map report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapReport
    
    For i = 1 To MAX_MAPS
        Buffer.WriteString Trim$(Map(i).Name)
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleMapReport", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) <= 0 Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) ' Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Game_Name & " by " & GetPlayerName(index) & "!", White)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", ADMIN_LOG)
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot kick yourself!", White)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKickPlayer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::
' :: Ban list packet ::
' :::::::::::::::::::::
Sub HandleBanList(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim F As Long
    Dim s As String
    Dim Name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    n = 1
    F = FreeFile
    Open App.Path & "\data\banlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s
        Input #F, Name
        Call PlayerMsg(index, n & ": Banned IP " & s & " by " & Name, White)
        n = n + 1
    Loop

    Close #F

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBanList", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::
' :: Ban destroy packet ::
' ::::::::::::::::::::::::
Sub HandleBanDestroy(ByVal index As Long, ByRef Data() As Byte)
    Dim filename As String
    Dim F As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    filename = App.Path & "\data\banlist.txt"

    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    Kill filename
    Call PlayerMsg(index, "Ban list destroyed.", White)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBanDestroy", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    ' The player index
    n = FindPlayer(Buffer.ReadString) ' Parse(1))
    Set Buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call BanIndex(n, index)
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If

        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "You cannot ban yourself!", White)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBanPlayer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SEditMap
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SItemEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong ' CLng(Parse(1))

    If n < 0 Or n > MAX_ITEMS Then
        Exit Sub
    End If

    ' Update the item
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
    
    ' Save it
    Call SendUpdateItemToAll(n)
    Call SaveItem(n)
    Call AddLog(GetPlayerName(index) & " saved item #" & n & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit Animation packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SAnimationEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::
' :: Save Animation packet ::
' ::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long, i As Byte
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong ' CLng(Parse(1))

    If n < 0 Or n > MAX_ANIMATIONS Then
        Exit Sub
    End If
    
    Animation(n).Name = Buffer.ReadString
    Animation(n).Sound = Buffer.ReadByte
    
    For i = 0 To 1
        Animation(n).Sprite(i) = Buffer.ReadLong
        Animation(n).Frames(i) = Buffer.ReadLong
        Animation(n).LoopCount(i) = Buffer.ReadLong
        Animation(n).LoopTime(i) = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(index) & " saved Animation #" & n & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit npc packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::
' :: Save npc packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte)
    Dim NpcNum As Long
    Dim Buffer As clsBuffer
    Dim i As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    NpcNum = Buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 0 Or NpcNum > MAX_NPCS Then
        Exit Sub
    End If

    Npc(NpcNum).Name = Buffer.ReadString
    Npc(NpcNum).AttackSay = Buffer.ReadString
    Npc(NpcNum).Sound = Buffer.ReadByte
    Npc(NpcNum).Sprite = Buffer.ReadLong
    Npc(NpcNum).SpawnSecs = Buffer.ReadLong
    Npc(NpcNum).Behaviour = Buffer.ReadByte
    Npc(NpcNum).Range = Buffer.ReadByte
    Npc(NpcNum).HP = Buffer.ReadLong
    Npc(NpcNum).EXP = Buffer.ReadLong
    Npc(NpcNum).Animation = Buffer.ReadLong
    Npc(NpcNum).Damage = Buffer.ReadLong
    Npc(NpcNum).Level = Buffer.ReadLong
    Npc(NpcNum).ShopNum = Buffer.ReadLong
    
    For i = 1 To MAX_NPC_DROPS
        Npc(NpcNum).DropChance(i) = Buffer.ReadByte
        Npc(NpcNum).DropItem(i) = Buffer.ReadInteger
        Npc(NpcNum).DropItemValue(i) = Buffer.ReadLong
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Npc(NpcNum).Stat(i) = Buffer.ReadByte
    Next
    
    For i = 1 To MAX_NPC_QUESTS
        Npc(NpcNum).Quest(i) = Buffer.ReadInteger
    Next
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(index) & " saved Npc #" & NpcNum & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Resource packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SResourceEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::
' :: Save Resource packet ::
' :::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte)
    Dim ResourceNum As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ResourceNum = Buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 0 Or ResourceNum > MAX_RESOURCES Then
        Exit Sub
    End If
    
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
    
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SShopEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte)
    Dim ShopNum As Long
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    ShopNum = Buffer.ReadLong

    ' Prevent hacking
    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then Exit Sub

    Shop(ShopNum).Name = Buffer.ReadString
    Shop(ShopNum).BuyRate = Buffer.ReadLong
    
    For i = 1 To MAX_TRADES
        Shop(ShopNum).TradeItem(i).Item = Buffer.ReadLong
        Shop(ShopNum).TradeItem(i).ItemValue = Buffer.ReadLong
        Shop(ShopNum).TradeItem(i).CostItem = Buffer.ReadLong
        Shop(ShopNum).TradeItem(i).CostValue = Buffer.ReadLong
    Next
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit Spell packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditSpell(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSpellEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Save Spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte)
    Dim SpellNum As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    SpellNum = Buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 0 Or SpellNum > MAX_SPELLS Then
        Exit Sub
    End If

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
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(index) & " saved Spell #" & SpellNum & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Exit Sub
    End If

    ' The index
    n = FindPlayer(Buffer.ReadString) ' Parse(1))
    ' The access
    i = Buffer.ReadLong ' CLng(Parse(2))
    Set Buffer = Nothing

    ' Check for invalid access level
    If i >= 0 Or i <= 3 Then

        ' Check if player is on
        If n > 0 Then

            ' check to see if same level access is trying to change another access of the very same level and boot them if they are.
            If GetPlayerAccess(n) = GetPlayerAccess(index) Then
                Call PlayerMsg(index, "Invalid access level.", Red)
                Exit Sub
            End If

            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If

            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "' s access.", ADMIN_LOG)
        Else
            Call PlayerMsg(index, "Player is not online.", White)
        End If

    Else
        Call PlayerMsg(index, "Invalid access level.", Red)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSetAccess", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::::
' :: Who online packet ::
' :::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendWhosOnline(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleWhosOnline", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMotd(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Exit Sub
    End If

    Options.MOTD = Trim$(Buffer.ReadString) ' Parse(1))
    SaveOptions
    Set Buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSetMotd", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte)
    Dim MapNum As Integer
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    x = Buffer.ReadLong ' CLng(Parse(1))
    y = Buffer.ReadLong ' CLng(Parse(2))
    Set Buffer = Nothing

    ' player map
    MapNum = GetPlayerMap(index)
    
    ' Prevent subscript out of range
    If x < 0 Or x > Map(MapNum).MaxX Or y < 0 Or y > Map(MapNum).MaxY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If MapNum = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        ' Change target
                        If TempPlayer(index).targetType = TARGET_TYPE_PLAYER And TempPlayer(index).target = i Then
                            TempPlayer(index).target = 0
                            TempPlayer(index).targetType = TARGET_TYPE_NONE
                            ' send target to player
                            SendTarget index
                        Else
                            TempPlayer(index).target = i
                            TempPlayer(index).targetType = TARGET_TYPE_PLAYER
                            ' send target to player
                            SendTarget index
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Check for an npc
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(MapNum).Npc(i).Num > 0 Then
            If MapNpc(MapNum).Npc(i).x = x Then
                If MapNpc(MapNum).Npc(i).y = y Then
                    If TempPlayer(index).target = i And TempPlayer(index).targetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).target = 0
                        TempPlayer(index).targetType = TARGET_TYPE_NONE
                        ' send target to player
                        SendTarget index
                    Else
                        ' Change target
                        TempPlayer(index).target = i
                        TempPlayer(index).targetType = TARGET_TYPE_NPC
                        ' send target to player
                        SendTarget index
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSearch", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendPlayerSpells(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCast(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Spell slot
    n = Buffer.ReadLong ' CLng(Parse(1))
    Set Buffer = Nothing
    ' set the Spell buffer before castin
    Call BufferSpell(index, n)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCast", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleQuit(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call CloseSocket(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuit", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchInvSlots index, oldSlot, newSlot

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwapInvSlots", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long, n As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If TempPlayer(index).InTrade > 0 Or TempPlayer(index).InBank Or TempPlayer(index).InShop Then Exit Sub
    
    If TempPlayer(index).SpellBuffer.Spell > 0 Then
        PlayerMsg index, "You cannot swap Spells whilst casting.", BrightRed
        Exit Sub
    End If
    
    For n = 1 To MAX_PLAYER_SPELLS
        If TempPlayer(index).SpellCD(n) > GetTickCount Then
            PlayerMsg index, "You cannot swap Spells whilst they' re cooling down.", BrightRed
            Exit Sub
        End If
    Next
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchSpellSlots index, oldSlot, newSlot

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwapSpellSlots", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSendPing
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    PlayerUnequipItem index, Buffer.ReadLong
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUnequip", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestPlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendItems index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestItems", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendAnimations index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestAnimations", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestNPCS(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendNpcs index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestNPCS", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendResources index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestResources", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequesSpells(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendSpells index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequesSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendShops index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestShops", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim tmpItem As Long
    Dim tmpAmount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    ' item
    tmpItem = Buffer.ReadLong
    tmpAmount = Buffer.ReadLong
        
    If GetPlayerAccess(index) < ADMIN_CREATOR Then Exit Sub
    
    SpawnItem tmpItem, tmpAmount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then Exit Sub

    ' Set player exp
    SetPlayerExp index, GetPlayerNextLevel(index)
    CheckPlayerLevelUp index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestLevelUp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleForgeSpell(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Spellslot As Long
    Dim HotbarSlot As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Spellslot = Buffer.ReadLong
    
    ' Prevent subscript out of range
    If Spellslot < 1 Or Spellslot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If
    
    ' dont let them forget a Spell which is in CD
    If TempPlayer(index).SpellCD(Spellslot) > GetTickCount Then
        PlayerMsg index, "Cannot forget a Spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' dont let them forget a Spell which is buffered
    If TempPlayer(index).SpellBuffer.Spell = Spellslot Then
        PlayerMsg index, "Cannot forget a Spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    ' Remove item of hotbar
    HotbarSlot = FindHotbar(index, GetPlayerSpell(index, Spellslot), 2)
    If HotbarSlot > 0 Then
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).SType = 0
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).Slot = 0
        SendHotbar index
    End If
    
    Call SetPlayerSpell(index, Spellslot, 0)
    SendPlayerSpells index
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleForgeSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TempPlayer(index).InShop = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim shopslot As Long
    Dim ShopNum As Long
    Dim itemamount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    shopslot = Buffer.ReadLong
    
    ' not in shop, exit out
    ShopNum = TempPlayer(index).InShop
    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then Exit Sub
    
    With Shop(ShopNum).TradeItem(shopslot)
        ' check trade exists
        If .Item < 1 Then Exit Sub
            
        ' check has the cost item
        itemamount = HasItem(index, .CostItem)
        If itemamount = 0 Or itemamount < .CostValue Then
            PlayerMsg index, "You do not have enough to buy this item.", BrightRed
            ResetShopAction index
            Exit Sub
        End If

        ' check if player have enough room in your inventory
        If FindOpenInvSlot(index, .Item) = 0 Then
            Call PlayerMsg(index, "You don' t have enough room in your inventory!", BrightRed)
            Exit Sub
        End If
                
        ' it' s fine, let' s go ahead
        TakeInvItem index, .CostItem, .CostValue
        GiveInvItem index, .Item, .ItemValue
    End With
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleBuyItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim ItemNum As Long
    Dim Price As Long
    Dim multiplier As Double

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    
    ' if invalid, exit out
    If invSlot < 1 Or invSlot > MAX_INV Then Exit Sub
    
    ' has item?
    If GetPlayerInvItemNum(index, invSlot) < 1 Or GetPlayerInvItemNum(index, invSlot) > MAX_ITEMS Then Exit Sub
    
    ' seems to be valid
    ItemNum = GetPlayerInvItemNum(index, invSlot)
    
    ' work out price
    multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
    Price = Item(ItemNum).Price * multiplier
    
    ' item has cost?
    If Price <= 0 Then
        PlayerMsg index, "The shop doesn' t want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' take item and give gold
    TakeInvItem index, ItemNum, 1
    GiveInvItem index, 1, Price
    
    ' send confirmation message & reset their shop action
    PlayerMsg index, "Trade successful.", BrightGreen
    ResetShopAction index
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSellItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleChangeBankSlots(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim newSlot As Long
    Dim oldSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    
    PlayerSwitchBankSlots index, oldSlot, newSlot
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleChangeBankSlots", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim BankSlot As Long
    Dim amount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    BankSlot = Buffer.ReadLong
    amount = Buffer.ReadLong
    
    TakeBankItem index, BankSlot, amount
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleWithdrawItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    amount = Buffer.ReadLong
    
    GiveBankItem index, invSlot, amount
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDepositItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SaveBank index
    SavePlayer index
    
    TempPlayer(index).InBank = False
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCloseBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim x As Long
    Dim y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    x = Buffer.ReadLong
    y = Buffer.ReadLong
    
    If GetPlayerAccess(index) >= ADMIN_MAPPER Then
        ' PlayerWarp index, GetPlayerMap(index), x, y
        SetPlayerX index, x
        SetPlayerY index, y
        SendPlayerXYToMap index
    End If
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAdminWarp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte)
    Dim tradeTarget As Long, sX As Long, sY As Long, tX As Long, tY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' can' t trade npcs
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub

    ' find the target
    tradeTarget = TempPlayer(index).target
    
    ' make sure we don' t error
    If tradeTarget <= 0 Or tradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' can' t trade with yourself..
    If tradeTarget = index Then
        PlayerMsg index, "You can' t trade with yourself.", BrightRed
        Exit Sub
    End If
    
    ' make sure they' re on the same map
    If Not GetPlayerMap(tradeTarget) = GetPlayerMap(index) Then Exit Sub
    
    ' make sure they' re stood next to each other
    tX = GetPlayerX(tradeTarget)
    tY = GetPlayerY(tradeTarget)
    sX = GetPlayerX(index)
    sY = GetPlayerY(index)
    
    ' within range?
    If tX < sX - 1 Or tX > sX + 1 Then
        PlayerMsg index, "You need to be standing next to Soundeone to request a trade.", BrightRed
        Exit Sub
    End If
    If tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to Soundeone to request a trade.", BrightRed
        Exit Sub
    End If
    
    ' make sure not already got a trade request
    If TempPlayer(tradeTarget).TradeRequest > 0 Then
        PlayerMsg index, "This player is busy.", BrightRed
        Exit Sub
    End If

    ' send the trade request
    TempPlayer(tradeTarget).TradeRequest = index
    SendDialogue tradeTarget, "Trade Request", GetPlayerName(index) & " has requested a trade. Would you like to accept?", DIALOGUE_TYPE_TRADE, YES

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeRequest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte)
    Dim tradeTarget As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If TempPlayer(index).InTrade > 0 Then
        TempPlayer(index).TradeRequest = 0
    Else
        tradeTarget = TempPlayer(index).TradeRequest
        ' let them know they' re trading
        PlayerMsg index, "You have accepted " & GetPlayerName(tradeTarget) & "' s trade request.", BrightGreen
        PlayerMsg tradeTarget, GetPlayerName(index) & " has accepted your trade request.", BrightGreen
        ' clear the tradeRequest server-side
        TempPlayer(index).TradeRequest = 0
        TempPlayer(tradeTarget).TradeRequest = 0
        ' set that they' re trading with each other
        TempPlayer(index).InTrade = tradeTarget
        TempPlayer(tradeTarget).InTrade = index
        ' clear out their trade offers
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        ' Used to init the trade window clientside
        SendTrade index, tradeTarget
        SendTrade tradeTarget, index
        ' Send the offer data - Used to clear their client
        SendTradeUpdate index, 0
        SendTradeUpdate index, 1
        SendTradeUpdate tradeTarget, 0
        SendTradeUpdate tradeTarget, 1
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAcceptTradeRequest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    PlayerMsg TempPlayer(index).TradeRequest, GetPlayerName(index) & " has declined your trade request.", BrightRed
    PlayerMsg index, "You decline the trade request.", BrightRed
    ' clear the tradeRequest server-side
    TempPlayer(index).TradeRequest = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDeclineTradeRequest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte)
    Dim tradeTarget As Long
    Dim i As Long
    Dim tmpTradeItem(MAX_INV) As PlayerInvRec
    Dim tmpTradeItem2(MAX_INV) As PlayerInvRec
    Dim ItemNum As Long
        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    tradeTarget = TempPlayer(index).InTrade
    
    If tradeTarget > 0 Then
        If GetPlayerMap(index) <> GetPlayerMap(TempPlayer(index).InTrade) Then Exit Sub

        TempPlayer(index).AcceptTrade = True

        ' if not both of them accept, then exit
        If Not TempPlayer(tradeTarget).AcceptTrade Then
            SendTradeStatus index, 2
            SendTradeStatus tradeTarget, 1
            Exit Sub
        End If
    
        ' if not have space in inventory of tradetarget
        If IsInventoryFull(tradeTarget, index) Then
            TempPlayer(index).InTrade = 0
            TempPlayer(tradeTarget).InTrade = 0
            TempPlayer(index).AcceptTrade = False
            TempPlayer(tradeTarget).AcceptTrade = False
            PlayerMsg tradeTarget, "You do not have enough space in inventory.", BrightRed
            PlayerMsg index, GetPlayerName(tradeTarget) & " do not have enough space in inventory.", BrightRed
            SendCloseTrade index
            SendCloseTrade tradeTarget
            Exit Sub
        End If
        
        ' take their items
        For i = 1 To MAX_INV
            ' player
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ItemNum = GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)
                If ItemNum > 0 Then
                    ' store temp
                    tmpTradeItem(i).Num = ItemNum
                    tmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot index, TempPlayer(index).TradeOffer(i).Num, tmpTradeItem(i).Value
                End If
            End If
            ' target
            If TempPlayer(tradeTarget).TradeOffer(i).Num > 0 Then
                ItemNum = GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
                If ItemNum > 0 Then
                    ' store temp
                    tmpTradeItem2(i).Num = ItemNum
                    tmpTradeItem2(i).Value = TempPlayer(tradeTarget).TradeOffer(i).Value
                    ' take item
                    TakeInvSlot tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num, tmpTradeItem2(i).Value
                End If
            End If
        Next
        
        ' taken all items. now they can' t not get items because of no inventory space.
        For i = 1 To MAX_INV
            ' player
            If tmpTradeItem2(i).Num > 0 Then
                ' give away!
                GiveInvItem index, tmpTradeItem2(i).Num, tmpTradeItem2(i).Value, False
            End If
            ' target
            If tmpTradeItem(i).Num > 0 Then
                ' give away!
                GiveInvItem tradeTarget, tmpTradeItem(i).Num, tmpTradeItem(i).Value, False
            End If
        Next

        SendInventory index
        SendInventory tradeTarget
    
        ' they now have all the items. Clear out values + let them out of the trade.
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
        TempPlayer(index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
    
        PlayerMsg index, "Trade completed.", BrightGreen
        PlayerMsg tradeTarget, "Trade completed.", BrightGreen
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAcceptTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long
    Dim tradeTarget As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    tradeTarget = TempPlayer(index).InTrade

    If tradeTarget > 0 Then
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next

        TempPlayer(index).InTrade = 0
        TempPlayer(tradeTarget).InTrade = 0
        TempPlayer(index).AcceptTrade = False
        TempPlayer(tradeTarget).AcceptTrade = False
    
        PlayerMsg index, "You declined the trade.", BrightRed
        PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
    
        SendCloseTrade index
        SendCloseTrade tradeTarget
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDeclineTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim invSlot As Long
    Dim amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    invSlot = Buffer.ReadLong
    amount = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If invSlot <= 0 Or invSlot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(index, invSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' make sure they have the amount they offer
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then
        Exit Sub
    End If
    
    ' check if amount is > 0
    If Item(ItemNum).Type = iCurrency Then
        If amount < 1 Then Exit Sub
    End If

    If Item(ItemNum).Type = iCurrency Then
        ' check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                ' add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + amount
                ' clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, invSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, invSlot)
                End If
                ' cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' exit early
                Exit Sub
            End If
        Next
    Else
        ' make sure they' re not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = invSlot Then
                PlayerMsg index, "You' ve already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    TempPlayer(index).TradeOffer(EmptySlot).Num = invSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = amount
    
    ' cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTradeItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim tradeSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    tradeSlot = Buffer.ReadLong
    
    Set Buffer = Nothing
    
    If tradeSlot <= 0 Or tradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(tradeSlot).Num <= 0 Then Exit Sub
    
    TempPlayer(index).TradeOffer(tradeSlot).Num = 0
    TempPlayer(index).TradeOffer(tradeSlot).Value = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleUntradeItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim SType As Long
    Dim Slot As Long
    Dim hotbarNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    SType = Buffer.ReadLong
    Slot = Buffer.ReadLong
    hotbarNum = Buffer.ReadLong
    
    Select Case SType
        Case 0 ' clear
            Call SetPlayerHotbarSlot(index, hotbarNum, 0)
            Call SetPlayerHotbarType(index, hotbarNum, 0)
        Case 1 ' inventory
            If Slot > 0 And Slot <= MAX_INV Then
                If GetPlayerInvItemNum(index, Slot) > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Call SetPlayerHotbarSlot(index, hotbarNum, GetPlayerInvItemNum(index, Slot))
                        Call SetPlayerHotbarType(index, hotbarNum, SType)
                    End If
                End If
            End If
        Case 2 ' Spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If GetPlayerSpell(index, Slot) > 0 Then
                    If Len(Trim$(Spell(Player(index).Char(TempPlayer(index).Char).Spell(Slot)).Name)) > 0 Then
                        Call SetPlayerHotbarSlot(index, hotbarNum, GetPlayerSpell(index, Slot))
                        Call SetPlayerHotbarType(index, hotbarNum, SType)
                    End If
                End If
            End If
        Case 3 ' title
            If Slot > 0 And Slot <= MAX_PLAYER_TITLES Then
                If GetPlayerTitle(index, Slot) > 0 Then
                    If Len(Trim$(Title(GetPlayerTitle(index, Slot)).Name)) > 0 Then
                        Player(index).Char(TempPlayer(index).Char).Hotbar(hotbarNum).Slot = GetPlayerTitle(index, Slot)
                        Player(index).Char(TempPlayer(index).Char).Hotbar(hotbarNum).SType = SType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbarChange", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleHotbarUse(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Slot As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    
    Slot = Buffer.ReadLong
    
    Select Case GetPlayerHotbarType(index, Slot)
        Case 1 ' inventory
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(index, i) > 0 Then
                    If GetPlayerInvItemNum(index, i) = GetPlayerHotbarSlot(index, Slot) Then
                        UseItem index, i
                        Exit Sub
                    End If
                End If
            Next
        Case 2 ' Spell
            For i = 1 To MAX_PLAYER_SPELLS
                If GetPlayerSpell(index, i) > 0 Then
                    If GetPlayerSpell(index, i) = GetPlayerHotbarSlot(index, Slot) Then
                        BufferSpell index, i
                        Exit Sub
                    End If
                End If
            Next
    End Select
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHotbarUse", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it' s a valid target
    If TempPlayer(index).targetType <> TARGET_TYPE_PLAYER Then Exit Sub
    If TempPlayer(index).target = index Then Exit Sub
    
    ' make sure they' re connected and on the same map
    If Not IsConnected(TempPlayer(index).target) Or Not IsPlaying(TempPlayer(index).target) Then Exit Sub
    If GetPlayerMap(TempPlayer(index).target) <> GetPlayerMap(index) Then Exit Sub
    
    ' init the request
    Party_Invite index, TempPlayer(index).target

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyRequest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsConnected(TempPlayer(index).partyInvite) Or Not IsPlaying(TempPlayer(index).partyInvite) Then
        TempPlayer(index).partyInvite = 0
        Exit Sub
    End If
    
    Party_InviteAccept TempPlayer(index).partyInvite, index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleAcceptParty", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Party_InviteDecline TempPlayer(index).partyInvite, index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDeclineParty", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Party_PlayerLeave index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandlePartyLeave", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestEditDoor(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SDoorEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditDoor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSaveDoor(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long, i As Byte
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong ' CLng(Parse(1))

    If n < 0 Or n > MAX_DOORS Then
        Exit Sub
    End If
    
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
    
    ' Save it
    Call SendUpdateDoorToAll(n)
    Call SaveDoor(n)
    Call AddLog(GetPlayerName(index) & " saved Door #" & n & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveDoor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestDoors(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendDoors index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestDoors", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleCheckDoor(ByVal index As Long, ByRef Data() As Byte)
    Dim x As Byte, y As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case GetPlayerDir(index)
        Case DIR_UP
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_DOWN
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
        Case DIR_UP_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) - 1
        Case DIR_UP_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) - 1
        Case DIR_DOWN_LEFT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) + 1
        Case DIR_DOWN_RIGHT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) + 1
    End Select
    
    OpenDoor index, x, y

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleCheckDoor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestEditQuest(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SQuestEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditQuest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSaveQuest(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long, i As Long, x As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong ' CLng(Parse(1))

    If n < 0 Or n > MAX_QUESTS Then
        Exit Sub
    End If

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
            Quest(n).Task(i).Message(x) = Buffer.ReadString
        Next
    Next
    Set Buffer = Nothing
    
    ' Save it
    Call SendUpdateQuestToAll(n)
    Call SaveQuest(n)
    Call AddLog(GetPlayerName(index) & " saved Quest #" & n & ".", ADMIN_LOG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveQuest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestQuests(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendQuests index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestQuests", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleQuestCommand(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer, Command As Byte, Value As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadByte
    Value = Buffer.ReadLong
    Set Buffer = Nothing
    
    Select Case Command
            ' Accept quest
        Case 1
            QuestAccept index
            ' Select quest
        Case 2
            CheckQuest index, TempPlayer(index).QuestSelect, Value
    End Select

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleQuestCommand", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestEditTitle(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteLong STitleEditor
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestEditTitle", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSaveTitle(ByVal index As Long, ByRef Data() As Byte)
    Dim n As Long, i As Long
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Exit Sub
    End If

    n = Buffer.ReadLong ' CLng(Parse(1))

    If n < 0 Or n > MAX_TITLES Then
        Exit Sub
    End If

    Title(n).Name = Buffer.ReadString
    Title(n).Description = Buffer.ReadString
    Title(n).Icon = Buffer.ReadInteger
    Title(n).Type = Buffer.ReadByte
    Title(n).color = Buffer.ReadByte
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
    
    ' Save it
    Call SendUpdateTitleToAll(n)
    Call SaveTitle(n)
    Call AddLog(GetPlayerName(index) & " saved Title #" & n & ".", ADMIN_LOG)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSaveTitle", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestTitles(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendTitles index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestTitles", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleSwapTitleSlots(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim oldSlot As Long, newSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    ' Old Slot
    oldSlot = Buffer.ReadLong
    newSlot = Buffer.ReadLong
    Set Buffer = Nothing
    PlayerSwitchTitleSlots index, oldSlot, newSlot
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleSwapTitleSlots", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleTitleCommand(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim Command As Byte, Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Command = Buffer.ReadByte
    Slot = Buffer.ReadLong
    
    Select Case Command
        Case 1: Call UseTitle(index, Slot)
        Case 2: Call RemoveTitle(index, Slot)
        Case 3: Call RemoveTUsing(index)
    End Select
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleTitleCommand", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestNewChar(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If IsPlaying(index) Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TempPlayer(index).Char = Buffer.ReadLong
    Call SendNewCharClasses(index)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestNewChar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestUseChar(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If IsPlaying(index) Then Exit Sub
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TempPlayer(index).Char = Buffer.ReadLong
    
    ' Actually log in
    JoinGame index
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestUseChar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleRequestDelChar(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out range
    If IsPlaying(index) Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    TempPlayer(index).Char = Buffer.ReadLong
    Set Buffer = Nothing

    ' Delete char name of char list
    If CharExist(index) Then
        Call DeleteName(Trim$(Player(index).Char(TempPlayer(index).Char).Name))
    End If

    ' Delete the char.
    ClearChar index, TempPlayer(index).Char
    SendCharactersData index
        
    ' Save the account
    SavePlayer index
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleRequestDelChar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleChatCreate(ByVal index As Long, ByRef Data() As Byte)
    Dim chatName As String
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' checks
    If TempPlayer(index).roomIndex > 0 Then
        PlayerMsg index, "You already are in a chat room. Please leave the current Room and try again.", AlertColor
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    chatName = Buffer.ReadString
    Call CreateChat(chatName, index)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleChatCreate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleChatJoin(ByVal index As Long, ByRef Data() As Byte)
    Dim chatName As String
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' checks
    If TempPlayer(index).roomIndex > 0 Then
        PlayerMsg index, "You already are in a chat room.", AlertColor
        Exit Sub
    End If

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    chatName = Buffer.ReadString
    Call JoinChat(chatName, index)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleChatJoin", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleChatWho(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long
    Dim n As Long
    Dim s As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If TempPlayer(index).roomIndex <= 0 Then
        PlayerMsg index, "You are not in a chat room.", AlertColor
        Exit Sub
    End If

    ' Find all the players that are in chat room
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If TempPlayer(i).roomIndex = TempPlayer(index).roomIndex Then
                    s = s & GetPlayerName(i) & ", "
                    n = n + 1
                End If
            End If
        End If
    Next
    
    ' Send message
    If n = 0 Then
        s = "There are no other players online in" & ChatRoom(TempPlayer(index).roomIndex).Name & "."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players in '" & ChatRoom(TempPlayer(index).roomIndex).Name & "' : " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleChatWho", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleChatMsg(ByVal index As Long, ByRef Data() As Byte)
    Dim Msg As String, Buffer As clsBuffer, s As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If TempPlayer(index).roomIndex <= 0 Then
        PlayerMsg index, "You are not in a chat room.", AlertColor
        Exit Sub
    End If

    ' Prevent hacking
    For i = 1 To Len(Msg)
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then
            Exit Sub
        End If
    Next

    Set Buffer = New clsBuffer
    Buffer.WriteBytes Data()
    Msg = Buffer.ReadString
    s = "[" & ChatRoom(TempPlayer(index).roomIndex).Name & "] " & GetPlayerName(index) & ": " & Msg
    ChatRoomMsg TempPlayer(index).roomIndex, s, Blue
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleChatMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleChatLeave(ByVal index As Long, ByRef Data() As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If TempPlayer(index).roomIndex > 0 Then
        LeaveChat (index)
    Else
        PlayerMsg index, "You are not in a chat room.", AlertColor
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleChatLeave", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
