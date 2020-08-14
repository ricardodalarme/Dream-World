Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    frmServer.Caption = Options.Game_Name & " <IP " & frmServer.Socket(0).LocalIP & " Port " & CStr(frmServer.Socket(0).LocalPort) & "> (" & TotalOnlinePlayers & ")"

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCaption", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsConnected", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Checks if the player is online
    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsPlaying", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Checks if the player is logged
    If IsConnected(index) Then
        If LenB(Trim$(Player(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsLoggedIn", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Checks if the player is already logged
    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsMultiAccounts", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function IsBanned(ByVal IP As String) As Boolean
    Dim filename As String
    Dim fIP As String
    Dim fName As String
    Dim F As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' File path
    filename = App.Path & "\data\banlist.txt"

    ' Check if file exists
    If Not FileExist("data\banlist.txt") Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP
        Input #F, fName

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsBanned", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
    Dim Buffer As clsBuffer
    Dim TempData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
              
        frmServer.Socket(index).SendData Buffer.ToArray()
        Set Buffer = Nothing
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataToAllBut", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToChatRoom(ByVal roomIndex As Long, ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).roomIndex = roomIndex Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataToChatRoom", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If PlayersOnMap(MapNum) Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    Call SendDataTo(i, Data)
                End If
            End If
        Next
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If PlayersOnMap(MapNum) Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If GetPlayerMap(i) = MapNum Then
                    If i <> index Then
                        Call SendDataTo(i, Data)
                    End If
                End If
            End If
        Next
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataToMapBut", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDataToParty(ByVal partyNum As Long, ByRef Data() As Byte)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Party(partyNum).MemberCount > 0 Then
        For i = 1 To Party(partyNum).MemberCount
            If Party(partyNum).Member(i) > 0 Then
                Call SendDataTo(Party(partyNum).Member(i), Data)
            End If
        Next
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDataToParty", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GlobalMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            SendDataTo i, Buffer.ToArray
        End If
    Next
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AdminMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    Buffer.WriteLong ServerPackets.SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub ChatRoomMsg(ByVal roomIndex As Long, ByVal Msg As String, ByVal color As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SChatMsg
    Buffer.WriteString Msg
    Buffer.WriteLong color

    SendDataToChatRoom roomIndex, Buffer.ToArray
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ChatRoomMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SAlertMsg
    Buffer.WriteString Msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    Call CloseSocket(index)
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AlertMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub PartyMsg(ByVal partyNum As Long, ByVal Msg As String, ByVal color As Byte)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' exist?
        If Party(partyNum).Member(i) > 0 Then
            ' make sure they' re logged on
            If IsConnected(Party(partyNum).Member(i)) And IsPlaying(Party(partyNum).Member(i)) Then
                PlayerMsg Party(partyNum).Member(i), Msg, color
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PartyMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If

        Call AlertMsg(index, "You have lost your connection with " & Options.Game_Name & ".")
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HackingAttempt", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If i <> 0 Then
            ' we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AcceptConnection", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SocketConnected(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index <> 0 Then
        ' make sure they' re not banned
        If Not IsBanned(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "You have been banned from " & Options.Game_Name & ", and can no longer play.")
        End If
        ' re-set the high index
        Player_HighIndex = 0
        For i = MAX_PLAYERS To 1 Step -1
            If IsConnected(i) Then
                Player_HighIndex = i
                Exit For
            End If
        Next
        ' send the new highindex to all logged in players
        SendHighIndex
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SocketConnected", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
    Dim Buffer() As Byte
    Dim pLength As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetTickCount >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = GetTickCount + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1500 Then
            Exit Sub
        End If
    
        ' Check for packet flooding
        If TempPlayer(index).DataPackets > 35 Then
            Exit Sub
        End If
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then
            Exit Sub
        End If
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then
                Exit Sub
            End If
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "IncomingData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CloseSocket(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        frmServer.Socket(index).Close
        Call UpdateCaption
        Call ClearPlayer(index)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CloseSocket", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CreateFull_MapCache()
    Dim i As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        MapCache_Create i
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateFull_MapCache", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub MapCache_Create(ByVal MapNum As Long)
    Dim x As Long
    Dim y As Long
    Dim i As Long, n As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteByte Map(MapNum).Music
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteByte Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).Panorama
    Buffer.WriteByte Map(MapNum).Red
    Buffer.WriteByte Map(MapNum).Green
    Buffer.WriteByte Map(MapNum).Blue
    Buffer.WriteByte Map(MapNum).Alpha
    Buffer.WriteByte Map(MapNum).Fog
    Buffer.WriteByte Map(MapNum).FogSpeed
    Buffer.WriteByte Map(MapNum).FogOpacity
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).UpLeft
    Buffer.WriteLong Map(MapNum).UpRight
    Buffer.WriteLong Map(MapNum).DownLeft
    Buffer.WriteLong Map(MapNum).DownRight
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteByte Map(MapNum).BootX
    Buffer.WriteByte Map(MapNum).BootY
    Buffer.WriteByte Map(MapNum).MaxX
    Buffer.WriteByte Map(MapNum).MaxY

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            With Map(MapNum).Tile(x, y)
                For i = 1 To MapLayer.Layer_Count - 1
                    For n = 1 To MAX_MAP_LAYERS
                        Buffer.WriteLong .Layer(i, n).x
                        Buffer.WriteLong .Layer(i, n).y
                        Buffer.WriteLong .Layer(i, n).Tileset
                    Next
                Next
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteByte .DirBlock
            End With

        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).Npc(x)
    Next

    MapCache(MapNum).Data = Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "MapCache_Create", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            If i <> index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If

    Next

    If n = 0 Then
        s = "There are no other players online."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWhosOnline", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerLevel(index)
    Buffer.WriteLong GetPlayerPOINTS(index)
    Buffer.WriteLong GetPlayerSprite(index)
    Buffer.WriteLong GetPlayerMap(index)
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteLong GetPlayerKillPlayers(index)
    Buffer.WriteLong GetPlayerTUsing(index)

    For i = 1 To MAX_PLAYER_TITLES
        Buffer.WriteLong GetPlayerTitle(index, i)
    Next
    
    For i = 1 To MAX_NPCS
        Buffer.WriteInteger GetPlayerKillNpcs(index, i)
    Next

    For i = 1 To MAX_PLAYER_QUESTS
        Buffer.WriteInteger GetPlayerQuestNum(index, i)
        Buffer.WriteByte GetPlayerQuestStatus(index, i)
        Buffer.WriteByte GetPlayerQuestPart(index, i)
    Next

    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Function
errorhandler:
    HandleError "PlayerData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next

    ' Send index' s player data to everyone on the map including himself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendJoinMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SLeft
    Buffer.WriteLong index
    SendDataToMapBut index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLeaveMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerData(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SendDataToMap GetPlayerMap(index), PlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong ServerPackets.SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next

    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapItemsTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapItemsToAll(ByVal MapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SMapItemData

    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteLong MapItem(MapNum, i).x
        Buffer.WriteLong MapItem(MapNum, i).y
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapItemsToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Long, ByVal MapNpcNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SMapNpcVitals
    Buffer.WriteLong MapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Vital(i)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapNpcVitals", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(Vitals.HP)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapNpcsTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SMapNpcData

    For i = 1 To MAX_MAP_NPCS
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Num
        Buffer.WriteLong MapNpc(MapNum).Npc(i).x
        Buffer.WriteLong MapNpc(MapNum).Npc(i).y
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Dir
        Buffer.WriteLong MapNpc(MapNum).Npc(i).Vital(Vitals.HP)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapNpcsToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ITEMS
        If LenB(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendItems", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ANIMATIONS
        If LenB(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAnimations", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_NPCS
        If LenB(Trim$(Npc(i).Name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNpcs", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_RESOURCES
        If LenB(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendResources", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendInventory(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, i)
        Buffer.WriteLong GetPlayerInvItemValue(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendInventory", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal invSlot As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SPlayerInvUpdate
    Buffer.WriteLong invSlot
    Buffer.WriteLong GetPlayerInvItemNum(index, invSlot)
    Buffer.WriteLong GetPlayerInvItemValue(index, invSlot)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendInventoryUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SPlayerWornEq
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWornEquipment", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SMapWornEq
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerEquipment(index, Armor)
    Buffer.WriteLong GetPlayerEquipment(index, Weapon)
    Buffer.WriteLong GetPlayerEquipment(index, Helmet)
    Buffer.WriteLong GetPlayerEquipment(index, Shield)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapEquipment", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer

    Select Case Vital
        Case Vitals.HP
            Buffer.WriteLong ServerPackets.SPlayerHp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case Vitals.MP
            Buffer.WriteLong ServerPackets.SPlayerMp
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataTo index, Buffer.ToArray()

    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendVital", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendEXP(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong ServerPackets.SPlayerEXP
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendEXP", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendStats(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerStats
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong GetPlayerStat(index, i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendStats", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendWelcome(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Send them MOTD
    If LenB(Options.MOTD) > 0 Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If

    ' Send whos online
    Call SendWhosOnline(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendWelcome", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendClasses(ByVal index As Long)
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SClassesData
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)

        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        
        ' send array size
        Buffer.WriteLong n
        
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendClasses", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim i As Long, n As Long, q As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNewCharClasses
    Buffer.WriteLong Max_Classes

    For i = 1 To Max_Classes
        Buffer.WriteString GetClassName(i)

        ' set sprite array size
        n = UBound(Class(i).MaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).MaleSprite(q)
        Next
        
        ' set sprite array size
        n = UBound(Class(i).FemaleSprite)
        ' send array size
        Buffer.WriteLong n
        ' loop around sending each sprite
        For q = 0 To n
            Buffer.WriteLong Class(i).FemaleSprite(q)
        Next
        
        For q = 1 To Stats.Stat_Count - 1
            Buffer.WriteLong Class(i).Stat(q)
        Next
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNewCharClasses", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerData
    Buffer.WriteLong index
    Buffer.WriteString vbNullString
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    Buffer.WriteLong 0
    SendDataToAllBut index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLeftGame", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerXY
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerXY", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerXYToMap(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerXYMap
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerXYToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteString Trim$(Item(ItemNum).Desc)
    Buffer.WriteByte Item(ItemNum).Sound
    Buffer.WriteLong Item(ItemNum).Pic
    Buffer.WriteByte Item(ItemNum).Type
    Buffer.WriteLong Item(ItemNum).Data1
    Buffer.WriteLong Item(ItemNum).Data2
    Buffer.WriteLong Item(ItemNum).Data3
    Buffer.WriteLong Item(ItemNum).ClassReq
    Buffer.WriteLong Item(ItemNum).AccessReq
    Buffer.WriteLong Item(ItemNum).LevelReq
    Buffer.WriteLong Item(ItemNum).Price
    Buffer.WriteByte Item(ItemNum).Rarity
    Buffer.WriteLong Item(ItemNum).Speed
    Buffer.WriteLong Item(ItemNum).Animation
    Buffer.WriteLong Item(ItemNum).Paperdoll
    Buffer.WriteLong Item(ItemNum).AddHP
    Buffer.WriteLong Item(ItemNum).AddMP
    Buffer.WriteLong Item(ItemNum).AddEXP
    Buffer.WriteLong Item(ItemNum).Damage
    Buffer.WriteLong Item(ItemNum).Protection
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Item(ItemNum).Add_Stat(i)
        Buffer.WriteByte Item(ItemNum).Stat_Req(i)
    Next
    
    For i = 1 To MAX_BAG
        Buffer.WriteLong Item(ItemNum).BagItem(i)
        Buffer.WriteLong Item(ItemNum).BagValue(i)
    Next
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateItemToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteString Trim$(Item(ItemNum).Name)
    Buffer.WriteString Trim$(Item(ItemNum).Desc)
    Buffer.WriteByte Item(ItemNum).Sound
    Buffer.WriteLong Item(ItemNum).Pic
    Buffer.WriteByte Item(ItemNum).Type
    Buffer.WriteLong Item(ItemNum).Data1
    Buffer.WriteLong Item(ItemNum).Data2
    Buffer.WriteLong Item(ItemNum).Data3
    Buffer.WriteLong Item(ItemNum).ClassReq
    Buffer.WriteLong Item(ItemNum).AccessReq
    Buffer.WriteLong Item(ItemNum).LevelReq
    Buffer.WriteLong Item(ItemNum).Price
    Buffer.WriteByte Item(ItemNum).Rarity
    Buffer.WriteLong Item(ItemNum).Speed
    Buffer.WriteLong Item(ItemNum).Animation
    Buffer.WriteLong Item(ItemNum).Paperdoll
    Buffer.WriteLong Item(ItemNum).AddHP
    Buffer.WriteLong Item(ItemNum).AddMP
    Buffer.WriteLong Item(ItemNum).AddEXP
    Buffer.WriteLong Item(ItemNum).Damage
    Buffer.WriteLong Item(ItemNum).Protection
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Item(ItemNum).Add_Stat(i)
        Buffer.WriteByte Item(ItemNum).Stat_Req(i)
    Next
    
    For i = 1 To MAX_BAG
        Buffer.WriteLong Item(ItemNum).BagItem(i)
        Buffer.WriteLong Item(ItemNum).BagValue(i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateItemTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteString Trim$(Animation(AnimationNum).Name)
    Buffer.WriteByte Animation(AnimationNum).Sound
    
    For i = 0 To 1
        Buffer.WriteLong Animation(AnimationNum).Sprite(i)
        Buffer.WriteLong Animation(AnimationNum).Frames(i)
        Buffer.WriteLong Animation(AnimationNum).LoopCount(i)
        Buffer.WriteLong Animation(AnimationNum).LoopTime(i)
    Next
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateAnimationToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteString Trim$(Animation(AnimationNum).Name)
    Buffer.WriteByte Animation(AnimationNum).Sound
    
    For i = 0 To 1
        Buffer.WriteLong Animation(AnimationNum).Sprite(i)
        Buffer.WriteLong Animation(AnimationNum).Frames(i)
        Buffer.WriteLong Animation(AnimationNum).LoopCount(i)
        Buffer.WriteLong Animation(AnimationNum).LoopTime(i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateAnimationTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateNpcToAll(ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteString Trim$(Npc(NpcNum).AttackSay)
    Buffer.WriteByte Npc(NpcNum).Sound
    Buffer.WriteLong Npc(NpcNum).Sprite
    Buffer.WriteLong Npc(NpcNum).SpawnSecs
    Buffer.WriteByte Npc(NpcNum).Behaviour
    Buffer.WriteByte Npc(NpcNum).Range
    Buffer.WriteLong Npc(NpcNum).HP
    Buffer.WriteLong Npc(NpcNum).EXP
    Buffer.WriteLong Npc(NpcNum).Animation
    Buffer.WriteLong Npc(NpcNum).Damage
    Buffer.WriteLong Npc(NpcNum).Level
    Buffer.WriteLong Npc(NpcNum).ShopNum
    
    For i = 1 To MAX_NPC_DROPS
        Buffer.WriteByte Npc(NpcNum).DropChance(i)
        Buffer.WriteInteger Npc(NpcNum).DropItem(i)
        Buffer.WriteLong Npc(NpcNum).DropItemValue(i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Npc(NpcNum).Stat(i)
    Next
    
    For i = 1 To MAX_NPC_QUESTS
        Buffer.WriteInteger Npc(NpcNum).Quest(i)
    Next
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateNpcToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal NpcNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer

    Buffer.WriteLong ServerPackets.SUpdateNpc
    Buffer.WriteLong NpcNum
    Buffer.WriteString Trim$(Npc(NpcNum).Name)
    Buffer.WriteString Trim$(Npc(NpcNum).AttackSay)
    Buffer.WriteByte Npc(NpcNum).Sound
    Buffer.WriteLong Npc(NpcNum).Sprite
    Buffer.WriteLong Npc(NpcNum).SpawnSecs
    Buffer.WriteByte Npc(NpcNum).Behaviour
    Buffer.WriteByte Npc(NpcNum).Range
    Buffer.WriteLong Npc(NpcNum).HP
    Buffer.WriteLong Npc(NpcNum).EXP
    Buffer.WriteLong Npc(NpcNum).Animation
    Buffer.WriteLong Npc(NpcNum).Damage
    Buffer.WriteLong Npc(NpcNum).Level
    Buffer.WriteLong Npc(NpcNum).ShopNum
    
    For i = 1 To MAX_NPC_DROPS
        Buffer.WriteByte Npc(NpcNum).DropChance(i)
        Buffer.WriteInteger Npc(NpcNum).DropItem(i)
        Buffer.WriteLong Npc(NpcNum).DropItemValue(i)
    Next
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Npc(NpcNum).Stat(i)
    Next
    
    For i = 1 To MAX_NPC_QUESTS
        Buffer.WriteInteger Npc(NpcNum).Quest(i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateNpcTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteString Trim$(Resource(ResourceNum).Name)
    Buffer.WriteString Trim$(Resource(ResourceNum).SuccessMessage)
    Buffer.WriteString Trim$(Resource(ResourceNum).EmptyMessage)
    Buffer.WriteByte Resource(ResourceNum).Sound
    Buffer.WriteByte Resource(ResourceNum).Type
    Buffer.WriteLong Resource(ResourceNum).ResourceImage
    Buffer.WriteLong Resource(ResourceNum).ItemReward
    Buffer.WriteLong Resource(ResourceNum).ToolRequired
    Buffer.WriteLong Resource(ResourceNum).Health
    Buffer.WriteLong Resource(ResourceNum).RespawnTime
    Buffer.WriteLong Resource(ResourceNum).ToolRequired
    Buffer.WriteLong Resource(ResourceNum).Animation
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateResourceToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteString Trim$(Resource(ResourceNum).Name)
    Buffer.WriteString Trim$(Resource(ResourceNum).SuccessMessage)
    Buffer.WriteString Trim$(Resource(ResourceNum).EmptyMessage)
    Buffer.WriteByte Resource(ResourceNum).Sound
    Buffer.WriteByte Resource(ResourceNum).Type
    Buffer.WriteLong Resource(ResourceNum).ResourceImage
    Buffer.WriteLong Resource(ResourceNum).ItemReward
    Buffer.WriteLong Resource(ResourceNum).ToolRequired
    Buffer.WriteLong Resource(ResourceNum).Health
    Buffer.WriteLong Resource(ResourceNum).RespawnTime
    Buffer.WriteLong Resource(ResourceNum).ToolRequired
    Buffer.WriteLong Resource(ResourceNum).Animation
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateResourceTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SHOPS

        If LenB(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendShops", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    Buffer.WriteLong Shop(ShopNum).BuyRate
    
    For i = 1 To MAX_TRADES
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).Item
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).ItemValue
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).CostItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).CostValue
    Next
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateShopToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteString Trim$(Shop(ShopNum).Name)
    Buffer.WriteLong Shop(ShopNum).BuyRate
    
    For i = 1 To MAX_TRADES
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).Item
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).ItemValue
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).CostItem
        Buffer.WriteLong Shop(ShopNum).TradeItem(i).CostValue
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateShopTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_SPELLS

        If LenB(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If

    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpells", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteString Trim$(Spell(SpellNum).Desc)
    Buffer.WriteByte Spell(SpellNum).Sound
    Buffer.WriteByte Spell(SpellNum).Type
    Buffer.WriteLong Spell(SpellNum).MPCost
    Buffer.WriteLong Spell(SpellNum).LevelReq
    Buffer.WriteLong Spell(SpellNum).AccessReq
    Buffer.WriteLong Spell(SpellNum).ClassReq
    Buffer.WriteLong Spell(SpellNum).CastTime
    Buffer.WriteLong Spell(SpellNum).CDTime
    Buffer.WriteLong Spell(SpellNum).Icon
    Buffer.WriteLong Spell(SpellNum).Map
    Buffer.WriteLong Spell(SpellNum).x
    Buffer.WriteLong Spell(SpellNum).y
    Buffer.WriteByte Spell(SpellNum).Dir
    Buffer.WriteLong Spell(SpellNum).Vital
    Buffer.WriteLong Spell(SpellNum).Duration
    Buffer.WriteLong Spell(SpellNum).Interval
    Buffer.WriteByte Spell(SpellNum).Range
    Buffer.WriteLong Spell(SpellNum).IsAoE
    Buffer.WriteLong Spell(SpellNum).CastAnim
    Buffer.WriteLong Spell(SpellNum).SpellAnim
    Buffer.WriteLong Spell(SpellNum).StunDuration
    Buffer.WriteLong Spell(SpellNum).BaseStat
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateSpellToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteString Trim$(Spell(SpellNum).Name)
    Buffer.WriteString Trim$(Spell(SpellNum).Desc)
    Buffer.WriteByte Spell(SpellNum).Sound
    Buffer.WriteByte Spell(SpellNum).Type
    Buffer.WriteLong Spell(SpellNum).MPCost
    Buffer.WriteLong Spell(SpellNum).LevelReq
    Buffer.WriteLong Spell(SpellNum).AccessReq
    Buffer.WriteLong Spell(SpellNum).ClassReq
    Buffer.WriteLong Spell(SpellNum).CastTime
    Buffer.WriteLong Spell(SpellNum).CDTime
    Buffer.WriteLong Spell(SpellNum).Icon
    Buffer.WriteLong Spell(SpellNum).Map
    Buffer.WriteLong Spell(SpellNum).x
    Buffer.WriteLong Spell(SpellNum).y
    Buffer.WriteByte Spell(SpellNum).Dir
    Buffer.WriteLong Spell(SpellNum).Vital
    Buffer.WriteLong Spell(SpellNum).Duration
    Buffer.WriteLong Spell(SpellNum).Interval
    Buffer.WriteByte Spell(SpellNum).Range
    Buffer.WriteLong Spell(SpellNum).IsAoE
    Buffer.WriteLong Spell(SpellNum).CastAnim
    Buffer.WriteLong Spell(SpellNum).SpellAnim
    Buffer.WriteLong Spell(SpellNum).StunDuration
    Buffer.WriteLong Spell(SpellNum).BaseStat
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateSpellTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSpells

    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerSpells", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count
    Buffer.WriteLong Resource_num

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then
        Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState
        Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
        Buffer.WriteLong ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y
    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendResourceCacheTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Long, ByVal Resource_num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SResourceCache
    Buffer.WriteLong ResourceCache(MapNum).Resource_Count
    Buffer.WriteLong Resource_num

    If ResourceCache(MapNum).Resource_Count > 0 Then

        For i = 0 To ResourceCache(MapNum).Resource_Count
            Buffer.WriteByte ResourceCache(MapNum).ResourceData(Resource_num).ResourceState
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(Resource_num).x
            Buffer.WriteLong ResourceCache(MapNum).ResourceData(Resource_num).y
        Next

    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendResourceCacheToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendActionMsg(ByVal MapNum As Long, ByVal Message As String, ByVal color As Long, ByVal MsgType As Long, ByVal x As Long, ByVal y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SActionMsg
    Buffer.WriteString Message
    Buffer.WriteLong color
    Buffer.WriteLong MsgType
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendActionMsg", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendBlood(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SBlood
    Buffer.WriteLong x
    Buffer.WriteLong y
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBlood", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAnimation(ByVal MapNum As Long, ByVal Anim As Long, ByVal x As Long, ByVal y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAnimation", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SCooldown
    Buffer.WriteLong Slot
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendCooldown", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendClearSpellBuffer(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendClearSpellBuffer", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SayMsg_Map(ByVal MapNum As Long, ByVal index As Long, ByVal Message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString Message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong saycolour
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg_Map", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal Message As String, ByVal saycolour As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString Message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong saycolour
    
    SendDataToAll Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SayMsg_Global", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SResetShopAction
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResetShopAction", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SStunned
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendStunned", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SBank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Bank(index).Item(i).Num
        Buffer.WriteLong Bank(index).Item(i).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendBank", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SOpenShop
    Buffer.WriteLong ShopNum
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendOpenShop", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPlayerMove
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerX(index)
    Buffer.WriteLong GetPlayerY(index)
    Buffer.WriteLong GetPlayerDir(index)
    Buffer.WriteLong Movement
    
    If Not sendToSelf Then
        SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    End If
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerMove", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTrade(ByVal index As Long, ByVal tradeTarget As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.STrade
    Buffer.WriteLong tradeTarget
    Buffer.WriteString GetPlayerName(tradeTarget)
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTrade", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SCloseTrade
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendCloseTrade", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal DataType As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim tradeTarget As Long
    Dim totalWorth As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    tradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.STradeUpdate
    Buffer.WriteByte DataType
    
    If DataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(i).Num).Type = iCurrency Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    ElseIf DataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(tradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Type = iCurrency Then
                    totalWorth = totalWorth + (Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price * TempPlayer(tradeTarget).TradeOffer(i).Value)
                Else
                    totalWorth = totalWorth + Item(GetPlayerInvItemNum(tradeTarget, TempPlayer(tradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    
    ' send total worth of trade
    Buffer.WriteLong totalWorth
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.STradeStatus
    Buffer.WriteLong Status
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTradeStatus", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTarget(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.STarget
    Buffer.WriteLong TempPlayer(index).target
    Buffer.WriteLong TempPlayer(index).targetType
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTarget", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendHotbar(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SHotbar
    For i = 1 To MAX_HOTBAR
        Buffer.WriteLong GetPlayerHotbarSlot(index, i)
        Buffer.WriteByte GetPlayerHotbarType(index, i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHotbar", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendLoginOk(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SLoginOk
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendLoginOk", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendInGame(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SInGame
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendInGame", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendHighIndex()
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SHighIndex
    Buffer.WriteLong Player_HighIndex
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendHighIndex", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerSound", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendMapSound(ByVal index As Long, ByVal x As Long, ByVal y As Long, ByVal entityType As Long, ByVal entityNum As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(GetPlayerMap(index)) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSound
    Buffer.WriteLong x
    Buffer.WriteLong y
    Buffer.WriteLong entityType
    Buffer.WriteLong entityNum
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendMapSound", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyUpdate(ByVal partyNum As Long)
    Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPartyUpdate
    Buffer.WriteByte 1
    Buffer.WriteLong Party(partyNum).Leader
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(partyNum).Member(i)
    Next
    Buffer.WriteLong Party(partyNum).MemberCount
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyUpdate", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long, partyNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPartyUpdate
    
    ' check if we' re in a party
    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(partyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(partyNum).Member(i)
        Next
        Buffer.WriteLong Party(partyNum).MemberCount
    Else
        ' send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyUpdateTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPartyVitals(ByVal partyNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SPartyVitals
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong GetPlayerVital(index, i)
    Next
    SendDataToParty partyNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPartyVitals", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendSpawnItemToMap(ByVal MapNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SSpawnItem
    Buffer.WriteLong index
    Buffer.WriteLong MapItem(MapNum, index).Num
    Buffer.WriteLong MapItem(MapNum, index).Value
    Buffer.WriteLong MapItem(MapNum, index).x
    Buffer.WriteLong MapItem(MapNum, index).y
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendSpawnItemToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendAttack(ByVal index As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SAttack
    Buffer.WriteLong index
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendAttack", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDoors(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOORS
        If LenB(Trim$(Door(i).Name)) > 0 Then
            Call SendUpdateDoorTo(index, i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDoors", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateDoorToAll(ByVal DoorNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateDoor
    Buffer.WriteLong DoorNum
    Buffer.WriteString Trim$(Door(DoorNum).Name)
    Buffer.WriteInteger Door(DoorNum).OpeningImage
    Buffer.WriteInteger Door(DoorNum).OpenWith
    Buffer.WriteLong Door(DoorNum).Respawn
    Buffer.WriteLong Door(DoorNum).Animation
    Buffer.WriteByte Door(DoorNum).Sound
    Buffer.WriteInteger Door(DoorNum).Map
    Buffer.WriteByte Door(DoorNum).x
    Buffer.WriteByte Door(DoorNum).y
    Buffer.WriteByte Door(DoorNum).Dir
    Buffer.WriteInteger Door(DoorNum).LevelReq
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Door(DoorNum).Stat_Req(i)
    Next
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateDoorToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateDoorTo(ByVal index As Long, ByVal DoorNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateDoor
    Buffer.WriteLong DoorNum
    Buffer.WriteString Door(DoorNum).Name
    Buffer.WriteInteger Door(DoorNum).OpeningImage
    Buffer.WriteInteger Door(DoorNum).OpenWith
    Buffer.WriteLong Door(DoorNum).Respawn
    Buffer.WriteLong Door(DoorNum).Animation
    Buffer.WriteByte Door(DoorNum).Sound
    Buffer.WriteInteger Door(DoorNum).Map
    Buffer.WriteByte Door(DoorNum).x
    Buffer.WriteByte Door(DoorNum).y
    Buffer.WriteByte Door(DoorNum).Dir
    Buffer.WriteInteger Door(DoorNum).LevelReq
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteByte Door(DoorNum).Stat_Req(i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateDoorTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDoorCacheTo(ByVal index As Long, ByVal Door_num As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SDoorCache
    Buffer.WriteLong DoorCache(GetPlayerMap(index)).Count
    Buffer.WriteLong Door_num

    If DoorCache(GetPlayerMap(index)).Count > 0 Then
        Buffer.WriteLong DoorCache(GetPlayerMap(index)).Data(Door_num).x
        Buffer.WriteLong DoorCache(GetPlayerMap(index)).Data(Door_num).y
        Buffer.WriteLong DoorCache(GetPlayerMap(index)).Data(Door_num).State
    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDoorCacheTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDoorCacheToMap(ByVal MapNum As Long, ByVal Door_num As Long)
    Dim Buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SDoorCache
    Buffer.WriteLong DoorCache(MapNum).Count
    Buffer.WriteLong Door_num

    If DoorCache(MapNum).Count > 0 Then
        Buffer.WriteLong DoorCache(MapNum).Data(Door_num).x
        Buffer.WriteLong DoorCache(MapNum).Data(Door_num).y
        Buffer.WriteLong DoorCache(MapNum).Data(Door_num).State
    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDoorCacheToMap", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendDialogue(ByVal index As Integer, ByVal Title As String, ByVal Text As String, ByVal dType As Byte, ByVal YesNo As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SDialogue
    Buffer.WriteString Title
    Buffer.WriteString Text
    Buffer.WriteByte dType
    Buffer.WriteLong YesNo
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendDialogue", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendQuests(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUESTS
        If LenB(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendQuests", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateQuestToAll(ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, x As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(Quest(QuestNum).Name)
    Buffer.WriteString Trim$(Quest(QuestNum).Description)
    Buffer.WriteLong Quest(QuestNum).Retry
    Buffer.WriteInteger Quest(QuestNum).LevelReq
    Buffer.WriteInteger Quest(QuestNum).QuestReq
    Buffer.WriteByte Quest(QuestNum).ClassReq
    Buffer.WriteInteger Quest(QuestNum).SpriteReq
    Buffer.WriteInteger Quest(QuestNum).LevelRew
    Buffer.WriteLong Quest(QuestNum).ExpRew
    Buffer.WriteByte Quest(QuestNum).ClassRew
    Buffer.WriteInteger Quest(QuestNum).SpriteRew
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong Quest(QuestNum).StatReq(i)
        Buffer.WriteLong Quest(QuestNum).StatRew(i)
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong Quest(QuestNum).VitalRew(i)
    Next
    
    For i = 1 To MAX_QUEST_TASKS
        Buffer.WriteByte Quest(QuestNum).Task(i).Type
        Buffer.WriteLong Quest(QuestNum).Task(i).Instant
        Buffer.WriteInteger Quest(QuestNum).Task(i).Num
        Buffer.WriteLong Quest(QuestNum).Task(i).Value
    
        For x = 1 To 3
            Buffer.WriteString Trim$(Quest(QuestNum).Task(i).Message(x))
        Next
    Next
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateQuestToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long, x As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateQuest
    Buffer.WriteLong QuestNum
    Buffer.WriteString Trim$(Quest(QuestNum).Name)
    Buffer.WriteString Trim$(Quest(QuestNum).Description)
    Buffer.WriteLong Quest(QuestNum).Retry
    Buffer.WriteInteger Quest(QuestNum).LevelReq
    Buffer.WriteInteger Quest(QuestNum).QuestReq
    Buffer.WriteByte Quest(QuestNum).ClassReq
    Buffer.WriteInteger Quest(QuestNum).SpriteReq
    Buffer.WriteInteger Quest(QuestNum).LevelRew
    Buffer.WriteLong Quest(QuestNum).ExpRew
    Buffer.WriteByte Quest(QuestNum).ClassRew
    Buffer.WriteInteger Quest(QuestNum).SpriteRew
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong Quest(QuestNum).StatReq(i)
        Buffer.WriteLong Quest(QuestNum).StatRew(i)
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong Quest(QuestNum).VitalRew(i)
    Next
    
    For i = 1 To MAX_QUEST_TASKS
        Buffer.WriteByte Quest(QuestNum).Task(i).Type
        Buffer.WriteLong Quest(QuestNum).Task(i).Instant
        Buffer.WriteInteger Quest(QuestNum).Task(i).Num
        Buffer.WriteLong Quest(QuestNum).Task(i).Value
    
        For x = 1 To 3
            Buffer.WriteString Trim$(Quest(QuestNum).Task(i).Message(x))
        Next
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateQuestTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendQuestCommand(ByVal index As Integer, ByVal Command As Byte, Optional Value As Long = 0)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SQuestCommand
    Buffer.WriteByte Command
    Buffer.WriteLong Value
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendQuestCommand", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendNpcXY(ByVal MapNum As Integer, ByVal MapNpcNum As Integer, ByVal Movement As Byte)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcMove
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
    Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
    Buffer.WriteLong Movement
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendNpcXY", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCloseShop(ByVal index As Integer)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SCloseShop
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendCloseShop", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCloseBank(ByVal index As Integer)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SCloseBank
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendCloseBank", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendTitles(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_TITLES
        If LenB(Trim$(Title(i).Name)) > 0 Then
            Call SendUpdateTitleTo(index, i)
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendTitles", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateTitleToAll(ByVal TitleNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateTitle
    Buffer.WriteLong TitleNum
    Buffer.WriteString Trim$(Title(TitleNum).Name)
    Buffer.WriteString Trim$(Title(TitleNum).Description)
    Buffer.WriteInteger Title(TitleNum).Icon
    Buffer.WriteByte Title(TitleNum).Type
    Buffer.WriteByte Title(TitleNum).color
    Buffer.WriteByte Title(TitleNum).Sound
    Buffer.WriteLong Title(TitleNum).UseAnimation
    Buffer.WriteLong Title(TitleNum).RemoveAnimation
    Buffer.WriteLong Title(TitleNum).Passive
    Buffer.WriteLong Title(TitleNum).LevelReq
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong Title(TitleNum).StatReq(i)
        Buffer.WriteLong Title(TitleNum).StatRew(i)
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong Title(TitleNum).VitalRew(i)
    Next
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateTitleToAll", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendUpdateTitleTo(ByVal index As Long, ByVal TitleNum As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SUpdateTitle
    Buffer.WriteLong TitleNum
    Buffer.WriteString Trim$(Title(TitleNum).Name)
    Buffer.WriteString Trim$(Title(TitleNum).Description)
    Buffer.WriteInteger Title(TitleNum).Icon
    Buffer.WriteByte Title(TitleNum).Type
    Buffer.WriteByte Title(TitleNum).color
    Buffer.WriteByte Title(TitleNum).Sound
    Buffer.WriteLong Title(TitleNum).UseAnimation
    Buffer.WriteLong Title(TitleNum).RemoveAnimation
    Buffer.WriteLong Title(TitleNum).Passive
    Buffer.WriteLong Title(TitleNum).LevelReq
    
    For i = 1 To Stats.Stat_Count - 1
        Buffer.WriteLong Title(TitleNum).StatReq(i)
        Buffer.WriteLong Title(TitleNum).StatRew(i)
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong Title(TitleNum).VitalRew(i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendUpdateTitleTo", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendPlayerTitles(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.STitles

    For i = 1 To MAX_PLAYER_TITLES
        Buffer.WriteLong GetPlayerTitle(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendPlayerTitles", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SendCharactersData(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SCharData
    For i = 1 To MAX_PLAYER_CHARS
        Buffer.WriteString Player(index).Char(i).Name
        Buffer.WriteLong Player(index).Char(i).Level
        Buffer.WriteLong Player(index).Char(i).Class
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SendCharactersData", "modServerTCP", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
