Attribute VB_Name = "modPlayer"
Option Explicit

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent someone who is already online try logging
    If IsPlaying(index) Then Exit Sub

    ' Set the flag so we know the person is in the game
    TempPlayer(index).InGame = True
    
    ' Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)

    ' send the login ok
    Call SendLoginOk(index)
    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send Sounde more little goodies, no need to explain these
    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call SendMapEquipment(index)
    Call SendPlayerSpells(index)
    Call SendHotbar(index)
    Call SendDoors(index)
    Call SendQuests(index)
    Call SendTitles(index)

    ' send vitals, exp + stats
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    Call SendEXP(index)
    Call SendStats(index)
    
    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
    ' Send a global message that he/she joined
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Game_Name & "!", White)
    End If
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        Call SendResourceCacheTo(index, i)
    Next

    For i = 0 To DoorCache(GetPlayerMap(index)).Count
        Call SendDoorCacheTo(index, i)
    Next
    
    ' Send the flag so they know they can start doing stuff
    Call SendInGame(index)

    ' Server log
    Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".", PLAYER_LOG)
    Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Game_Name & ".")
    
    ' Update the maximum players who is online
    Call UpdateCaption

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "JoinGame", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LeftGame(ByVal index As Long)
    Dim i As Long
    Dim tradeTarget As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Not relog the player if it is already online
    If Not TempPlayer(index).InGame Then Exit Sub

    ' If the player is in a group conversation, remove it from the group
    If TempPlayer(index).roomIndex > 0 Then
        If ChatRoom(TempPlayer(index).roomIndex).Members - 1 = 0 Then
            RemoveChat (TempPlayer(index).roomIndex)
        End If
    End If
         
    ' Removes the player of the game
    TempPlayer(index).InGame = False

    ' Check if player was the only player on the map and stop npc processing if so
    If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
        PlayersOnMap(GetPlayerMap(index)) = NO
    End If
        
    ' cancel any trade they' re in
    If TempPlayer(index).InTrade > 0 Then
        tradeTarget = TempPlayer(index).InTrade
        PlayerMsg tradeTarget, GetPlayerName(index) & " has declined the trade.", BrightRed
        ' clear out trade
        For i = 1 To MAX_INV
            TempPlayer(tradeTarget).TradeOffer(i).Num = 0
            TempPlayer(tradeTarget).TradeOffer(i).Value = 0
        Next
        TempPlayer(tradeTarget).InTrade = 0
        SendCloseTrade tradeTarget
    End If
        
    ' leave party.
    Call Party_PlayerLeave(index)
        
    ' clears
    Call ClearTargets(index)
        
    ' save and clear data.
    Call SavePlayer(index)
    Call SaveBank(index)
    Call ClearBank(index)

    ' Send a global message that he/she left
    If GetPlayerAccess(index) <= ADMIN_MONITOR Then
        Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", JoinLeftColor)
    Else
        Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Game_Name & "!", White)
    End If

    Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Game_Name & ".")
        
    ' Disconnects the player of the game
    Call SendLeftGame(index)
    Call ClearPlayer(index)
    TotalPlayersOnline = TotalPlayersOnline - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LeftGame", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim i As Byte
    Dim x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    ' Damage by stat
    x = GetPlayerStat(index, Stats.Strength) \ 2.6
    
    ' Damage by level
    x = x + (GetPlayerLevel(index) \ 4)

    ' Damage by equipments
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            x = x + Item(GetPlayerEquipment(index, i)).Damage \ 1.8
        End If
    Next

    ' Return value of function
    GetPlayerDamage = x

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDamage", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim i As Byte
    Dim x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Or index <= 0 Or index > Player_HighIndex Then Exit Function
    
    ' Protection by stat
    x = GetPlayerStat(index, Stats.Endurance) \ 2.8

    ' Protection by level
    x = x + (GetPlayerLevel(index) \ 4)
    
    ' Protection by equipments
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            x = x + Item(GetPlayerEquipment(index, i)).Protection \ 1.8
        End If
    Next

    ' Return value of function
    GetPlayerProtection = x

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerProtection", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim OldMap As Long
    Dim i As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check if you are out of bounds
    If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
    If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    If x < 0 Then x = 0
    If y < 0 Then y = 0
    
    ' clear target
    TempPlayer(index).target = 0
    TempPlayer(index).targetType = TARGET_TYPE_NONE
    Call SendTarget(index)

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)

    If OldMap <> MapNum Then
        Call SendLeaveMap(index, OldMap)
    End If

    ' Set player location
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
    
    ' if same map then just send their co-ordinates
    If MapNum <> GetPlayerMap(index) Then
        SendPlayerXYToMap index
    End If
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(MapNum) = YES
    TempPlayer(index).GettingMap = YES
    
    ' Clears the target of players that are marked on the player index
    Call ClearTargets(index)

    ' Check if the player completed a quest
    Call CheckCompleteQuest(index, MapNum)
    
    ' send player' s equipment to new map
    Call SendMapEquipment(index)
    
    ' Sets it so we know to process npcs on the map
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SCheckForMap
    Buffer.WriteLong MapNum
    Buffer.WriteLong Map(MapNum).Revision
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerWarp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Optional ByVal sendToSelf As Boolean = False)
    Dim MapNum As Long
    Dim x As Long, y As Long
    Dim Moved As Byte
    Dim NewMapX As Byte, NewMapY As Byte
    Dim VitalType As Long, Colour As Long, amount As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub

    Call SetPlayerDir(index, Dir)
    MapNum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Or GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) And Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index) - 1, GetPlayerY(index) - 1) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).UpLeft > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).UpLeft).MaxX
                    NewMapY = Map(Map(GetPlayerMap(index)).UpLeft).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).UpLeft, NewMapX, NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
          
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Or GetPlayerX(index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) And Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index) + 1, GetPlayerY(index) - 1) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).UpRight > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).UpRight).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).UpRight, 0, NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
          
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Or GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index) - 1, GetPlayerY(index) + 1) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If
            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).DownLeft > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).DownLeft).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).DownLeft, NewMapX, 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
          
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Or GetPlayerX(index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) And Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index) + 1, GetPlayerY(index) + 1) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).DownRight > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).DownRight, 0, 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index), GetPlayerY(index) - 1) Then
                        Call SetPlayerY(index, GetPlayerY(index) - 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If

            Else

                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    NewMapY = Map(Map(GetPlayerMap(index)).Up).MaxY
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), NewMapY)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index), GetPlayerY(index) + 1) Then
                        Call SetPlayerY(index, GetPlayerY(index) + 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index) - 1, GetPlayerY(index)) Then
                        Call SetPlayerX(index, GetPlayerX(index) - 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    NewMapX = Map(Map(GetPlayerMap(index)).Left).MaxX
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, NewMapX, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(MapNum).MaxX Then

                ' Check to make sure that the tile is walkable
                If Not isDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Not CanBlockedTile(index, GetPlayerX(index) + 1, GetPlayerY(index)) Then
                        Call SetPlayerX(index, GetPlayerX(index) + 1)
                        SendPlayerMove index, Movement, sendToSelf
                        Moved = YES
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    ' clear their target
                    TempPlayer(index).target = 0
                    TempPlayer(index).targetType = TARGET_TYPE_NONE
                    SendTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            x = .Data2
            y = .Data3
            Call PlayerWarp(index, MapNum, x, y)
            Moved = YES
        End If

        ' Check for key trigger open
        If .Type = TILE_TYPE_KEYOPEN Then
            OpenDoor index, .Data1, .Data2
        End If
        
        ' Check door
        If .Type = TILE_TYPE_DOOR Then
            CheckDoor index, GetPlayerX(index), GetPlayerY(index)
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            x = .Data1
            If x > 0 Then ' shop exists?
                If Len(Trim$(Shop(x).Name)) > 0 Then ' name exists?
                    SendOpenShop index, x
                    TempPlayer(index).InShop = x ' stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            TempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it' s a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            amount = .Data2
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If VitalType = Vitals.HP Then
                    Colour = BrightGreen
                Else
                    Colour = BrightBlue
                End If
                SendActionMsg GetPlayerMap(index), "+" & amount, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + amount
                PlayerMsg index, "You feel rejuvinating forces flowing through your boy.", BrightGreen
                Call SendVital(index, VitalType)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it' s a trap tile
        If .Type = TILE_TYPE_TRAP Then
            amount = .Data1
            SendActionMsg GetPlayerMap(index), "-" & amount, BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
            If GetPlayerVital(index, Vitals.HP) - amount <= 0 Then
                KillPlayer index
                PlayerMsg index, "You' re killed by a trap.", BrightRed
            Else
                SetPlayerVital index, Vitals.HP, GetPlayerVital(index, HP) - amount
                PlayerMsg index, "You' re injured by a trap.", BrightRed
                Call SendVital(index, Vitals.HP)
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            Select Case .Data1
                Case DIR_UP
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_LEFT
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_DOWN
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_BLOCKED Then Exit Sub
                Case DIR_RIGHT
                    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_RESOURCE Or Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_BLOCKED Then Exit Sub
            End Select
            ForcePlayerMove index, MOVING_WALKING, .Data1
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        PlayerWarp index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMove", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal Movement As Long, ByVal Direction As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If Direction < DIR_UP Or Direction > DIR_DOWN_RIGHT Then Exit Sub
    If Movement < 1 Or Movement > 2 Then Exit Sub
   
    ' Prevent subscript out range
    If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Or GetPlayerX(index) = 0 Then Exit Sub
    If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Or GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub

    ' Move the player
    PlayerMove index, Direction, Movement, True

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ForcePlayerMove", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim ItemNum As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ' The item
        ItemNum = GetPlayerEquipment(index, i)

        ' Equip the equipment
        If ItemNum > 0 Then
            Select Case i
                Case Equipment.Weapon

                    If Item(ItemNum).Type <> iWeapon Then SetPlayerEquipment index, 0, i
                Case Equipment.Armor

                    If Item(ItemNum).Type <> iArmor Then SetPlayerEquipment index, 0, i
                Case Equipment.Helmet

                    If Item(ItemNum).Type <> iHelmet Then SetPlayerEquipment index, 0, i
                Case Equipment.Shield

                    If Item(ItemNum).Type <> iShield Then SetPlayerEquipment index, 0, i
            End Select
        Else
            SetPlayerEquipment index, 0, i
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckEquippedItems", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    If Item(ItemNum).Type = iCurrency Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next
    End If

    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenInvSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindOpenInvSlots(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    If Item(ItemNum).Type = iCurrency Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlots = 1
                Exit Function
            End If
        Next
    End If

    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlots = FindOpenInvSlots + 1
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenInvSlots", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    ' Return function value
    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = ItemNum Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenBankSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = iCurrency Then
                HasItem = HasItem + GetPlayerInvItemValue(index, i)
            Else
                HasItem = HasItem + 1
            End If
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "HasItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TakeInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal QuestMessage As Boolean = True) As Boolean
    Dim i As Long
    Dim x As Byte
    Dim HotbarSlot As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    If ItemVal = 0 Then ItemVal = 1

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            x = x + 1
            If Item(ItemNum).Type = iCurrency Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                ElseIf ItemVal < GetPlayerInvItemValue(index, i) Then
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                If x > ItemVal Then
                    Exit Function
                Else
                    TakeInvItem = True
                End If
            End If
            
            ' Take?
            If TakeInvItem Then
                ' Take item
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                        
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
            End If
        End If
    Next

    ' Check if the player completed a quest
    Call CheckCompleteQuest(index, , QuestMessage)
                        
    ' Remove item of hotbar
    HotbarSlot = FindHotbar(index, ItemNum, 1)
    If HotbarSlot > 0 Then
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).SType = 0
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).Slot = 0
        SendHotbar index
    End If
                
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TakeInvItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TakeInvSlot(ByVal index As Long, ByVal invSlot As Long, ByVal ItemVal As Long) As Boolean
    Dim ItemNum As Long
    Dim HotbarSlot As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Or invSlot <= 0 Or invSlot > MAX_ITEMS Then Exit Function

    ' The item
    ItemNum = GetPlayerInvItemNum(index, invSlot)

    ' Prevent subscript out of range
    If ItemNum <= 0 Then Exit Function
    
    If Item(ItemNum).Type = iCurrency Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, invSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, invSlot, GetPlayerInvItemValue(index, invSlot) - ItemVal)
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        ' Take item
        Call SetPlayerInvItemNum(index, invSlot, 0)
        Call SetPlayerInvItemValue(index, invSlot, 0)
        Call SendInventoryUpdate(index, invSlot)
    End If

    ' Check if the player completed a quest
    Call CheckCompleteQuest(index, , True)

    ' Remove item of hotbar
    HotbarSlot = FindHotbar(index, ItemNum, 1)
    If HotbarSlot > 0 Then
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).SType = 0
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).Slot = 0
        SendHotbar index
    End If
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "TakeInvSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GiveInvItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, Optional ByVal sendUpdate As Boolean = True) As Boolean
    Dim i As Long
    Dim x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function
    If ItemVal = 0 Then ItemVal = 1
    
    If Item(ItemNum).Type = iCurrency Then
        ' Find an empty slot in player inventory
        i = FindOpenInvSlot(index, ItemNum)

        ' Check to see if inventory is full
        If i <> 0 Then
            ' Give item
            Call SetPlayerInvItemNum(index, i, ItemNum)
            Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
            
            ' Update?
            If sendUpdate Then Call SendInventoryUpdate(index, i)
            GiveInvItem = True
        Else
            Call PlayerMsg(index, "Your inventory is full.", BrightRed)
            Exit Function
        End If
    Else
        ' Check to see if inventory is full
        If FindOpenInvSlots(index, ItemNum) < ItemVal Then
            Call PlayerMsg(index, "You can not carry all these items.", BrightRed)
            Exit Function
        End If
        
        For x = 1 To ItemVal
            ' Find an empty slot in player inventory
            i = FindOpenInvSlot(index, ItemNum)

            ' Give item to player
            Call SetPlayerInvItemNum(index, i, ItemNum)
            Call SetPlayerInvItemValue(index, i, 1)
                
            ' Update?
            If sendUpdate Then Call SendInventoryUpdate(index, i)
            GiveInvItem = True
        Next
    End If
    
    ' Check if the player completed a quest
    Call CheckCompleteQuest(index, , True)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GiveInvItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_SPELLS
        ' Return function value
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "HasSpell", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_SPELLS
        ' Return function value
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenSpellSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim MapNum As Long
    Dim Msg As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Player map
    MapNum = GetPlayerMap(index)

    For i = MAX_MAP_ITEMS To 1 Step -1
        ' See if theres even an item here
        If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).x = GetPlayerX(index)) And (MapItem(MapNum, i).y = GetPlayerY(index)) Then
                ' Find open slot
                n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)
    
                ' Open slot available?
                If n <> 0 Then
                    ' Set item in players inventor
                    Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)
    
                    If Item(GetPlayerInvItemNum(index, n)).Type = iCurrency Then
                        Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                        Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                    Else
                        Call SetPlayerInvItemValue(index, n, 0)
                        Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                    End If
    
                    ' Erase item from the map
                    ClearMapItem i, MapNum
                            
                    ' Check if the player completed a quest
                    Call CheckCompleteQuest(index, , True)
                            
                    ' Update player inventory
                    Call SendInventoryUpdate(index, n)
                    Call SpawnItemSlot(i, 0, 0, GetPlayerMap(index), 0, 0)
                    SendActionMsg GetPlayerMap(index), Msg, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    Exit Sub
                Else
                    Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                    Exit Sub
                End If
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMapGetItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal invNum As Long, ByVal amount As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Or invNum <= 0 Or invNum > MAX_INV Then Exit Sub

    ' check the player isn' t doing Soundething
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub

    ' Prevent subscript out of range
    If GetPlayerInvItemNum(index, invNum) <= 0 Or GetPlayerInvItemNum(index, invNum) > MAX_ITEMS Then Exit Sub
    
    ' Finds some empty slot on the map to add the item
    i = FindOpenMapItemSlot(GetPlayerMap(index))

    ' Checks if can drop the item
    If i <= 0 Then
        Call PlayerMsg(index, "Too many items already on the ground.", BrightRed)
        Exit Sub
    End If
    
    ' Place the item on the ground
    MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, invNum)
    MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
    MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
    MapItem(GetPlayerMap(index), i).canDespawn = True
    MapItem(GetPlayerMap(index), i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME

    If Item(GetPlayerInvItemNum(index, invNum)).Type = iCurrency Then
        ' Check if its more then they have and if so drop it all
        If amount >= GetPlayerInvItemValue(index, invNum) Then
            MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, invNum)
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, invNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
            Call SetPlayerInvItemNum(index, invNum, 0)
            Call SetPlayerInvItemValue(index, invNum, 0)
        Else
            MapItem(GetPlayerMap(index), i).Value = amount
            Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & amount & " " & Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name) & ".", Yellow)
            Call SetPlayerInvItemValue(index, invNum, GetPlayerInvItemValue(index, invNum) - amount)
        End If
    Else
        ' Its not a currency object so this is easy
        MapItem(GetPlayerMap(index), i).Value = 0
                    
        ' send message
        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, invNum)).Name)) & ".", Yellow)
        Call SetPlayerInvItemNum(index, invNum, 0)
        Call SetPlayerInvItemValue(index, invNum, 0)
    End If

    ' Check if the player completed a quest
    Call CheckCompleteQuest(index, , True)
                
    ' Send inventory update
    Call SendInventoryUpdate(index, invNum)
                
    ' Spawn the item before we set the num or we' ll get a different free map item slot
    Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, amount, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), MapItem(GetPlayerMap(index), i).canDespawn)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerMapDropItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim expRollover As Long
    Dim level_count As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
        expRollover = GetPlayerExp(index) - GetPlayerNextLevel(index)
        
        ' can level up?
        If Not SetPlayerLevel(index, GetPlayerLevel(index) + 1) Then Exit Sub
        level_count = level_count + 1

        ' Rewards to pass of level
        Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + 3)
        Call SetPlayerExp(index, expRollover)
    Loop
    
    If level_count > 0 Then
        If level_count = 1 Then
            ' singular
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " level!", Brown
        Else
            ' plural
            GlobalMsg GetPlayerName(index) & " has gained " & level_count & " levels!", Brown
        End If
        
        ' Update player
        Call CheckTitle(index)
        Call SendEXP(index)
        Call SendPlayerData(index)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPlayerLevelUp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' //////////////////////
' // PLAYER FUNCTIONS //
' //////////////////////
Function GetPlayerLogin(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerLogin = Trim$(Player(index).Login)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLogin", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Login = Login

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerLogin", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerPassword = Trim$(Player(index).Password)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPassword", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Password = Password

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPassword", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerName(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerName = Trim$(Player(index).Char(TempPlayer(index).Char).Name)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerName", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Name = Name

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerName", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerClass = Player(index).Char(TempPlayer(index).Char).Class

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerClass", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Class = ClassNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerClass", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerSprite = Player(index).Char(TempPlayer(index).Char).Sprite

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSprite", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Sprite = Sprite

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSprite", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerLevel = Player(index).Char(TempPlayer(index).Char).Level

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function SetPlayerLevel(ByVal index As Long, ByVal Level As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SetPlayerLevel = False
    If Level > MAX_LEVELS Then Exit Function
    Player(index).Char(TempPlayer(index).Char).Level = Level
    SetPlayerLevel = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "SetPlayerLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerNextLevel(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerNextLevel", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerExp = Player(index).Char(TempPlayer(index).Char).EXP

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerExp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).EXP = EXP

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerExp", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerAccess = Player(index).Char(TempPlayer(index).Char).Access

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerAccess", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Access = Access

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerAccess", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerPK = Player(index).Char(TempPlayer(index).Char).PK

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPK", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).PK = PK

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPK", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerVital = Player(index).Char(TempPlayer(index).Char).Vital(Vital)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVital", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then Player(index).Char(TempPlayer(index).Char).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    If GetPlayerVital(index, Vital) < 0 Then Player(index).Char(TempPlayer(index).Char).Vital(Vital) = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerVital", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    ' Original player stat
    x = Player(index).Char(TempPlayer(index).Char).Stat(Stat)
    
    ' Stat by title
    x = GetTitleStat(index, Stat, x)

    ' Stat by equipment
    For i = 1 To Equipment.Equipment_Count - 1
        If Player(index).Char(TempPlayer(index).Char).Equipment(i) > 0 Then
            If Item(Player(index).Char(TempPlayer(index).Char).Equipment(i)).Add_Stat(Stat) > 0 Then
                x = x + Item(Player(index).Char(TempPlayer(index).Char).Equipment(i)).Add_Stat(Stat)
            End If
        End If
    Next
    
    ' Return function value
    GetPlayerStat = x

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerStat", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerRawStat = Player(index).Char(TempPlayer(index).Char).Stat(Stat)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerRawStat", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Stat(Stat) = Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerStat", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerPOINTS = Player(index).Char(TempPlayer(index).Char).POINTS

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerPOINTS", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If POINTS <= 0 Then POINTS = 0
    Player(index).Char(TempPlayer(index).Char).POINTS = POINTS

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerPOINTS", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerMap = Player(index).Char(TempPlayer(index).Char).Map

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMap", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(TempPlayer(index).Char).Map = MapNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerMap", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerX = Player(index).Char(TempPlayer(index).Char).x

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerX", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).x = x

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerX", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerY = Player(index).Char(TempPlayer(index).Char).y

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerY", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).y = y

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerY", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerDir = Player(index).Char(TempPlayer(index).Char).Dir

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDir", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Dir = Dir

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerDir", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerIP", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    If invSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Player(index).Char(TempPlayer(index).Char).Inv(invSlot).Num

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal invSlot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Inv(invSlot).Num = ItemNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerInvItemValue = Player(index).Char(TempPlayer(index).Char).Inv(invSlot).Value

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerInvItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal invSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Inv(invSlot).Value = ItemValue

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerInvItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal Spellslot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerSpell = Player(index).Char(TempPlayer(index).Char).Spell(Spellslot)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSpell", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal Spellslot As Long, ByVal SpellNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Spell(Spellslot) = SpellNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSpell", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Equipment) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(index).Char(TempPlayer(index).Char).Equipment(EquipmentSlot)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerEquipment", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal invNum As Long, ByVal EquipmentSlot As Equipment)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Equipment(EquipmentSlot) = invNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerEquipment", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerBankItemNum = Bank(index).Item(BankSlot).Num

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerBankItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Bank(index).Item(BankSlot).Num = ItemNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerBankItemNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerBankItemValue = Bank(index).Item(BankSlot).Value

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerBankItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Bank(index).Item(BankSlot).Value = ItemValue

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerBankItemValue", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerQuestNum(ByVal index As Long, ByVal Slot As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerQuestNum = Player(index).Char(TempPlayer(index).Char).Quests(Slot).Num

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerQuestNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerQuestNum(ByVal index As Long, ByVal Slot As Byte, ByVal Num As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Quests(Slot).Num = Num

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerQuestNum", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerQuestStatus(ByVal index As Long, ByVal Slot As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerQuestStatus = Player(index).Char(TempPlayer(index).Char).Quests(Slot).Status

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerQuestStatus", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerQuestStatus(ByVal index As Long, ByVal Slot As Byte, ByVal Status As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Quests(Slot).Status = Status

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerQuestStatus", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerQuestPart(ByVal index As Long, ByVal Slot As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerQuestPart = Player(index).Char(TempPlayer(index).Char).Quests(Slot).Part

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerQuestPart", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerQuestPart(ByVal index As Long, ByVal Slot As Byte, ByVal Part As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Quests(Slot).Part = Part

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerQuestPart", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerKillNpcs(ByVal index As Long, ByVal Slot As Byte) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or Slot = 0 Then Exit Function
    GetPlayerKillNpcs = Player(index).Char(TempPlayer(index).Char).KillNpcs(Slot)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerKillNpcs", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerKillNpcs(ByVal index As Long, ByVal Slot As Byte, ByVal Num As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).KillNpcs(Slot) = Num

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerKillNpcs", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerKillPlayers(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerKillPlayers = Player(index).Char(TempPlayer(index).Char).KillPlayers

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerKillPlayers", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerKillPlayers(ByVal index As Long, ByVal NumPlayers As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).KillPlayers = NumPlayers

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerKillPlayers", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerSex(ByVal index As Integer) As Byte
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerSex = Player(index).Char(TempPlayer(index).Char).Sex

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerSex", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerSex(ByVal index As Integer, ByVal Sex As Byte)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Sex = Sex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerSex", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerHotbarSlot(ByVal index As Integer, ByVal Slot As Integer) As Integer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerHotbarSlot = Player(index).Char(TempPlayer(index).Char).Hotbar(Slot).Slot

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerHotbarSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerHotbarSlot(ByVal index As Integer, ByVal Slot As Integer, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Hotbar(Slot).Slot = Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerHotbarSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetPlayerHotbarType(ByVal index As Integer, ByVal SType As Integer) As Integer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerHotbarType = Player(index).Char(TempPlayer(index).Char).Hotbar(SType).SType

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerHotbarType", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SetPlayerHotbarType(ByVal index As Integer, ByVal SType As Integer, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Hotbar(SType).SType = Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerHotbarType", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ToDo
Sub OnDeath(ByVal index As Long)
    Dim i As Long
    Dim ClassNum As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' player class
    ClassNum = GetPlayerClass(index)
    
    ' Set HP to nothing
    Call SetPlayerVital(index, Vitals.HP, 0)

    ' Drop all worn items
    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(index, i) > 0 Then
            PlayerMapDropItem index, GetPlayerEquipment(index, i), 0
        End If
    Next

    ' Warp player away
    Call SetPlayerDir(index, DIR_DOWN)
    
    With Map(GetPlayerMap(index))
        ' to the bootmap if it is set
        If .BootMap > 0 Then
            PlayerWarp index, .BootMap, .BootX, .BootY
        Else
            Call PlayerWarp(index, Class(ClassNum).StartMap, Class(ClassNum).StartX, Class(ClassNum).StartY)
        End If
    End With
    
    ' clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With TempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear Spell casting
    TempPlayer(index).SpellBuffer.Spell = 0
    TempPlayer(index).SpellBuffer.Timer = 0
    TempPlayer(index).SpellBuffer.target = 0
    TempPlayer(index).SpellBuffer.tType = 0
    Call SendClearSpellBuffer(index)
    
    ' Close bank for player
    TempPlayer(index).InBank = False
    SendCloseBank index
    
    ' Close shop for player
    TempPlayer(index).InShop = 0
    SendCloseShop index
    
    ' Close trade for player
    If TempPlayer(index).InTrade > 0 Then
        ' Clear trade
        For i = 1 To MAX_INV
            TempPlayer(index).TradeOffer(i).Num = 0
            TempPlayer(index).TradeOffer(i).Value = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Num = 0
            TempPlayer(TempPlayer(index).InTrade).TradeOffer(i).Value = 0
        Next

        ' Close trade
        TempPlayer(index).InTrade = 0
        TempPlayer(TempPlayer(index).InTrade).InTrade = 0
        
        ' Update trade
        SendCloseTrade index
        SendCloseTrade TempPlayer(index).InTrade
    End If
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
    Call SendVital(index, Vitals.HP)
    Call SendVital(index, Vitals.MP)
    
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index

    ' If the player the attacker killed was a pk then take it away
    If GetPlayerPK(index) = YES Then
        Call SetPlayerPK(index, NO)
        Call SendPlayerData(index)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OnDeath", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckResource(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim Resource_num As Long
    Dim Resource_index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check if tile has resource
    If Map(GetPlayerMap(index)).Tile(x, y).Type <> TILE_TYPE_RESOURCE Then Exit Sub
    
    ' The resource
    Resource_num = 0
    Resource_index = Map(GetPlayerMap(index)).Tile(x, y).Data1

    ' Get the cache number
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        If ResourceCache(GetPlayerMap(index)).ResourceData(i).x = x Then
            If ResourceCache(GetPlayerMap(index)).ResourceData(i).y = y Then
                Resource_num = i
            End If
        End If
    Next

    ' Prevent subscript out range
    If Resource_num <= 0 Then Exit Sub
    
    ' Check if player has weapon
    If GetPlayerEquipment(index, Weapon) <= 0 Then
        PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
        Exit Sub
    End If
            
    ' Check if player weapon is the requerid
    If Item(GetPlayerEquipment(index, Weapon)).Data3 <> Resource(Resource_index).ToolRequired Then
        PlayerMsg index, "You have the wrong type of tool equiped.", BrightRed
        Exit Sub
    End If

    ' inv space?
    If Resource(Resource_index).ItemReward > 0 Then
        If FindOpenInvSlot(index, Resource(Resource_index).ItemReward) = 0 Then
            PlayerMsg index, "You have no inventory space.", BrightRed
            Exit Sub
        End If
    End If

    ' check if already cut down
    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 0 Then
        ' send message if it exists
        If Len(Trim$(Resource(Resource_index).EmptyMessage)) > 0 Then SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    End If
                        
    ' Resource location
    rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).x
    rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).y

    ' The damage
    Damage = Item(GetPlayerEquipment(index, Weapon)).Damage
                    
    ' check if damage is more than health
    If Damage <= 0 Then
        ' too weak
        SendActionMsg GetPlayerMap(index), "Miss!", BrightRed, 1, (rX * 32), (rY * 32)
        Exit Sub
    End If
    
    ' cut it down!
    If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage <= 0 Then
        ' Message of resource health
        SendActionMsg GetPlayerMap(index), "-" & ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health, BrightRed, 1, (rX * 32), (rY * 32)
                                
        ' Cut resource!
        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceState = 1 ' Cut
        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).ResourceTimer = GetTickCount
        
        ' Update map resource
        SendResourceCacheToMap GetPlayerMap(index), Resource_num
        
        ' send message if it exists
        If Len(Trim$(Resource(Resource_index).SuccessMessage)) > 0 Then SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

        ' carry on
        GiveInvItem index, Resource(Resource_index).ItemReward, 1, True
        SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
    Else
        ' just do the damage
        ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_num).cur_health - Damage
        SendActionMsg GetPlayerMap(index), "-" & Damage, BrightRed, 1, (rX * 32), (rY * 32)
        SendAnimation GetPlayerMap(index), Resource(Resource_index).Animation, rX, rY
    End If
        
    ' send the sound
    SendMapSound index, rX, rY, SoundEntity.seResource, Resource_index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResource", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal invSlot As Long, ByVal amount As Long)
    Dim BankSlot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If invSlot < 0 Or invSlot > MAX_INV Then Exit Sub
    If amount < 0 Or amount > GetPlayerInvItemValue(index, invSlot) Then Exit Sub
    If amount < 1 Then amount = 1

    ' Find the slot of item
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, invSlot))
        
    ' Prevent subscript out range
    If BankSlot <= 0 Then Exit Sub

    ' Adds the item to the bank
    If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, invSlot) Then
        Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
        Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 1)
    Else
        Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, invSlot))
        Call SetPlayerBankItemValue(index, BankSlot, 1)
        Call TakeInvItem(index, GetPlayerInvItemNum(index, invSlot), 1)
    End If

    ' Update bank
    Call SaveBank(index)
    Call SavePlayer(index)
    Call SendBank(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GiveBankItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Long, ByVal amount As Long)
    Dim invSlot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If BankSlot < 0 Or BankSlot > MAX_BANK Then Exit Sub
    If amount < 0 Or amount > GetPlayerBankItemValue(index, BankSlot) Then Exit Sub
    If amount < 1 Then amount = 1

    ' Find the slot of item
    invSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    ' Prevent subscript out range
    If invSlot <= 0 Then Exit Sub
    
    ' Place the item in the bank
    Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), amount)
    Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - amount)
            
    If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
        Call SetPlayerBankItemNum(index, BankSlot, 0)
        Call SetPlayerBankItemValue(index, BankSlot, 0)
    End If
    
    ' Update bank
    Call SaveBank(index)
    Call SavePlayer(index)
    Call SendBank(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TakeBankItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub KillPlayer(ByVal index As Long)
    Dim EXP As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Calculate exp to give attacker
    EXP = GetPlayerExp(index) \ 3

    ' Make sure we dont get less then 0
    If EXP < 0 Then EXP = 0
    If EXP = 0 Then
        Call PlayerMsg(index, "You lost no exp.", BrightRed)
    Else
        Call SetPlayerExp(index, GetPlayerExp(index) - EXP)
        SendEXP index
        Call PlayerMsg(index, "You lost " & EXP & " exp.", BrightRed)
    End If
    
    Call OnDeath(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "KillPlayer", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UseItem(ByVal index As Long, ByVal invNum As Long)
    Dim n As Long, i As Long, tempItem As Long, x As Long, y As Long, ItemNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent hacking
    If invNum < 1 Or invNum > MAX_ITEMS Then Exit Sub
    If TempPlayer(index).InTrade > 0 Then Exit Sub

    ' The item
    ItemNum = GetPlayerInvItemNum(index, invNum)

    ' Prevent subscript out range
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
        
    ' stat requirements
    For i = 1 To Stats.Stat_Count - 1
        If GetPlayerRawStat(index, i) < Item(ItemNum).Stat_Req(i) Then
            PlayerMsg index, "You do not meet the stat requirements to equip this item.", BrightRed
            Exit Sub
        End If
    Next
                
    ' level requirement
    If GetPlayerLevel(index) < Item(ItemNum).LevelReq Then
        PlayerMsg index, "You do not meet the level requirement to equip this item.", BrightRed
        Exit Sub
    End If
                
    ' class requirement
    If Item(ItemNum).ClassReq > 0 Then
        If Not GetPlayerClass(index) = Item(ItemNum).ClassReq Then
            PlayerMsg index, "You do not meet the class requirement to equip this item.", BrightRed
            Exit Sub
        End If
    End If
                
    ' access requirement
    If Not GetPlayerAccess(index) >= Item(ItemNum).AccessReq Then
        PlayerMsg index, "You do not meet the access requirement to equip this item.", BrightRed
        Exit Sub
    End If

    ' Find out what kind of item it is
    Select Case Item(ItemNum).Type
        Case iArmor
            If GetPlayerEquipment(index, Armor) > 0 Then
                tempItem = GetPlayerEquipment(index, Armor)
            End If

            SetPlayerEquipment index, ItemNum, Armor
            PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
            TakeInvItem index, ItemNum, 1

            If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
            End If

            ' Update player equipment
            Call SendMapEquipment(index)
                
            ' send vitals
            Call SendVital(index, Vitals.HP)
            Call SendVital(index, Vitals.MP)
                
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
        Case iWeapon
            If GetPlayerEquipment(index, Weapon) > 0 Then
                tempItem = GetPlayerEquipment(index, Weapon)
            End If

            SetPlayerEquipment index, ItemNum, Weapon
            PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
            TakeInvItem index, ItemNum, 1

            If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
            End If

            ' Update player equipment
            Call SendMapEquipment(index)
                
            ' send vitals
            Call SendVital(index, Vitals.HP)
            Call SendVital(index, Vitals.MP)
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
        Case iHelmet
            If GetPlayerEquipment(index, Helmet) > 0 Then
                tempItem = GetPlayerEquipment(index, Helmet)
            End If

            SetPlayerEquipment index, ItemNum, Helmet
            PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
            TakeInvItem index, ItemNum, 1

            If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
            End If

            ' Update player equipment
            Call SendMapEquipment(index)
                
            ' send vitals
            Call SendVital(index, Vitals.HP)
            Call SendVital(index, Vitals.MP)
                
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
        Case iShield
            If GetPlayerEquipment(index, Shield) > 0 Then
                tempItem = GetPlayerEquipment(index, Shield)
            End If

            SetPlayerEquipment index, ItemNum, Shield
            PlayerMsg index, "You equip " & CheckGrammar(Item(ItemNum).Name), BrightGreen
            TakeInvItem index, ItemNum, 1

            If tempItem > 0 Then
                GiveInvItem index, tempItem, 0 ' give back the stored item
                tempItem = 0
            End If
                
            ' Update player equipment
            Call SendMapEquipment(index)
                
            ' send vitals
            Call SendVital(index, Vitals.HP)
            Call SendVital(index, Vitals.MP)
                
            ' send vitals to party if in one
            If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
            ' consumable
        Case iConsume
            ' add hp
            If Item(ItemNum).AddHP > 0 Then
                Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Item(ItemNum).AddHP)
                SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendVital index, Vitals.HP
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
                
            ' add mp
            If Item(ItemNum).AddMP > 0 Then
                Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Item(ItemNum).AddMP)
                SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendVital index, Vitals.MP
                ' send vitals to party if in one
                If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
            End If
                
            ' add exp
            If Item(ItemNum).AddEXP > 0 Then
                SetPlayerExp index, GetPlayerExp(index) + Item(ItemNum).AddEXP
                CheckPlayerLevelUp index
                SendActionMsg GetPlayerMap(index), "+" & Item(ItemNum).AddEXP & " EXP", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                SendEXP index
            End If
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call TakeInvItem(index, GetPlayerInvItemNum(index, invNum), 1)
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
        Case iKey
            ' Get the next tile
            Call GetPlayerNextTile(index, GetPlayerDir(index), x, y)
                
            ' Prevent subscript out range
            If GetPlayerY(index) <= 0 Or GetPlayerY(index) >= Map(GetPlayerMap(index)).MaxY Then Exit Sub
            If GetPlayerX(index) <= 0 Or GetPlayerX(index) >= Map(GetPlayerMap(index)).MaxX Then Exit Sub

            ' open door
            Call OpenDoor(index, x, y, ItemNum)
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
        Case iSpell
            ' Get the Spell num
            n = Item(ItemNum).Data1

            If n <= 0 Then
                Call PlayerMsg(index, "This Spell can only be learned by " & CheckGrammar(GetClassName(Spell(n).ClassReq)) & ".", BrightRed)
                Exit Sub
            End If
                
            ' Make sure they are the right class
            If Spell(n).ClassReq <> GetPlayerClass(index) And Spell(n).ClassReq > 0 Then
                Call PlayerMsg(index, "You must be class " & Spell(n).ClassReq & " to learn this skill.", BrightRed)
                Exit Sub
            End If
                
            ' Make sure they are the right level
            If Spell(n).LevelReq > GetPlayerLevel(index) Then
                Call PlayerMsg(index, "You must be level " & Spell(n).LevelReq & " to learn this skill.", BrightRed)
                Exit Sub
            End If
                
            i = FindOpenSpellSlot(index)

            ' Make sure they have an open Spell slot
            If i <= 0 Then
                Call PlayerMsg(index, "You cannot learn any more skills.", BrightRed)
                Exit Sub
            End If

            ' Make sure they dont already have the Spell
            If HasSpell(index, n) Then
                Call PlayerMsg(index, "You already have knowledge of this skill!.", BrightRed)
                Exit Sub
            End If

            ' Give spell to the player
            Call SetPlayerSpell(index, i, n)
            Call SendAnimation(GetPlayerMap(index), Item(ItemNum).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call PlayerMsg(index, "You feel the rush of knowledge fill your mind. You can now use " & Trim$(Spell(n).Name) & ".", BrightGreen)
                                    
            ' Take item
            Call TakeInvItem(index, ItemNum, 1)
                                    
            ' Update player
            Call SendPlayerSpells(index)
                                    
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
        Case iBag
            ' Search the last item of surprise
            For i = 1 To MAX_BAG
                If Item(ItemNum).BagItem(i) = 0 Then
                    n = i - 1
                    Exit For
                End If
            Next
                
            ' Prevent subscript out range
            If n > 0 Then Exit Sub
                
            ' Give item
            n = RAND(1, i)
            Call GiveInvItem(index, Item(ItemNum).BagItem(n), Item(ItemNum).BagValue(n))
            Call TakeInvItem(index, ItemNum, 0)
                
            ' send the sound
            Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, ItemNum)
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseItem", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CanBlockedTile(ByVal index As Long, ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Function
    
    ' Check to make sure that the tile is walkable
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_BLOCKED Then CanBlockedTile = True
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_RESOURCE Then CanBlockedTile = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanBlockedTile", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindOpenQuestSlot(ByVal index As Long) As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Function

    For i = 1 To MAX_PLAYER_QUESTS
        If GetPlayerQuestNum(index, i) = 0 Then
            FindOpenQuestSlot = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenQuestSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindQuestSlot(ByVal index As Long, ByVal QuestNum As Integer) As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Or QuestNum = 0 Then Exit Function

    For i = 1 To MAX_PLAYER_QUESTS
        If GetPlayerQuestNum(index, i) = QuestNum Then
            FindQuestSlot = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindQuestSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function QuestMaxTasks(ByVal QuestNum As Integer) As Byte
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_QUEST_TASKS
        If Quest(QuestNum).Task(i).Type = QUEST_TYPE_NONE Then
            QuestMaxTasks = i - 1
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "QuestMaxTasks", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CheckTitle(ByVal index As Long)
    Dim i As Integer, x As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_TITLES
        ' Check if title is valid
        If Len(Trim$(Title(i).Name)) <= 0 Or Title(i).Type <> TITLE_TYPE_NORMAL Then GoTo NextI
        
        ' Check if player has the requeriments
        If GetPlayerLevel(index) < Title(i).LevelReq Then GoTo NextI
        
        For x = 1 To Stats.Stat_Count - 1
            If GetPlayerStat(index, x) < Title(i).StatReq(x) Then GoTo NextI
        Next
        
        ' Check player has title
        If FindTitle(index, i) Then GoTo NextI
                                  
        ' Give player title
        Call SetPlayerTitle(index, FindOpenTitleSlot(index), i)
        Call PlayerMsg(index, "You gained the title " & Trim$(Title(i).Name), 15)
        
        ' Title sound
        Call SendPlayerSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seTitle, i)
NextI:
    Next
    
    ' Update player
    For x = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, x)
    Next

    Call SendPlayerData(index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTitle", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub UseTitle(ByVal index As Long, ByVal Slot As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    If GetPlayerTUsing(index) = GetPlayerTitle(index, Slot) Or GetPlayerTitle(index, Slot) = 0 Then Exit Sub

    ' Set the player title
    Call SetPlayerTUsing(index, GetPlayerTitle(index, Slot))
    Call PlayerMsg(index, "Now you' re using the title " & Trim$(Title(GetPlayerTitle(index, Slot)).Name), 15)
    
    ' Animation
    If Title(GetPlayerTitle(index, Slot)).UseAnimation > 0 Then Call SendAnimation(GetPlayerMap(index), Title(GetPlayerTitle(index, Slot)).UseAnimation, GetPlayerX(index), GetPlayerY(index))
    
    ' Update player
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next

    Call SendPlayerData(index)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UseTitle", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RemoveTitle(ByVal index As Long, ByVal Slot As Long)
    Dim i As Long
    Dim TitleNum As Integer
    Dim HotbarSlot As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Declaration
    TitleNum = GetPlayerTitle(index, Slot)
    
    ' Prevent subscript out range
    If TitleNum = 0 Then Exit Sub
    
    ' Remove using title
    If GetPlayerTUsing(index) = TitleNum Then Call SetPlayerTUsing(index, 0)
     
    ' Animation
    If Title(TitleNum).RemoveAnimation > 0 Then Call SendAnimation(GetPlayerMap(index), Title(TitleNum).RemoveAnimation, GetPlayerX(index), GetPlayerY(index))
    
    ' Remove player title
    Call PlayerMsg(index, "You removed the title " & Trim$(Title(TitleNum).Name), 15)
    Call SetPlayerTitle(index, Slot, 0)
    
    ' Remove item of hotbar
    HotbarSlot = FindHotbar(index, TitleNum, 1)
    If HotbarSlot > 0 Then
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).SType = 0
        Player(index).Char(TempPlayer(index).Char).Hotbar(HotbarSlot).Slot = 0
        SendHotbar index
    End If
    
    ' Update player
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
    
    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RemoveTitle", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RemoveTUsing(ByVal index As Integer)
    Dim i As Long
    Dim TitleNum As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Declaration
    TitleNum = GetPlayerTUsing(index)
    
    ' Prevent subscript out range
    If TitleNum = 0 Then Exit Sub
    
    ' Animation
    If Title(TitleNum).RemoveAnimation > 0 Then Call SendAnimation(GetPlayerMap(index), Title(TitleNum).RemoveAnimation, GetPlayerX(index), GetPlayerY(index))
    
    ' Remove player title
    Call PlayerMsg(index, "You are no longer using the title " & Trim$(Title(TitleNum).Name), 15)
    Call SetPlayerTUsing(index, 0)
    
    ' Update player
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next

    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RemoveTUsing", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetTitleStat(ByVal index As Integer, ByVal Stat As Stats, ByVal Value As Long) As Long
    Dim i As Long
    Dim TitleType As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetTitleStat = Value

    For i = 1 To MAX_PLAYER_TITLES
        ' Prevent subscript out of range
        If GetPlayerTitle(index, i) > 0 Then
            ' Check if title is passive
            If Title(GetPlayerTitle(index, i)).Passive = True Then
                GetTitleStat = Value + Title(GetPlayerTitle(index, i)).StatRew(Stat)
            End If
        End If
    Next
    
    ' Check if title not is passive
    If GetPlayerTUsing(index) > 0 Then
        GetTitleStat = Value + Title(GetPlayerTUsing(index)).StatRew(Stat)
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetTitleStat", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetTitleVital(ByVal index As Integer, ByVal Vital As Vitals, ByVal Value As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetTitleVital = Value

    For i = 1 To MAX_PLAYER_TITLES
        ' Prevent subscript out of range
        If GetPlayerTitle(index, i) > 0 Then
            ' Check if title is passive
            If Title(GetPlayerTitle(index, i)).Passive = True Then
                GetTitleVital = Value + Title(GetPlayerTitle(index, i)).VitalRew(Vital)
            End If
        End If
    Next
    
    ' Check if title not is passive
    If GetPlayerTUsing(index) > 0 Then
        GetTitleVital = Value + Title(GetPlayerTUsing(index)).VitalRew(Vital)
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetTitleVital", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerTitle(ByVal index As Long, ByVal Slot As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerTitle = Player(index).Char(TempPlayer(index).Char).Title.Title(Slot)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerTitle", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerTitle(ByVal index As Long, ByVal Slot As Long, ByVal Title As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Title.Title(Slot) = Title

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerTitle", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerTUsing(ByVal index As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function
    GetPlayerTUsing = Player(index).Char(TempPlayer(index).Char).Title.Using
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerTUsing", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SetPlayerTUsing(ByVal index As Long, ByVal Title As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Player(index).Char(TempPlayer(index).Char).Title.Using = Title
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetPlayerTUsing", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function FindOpenTitleSlot(ByVal index As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_TITLES
        If GetPlayerTitle(index, i) = 0 Then
            FindOpenTitleSlot = i
            Exit Function
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenTitleSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindTitle(ByVal index As Long, ByVal TitleNum As Long) As Boolean
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_PLAYER_TITLES
        If GetPlayerTitle(index, i) = TitleNum Then
            FindTitle = True
            Exit Function
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindTitle", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindTitleSlot(ByVal index As Long, ByVal TitleNum As Long) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_TITLES
        If GetPlayerTitle(index, i) = TitleNum Then
            FindTitleSlot = i
            Exit Function
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindTitleSlot", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindHotbar(ByVal index As Long, ByVal Num As Long, ByVal SType As Byte) As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_HOTBAR
        If Player(index).Char(TempPlayer(index).Char).Hotbar(i).SType = SType Then
            If Player(index).Char(TempPlayer(index).Char).Hotbar(i).Slot = Num Then
                FindHotbar = i
                Exit Function
            End If
        End If
    Next
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindHotbar", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub ClearTargets(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If PlayersOnMap(GetPlayerMap(index)) = NO Then Exit Sub
    
    For i = 1 To Player_HighIndex
        ' Prevent subscript out range
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                ' clear players target
                If TempPlayer(i).targetType = TARGET_TYPE_PLAYER Then
                    If TempPlayer(i).target = index Then
                        ' clear
                        TempPlayer(i).target = 0
                        TempPlayer(i).targetType = TARGET_TYPE_NONE
                        SendTarget i
                    End If
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTargets", "modPlayer", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
