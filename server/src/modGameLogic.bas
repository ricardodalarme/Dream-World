Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenPlayerSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS

        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If

    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindOpenMapItemSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "TotalOnlinePlayers", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then

            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If

    Next

    FindPlayer = 0

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindPlayer", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    Call SpawnItemSlot(i, ItemNum, ItemVal, MapNum, x, y)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal canDespawn As Boolean = True)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    i = MapItemSlot

    If i <> 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).canDespawn = canDespawn
            MapItem(MapNum, i).despawnTimer = GetTickCount + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).x = x
            MapItem(MapNum, i).y = y
            ' send to map
            SendSpawnItemToMap MapNum, i
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnItemSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnAllMapsItems", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
    Dim x As Long
    Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn what we have
    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            ' Check if the tile type is an item or a saved tile incase Soundeone drops Soundething
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then

                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Type = iCurrency And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If

        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnMapItems", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function Random(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Random = ((High - Low + 1) * Rnd) + Low

    ' Error handler
    Exit Function
errorhandler:
    HandleError "Random", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Dim NpcNum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Spawned As Boolean

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    ' Declartion
    NpcNum = Map(MapNum).Npc(MapNpcNum)

    If NpcNum > 0 Then
        MapNpc(MapNum).Npc(MapNpcNum).Num = NpcNum
        MapNpc(MapNum).Npc(MapNpcNum).target = 0
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 0 ' clear
        
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.MP) = GetNpcMaxVital(NpcNum, Vitals.MP)
        
        MapNpc(MapNum).Npc(MapNpcNum).Dir = Int(Rnd * 4)
        
        ' Check if theres a spawn tile for the specific npc
        For x = 0 To Map(MapNum).MaxX
            For y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(x, y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(x, y).Data1 = MapNpcNum Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).y = y
                        MapNpc(MapNum).Npc(MapNpcNum).Dir = Map(MapNum).Tile(x, y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next y
        Next x
        
        If Not Spawned Then
    
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                x = Random(0, Map(MapNum).MaxX)
                y = Random(0, Map(MapNum).MaxY)
    
                If x > Map(MapNum).MaxX Then x = Map(MapNum).MaxX
                If y > Map(MapNum).MaxY Then y = Map(MapNum).MaxY
    
                ' Check if the tile is walkable
                If NpcTileIsOpen(MapNum, x, y) Then
                    MapNpc(MapNum).Npc(MapNpcNum).x = x
                    MapNpc(MapNum).Npc(MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
    
            Next
            
        End If

        ' Didn' t spawn, so now we' ll just try to find a free tile
        If Not Spawned Then

            For x = 0 To Map(MapNum).MaxX
                For y = 0 To Map(MapNum).MaxY

                    If NpcTileIsOpen(MapNum, x, y) Then
                        MapNpc(MapNum).Npc(MapNpcNum).x = x
                        MapNpc(MapNum).Npc(MapNpcNum).y = y
                        Spawned = True
                    End If

                Next
            Next

        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set Buffer = New clsBuffer
            Buffer.WriteLong ServerPackets.SSpawnNpc
            Buffer.WriteLong MapNpcNum
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Num
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).x
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).y
            Buffer.WriteLong MapNpc(MapNum).Npc(MapNpcNum).Dir
            SendDataToMap MapNum, Buffer.ToArray()
            Set Buffer = Nothing
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnNpc", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function NpcTileIsOpen(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long) As Boolean
    Dim LoopI As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NpcTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex

            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = x Then
                    If GetPlayerY(LoopI) = y Then
                        NpcTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If

        Next

    End If

    For LoopI = 1 To MAX_MAP_NPCS

        If MapNpc(MapNum).Npc(LoopI).Num > 0 Then
            If MapNpc(MapNum).Npc(LoopI).x = x Then
                If MapNpc(MapNum).Npc(LoopI).y = y Then
                    NpcTileIsOpen = False
                    Exit Function
                End If
            End If
        End If

    Next

    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_ITEM Then
                NpcTileIsOpen = False
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "NpcTileIsOpen", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub SpawnMapNpcs(ByVal MapNum As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, MapNum)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnMapNpcs", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        Call SpawnMapNpcs(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpawnAllMapNpcs", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CanNpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim x As Long
    Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then
        Exit Function
    End If

    x = MapNpc(MapNum).Npc(MapNpcNum).x
    y = MapNpc(MapNum).Npc(MapNpcNum).y
    CanNpcMove = True

    Select Case Dir
        Case DIR_UP_LEFT
            ' Check to make sure not outside of boundries
            If y > 0 And x > 0 Then
                n = Map(MapNum).Tile(x - 1, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
               
        Case DIR_UP_RIGHT
            ' Check to make sure not outside of boundries
            If y > 0 And x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_UP + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_LEFT
            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY And x > 0 Then
                n = Map(MapNum).Tile(x - 1, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
                
        Case DIR_DOWN_RIGHT
            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY And x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                        
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_DOWN + 1) And isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If
        Case DIR_UP

            ' Check to make sure not outside of boundries
            If y > 0 Then
                n = Map(MapNum).Tile(x, y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_UP + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN

            ' Check to make sure not outside of boundries
            If y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(x, y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_DOWN + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT

            ' Check to make sure not outside of boundries
            If x > 0 Then
                n = Map(MapNum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_LEFT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT

            ' Check to make sure not outside of boundries
            If x < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(x + 1, y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (GetPlayerY(i) = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                            CanNpcMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (i <> MapNpcNum) And (MapNpc(MapNum).Npc(i).Num > 0) And (MapNpc(MapNum).Npc(i).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1) And (MapNpc(MapNum).Npc(i).y = MapNpc(MapNum).Npc(MapNpcNum).y) Then
                        CanNpcMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If isDirBlocked(Map(MapNum).Tile(MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y).DirBlock, DIR_RIGHT + 1) Then
                    CanNpcMove = False
                    Exit Function
                End If
            Else
                CanNpcMove = False
            End If

    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanNpcMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub NpcMove(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub

    ' Checks if the npc can move
    If Not CanNpcMove(MapNum, MapNpcNum, Dir) Then Exit Sub
    
    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
        Case DIR_UP_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
        Case DIR_DOWN_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
        Case DIR_DOWN_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
        Case DIR_UP
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
        Case DIR_DOWN
            MapNpc(MapNum).Npc(MapNpcNum).y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
        Case DIR_LEFT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
        Case DIR_RIGHT
            MapNpc(MapNum).Npc(MapNpcNum).x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
    End Select

    Call SendNpcXY(MapNum, MapNpcNum, Movement)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub NpcDir(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWN_RIGHT Then Exit Sub

    MapNpc(MapNum).Npc(MapNpcNum).Dir = Dir
    
    ' Check if have player on map
    If PlayersOnMap(MapNum) = NO Then Exit Sub

    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcDir
    Buffer.WriteLong MapNpcNum
    Buffer.WriteLong Dir
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcDir", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Long) As Long
    Dim i As Long
    Dim n As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    n = 0

    For i = 1 To Player_HighIndex

        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If

    Next

    GetTotalMapPlayers = n

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetTotalMapPlayers", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheResources(ByVal MapNum As Long)
    Dim x As Long, y As Long, Resource_Count As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Resource_Count = 0

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY

            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).x = x
                ResourceCache(MapNum).ResourceData(Resource_Count).y = y
                ResourceCache(MapNum).ResourceData(Resource_Count).cur_health = Resource(Map(MapNum).Tile(x, y).Data1).Health
            End If

        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchBankSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If
    
    OldNum = GetPlayerBankItemNum(index, oldSlot)
    OldValue = GetPlayerBankItemValue(index, oldSlot)
    NewNum = GetPlayerBankItemNum(index, newSlot)
    NewValue = GetPlayerBankItemValue(index, newSlot)
    
    SetPlayerBankItemNum index, newSlot, OldNum
    SetPlayerBankItemValue index, newSlot, OldValue
    
    SetPlayerBankItemNum index, oldSlot, NewNum
    SetPlayerBankItemValue index, oldSlot, NewValue
        
    SendBank index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSwitchBankSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchInvSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerInvItemNum(index, oldSlot)
    OldValue = GetPlayerInvItemValue(index, oldSlot)
    NewNum = GetPlayerInvItemNum(index, newSlot)
    NewValue = GetPlayerInvItemValue(index, newSlot)
    SetPlayerInvItemNum index, newSlot, OldNum
    SetPlayerInvItemValue index, newSlot, OldValue
    SetPlayerInvItemNum index, oldSlot, NewNum
    SetPlayerInvItemValue index, oldSlot, NewValue
    SendInventory index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSwitchInvSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchSpellSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If oldSlot = 0 Or newSlot = 0 Then
        Exit Sub
    End If

    OldNum = GetPlayerSpell(index, oldSlot)
    NewNum = GetPlayerSpell(index, newSlot)
    SetPlayerSpell index, oldSlot, NewNum
    SetPlayerSpell index, newSlot, OldNum
    SendPlayerSpells index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSwitchSpellSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerUnequipItem(ByVal index As Long, ByVal EqSlot As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EqSlot <= 0 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub ' exit out early if error' d
    If FindOpenInvSlot(index, GetPlayerEquipment(index, EqSlot)) > 0 Then
        GiveInvItem index, GetPlayerEquipment(index, EqSlot), 0
        PlayerMsg index, "You unequip " & CheckGrammar(Item(GetPlayerEquipment(index, EqSlot)).Name), Yellow
        ' send the sound
        SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerEquipment(index, EqSlot)
        ' remove equipment
        SetPlayerEquipment index, 0, EqSlot
        SendMapEquipment index
        SendStats index
        ' send vitals
        Call SendVital(index, Vitals.HP)
        Call SendVital(index, Vitals.MP)
        ' send vitals to party if in one
        If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
    Else
        PlayerMsg index, "Your inventory is full.", BrightRed
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerUnequipItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
    Dim FirstLetter As String * 1
   
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
        CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
        Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CheckGrammar", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function isInRange(ByVal Range As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Dim nVal As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    isInRange = False
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    If nVal <= Range Then isInRange = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isInRange", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function isDirBlocked(ByRef blockvar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not blockvar And (2 ^ Dir) Then
        isDirBlocked = False
    Else
        isDirBlocked = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function RAND(ByVal Low As Long, ByVal High As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Randomize
    RAND = Int((High - Low + 1) * Rnd) + Low

    ' Error handler
    Exit Function
errorhandler:
    HandleError "RAND", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' #####################
' ## Party functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal index As Long)
    Dim partyNum As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    partyNum = TempPlayer(index).inParty
    If partyNum > 0 Then
        ' find out how many members we have
        Party_CountMembers partyNum
        ' make sure there' s more than 2 people
        If Party(partyNum).MemberCount > 2 Then
            ' check if leader
            If Party(partyNum).Leader = index Then
                ' set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) > 0 And Party(partyNum).Member(i) <> index Then
                        Party(partyNum).Leader = Party(partyNum).Member(i)
                        PartyMsg partyNum, GetPlayerName(Party(partyNum).Leader) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                ' leave party
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                TempPlayer(index).inParty = 0
                TempPlayer(index).partyInvite = 0
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            Else
                ' not the leader, just leave
                PartyMsg partyNum, GetPlayerName(index) & " has left the party.", BrightRed
                ' remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(i) = index Then
                        Party(partyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                TempPlayer(index).inParty = 0
                TempPlayer(index).partyInvite = 0
                ' recount party
                Party_CountMembers partyNum
                ' set update to all
                SendPartyUpdate partyNum
                ' send clear to player
                SendPartyUpdateTo index
            End If
        Else
            ' find out how many members we have
            Party_CountMembers partyNum
            ' only 2 people, disband
            PartyMsg partyNum, "Party disbanded.", BrightRed
            ' clear out everyone' s party
            For i = 1 To MAX_PARTY_MEMBERS
                index = Party(partyNum).Member(i)
                ' player exist?
                If index > 0 Then
                    ' remove them
                    TempPlayer(index).inParty = 0
                    ' send clear to players
                    SendPartyUpdateTo index
                End If
            Next
            ' clear out the party itself
            ClearParty partyNum
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Party_PlayerLeave", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_Invite(ByVal index As Long, ByVal targetPlayer As Long)
    Dim partyNum As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check if the person is a valid target
    If Not IsConnected(targetPlayer) Or Not IsPlaying(targetPlayer) Then Exit Sub
    
    ' make sure they' re not busy
    If TempPlayer(targetPlayer).partyInvite > 0 Or TempPlayer(targetPlayer).TradeRequest > 0 Then
        ' they' ve already got a request for trade/party
        PlayerMsg index, "This player is busy.", BrightRed
        ' exit out early
        Exit Sub
    End If
    ' make syure they' re not in a party
    If TempPlayer(targetPlayer).inParty > 0 Then
        ' they' re already in a party
        PlayerMsg index, "This player is already in a party.", BrightRed
        ' exit out early
        Exit Sub
    End If
    
    ' check if we' re in a party
    If TempPlayer(index).inParty > 0 Then
        partyNum = TempPlayer(index).inParty
        ' make sure we' re the leader
        If Party(partyNum).Leader = index Then
            ' got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(partyNum).Member(i) = 0 Then
                    ' send the invitation
                    SendDialogue targetPlayer, "Party Invitation", GetPlayerName(index) & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, YES
                    ' set the invite target
                    TempPlayer(targetPlayer).partyInvite = index
                    ' let them know
                    PlayerMsg index, "Invitation sent.", Pink
                    Exit Sub
                End If
            Next
            ' no room
            PlayerMsg index, "Party is full.", BrightRed
            Exit Sub
        Else
            ' not the leader
            PlayerMsg index, "You are not the party leader.", BrightRed
            Exit Sub
        End If
    Else
        ' not in a party - doesn' t matter!
        SendDialogue targetPlayer, "Party Invitation", GetPlayerName(index) & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, YES
        ' set the invite target
        TempPlayer(targetPlayer).partyInvite = index
        ' let them know
        PlayerMsg index, "Invitation sent.", Pink
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Party_Invite", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_InviteAccept(ByVal index As Long, ByVal targetPlayer As Long)
    Dim partyNum As Long, i As Long, x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check if already in a party
    If TempPlayer(index).inParty > 0 Then
        ' get the partynumber
        partyNum = TempPlayer(index).inParty
        ' got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(partyNum).Member(i) = 0 Then
                ' add to the party
                Party(partyNum).Member(i) = targetPlayer
                ' recount party
                Party_CountMembers partyNum
                ' send update to all - including new player
                SendPartyUpdate partyNum
                For x = 1 To MAX_PARTY_MEMBERS
                    If Party(partyNum).Member(x) > 0 Then
                        SendPartyVitals partyNum, Party(partyNum).Member(x)
                    End If
                Next
                ' let everyone know they' ve joined
                PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
                ' add them in
                TempPlayer(targetPlayer).inParty = partyNum
                Exit Sub
            End If
        Next
        ' no empty slots - let them know
        PlayerMsg index, "Party is full.", BrightRed
        PlayerMsg targetPlayer, "Party is full.", BrightRed
        TempPlayer(targetPlayer).partyInvite = 0
        Exit Sub
    Else
        ' not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' find blank party
            If Not Party(i).Leader > 0 Then
                partyNum = i
                Exit For
            End If
        Next
        ' create the party
        Party(partyNum).MemberCount = 2
        Party(partyNum).Leader = index
        Party(partyNum).Member(1) = index
        Party(partyNum).Member(2) = targetPlayer
        SendPartyUpdate partyNum
        SendPartyVitals partyNum, index
        SendPartyVitals partyNum, targetPlayer
        ' let them know it' s created
        PartyMsg partyNum, "Party created.", BrightGreen
        PartyMsg partyNum, GetPlayerName(index) & " has joined the party.", Pink
        PartyMsg partyNum, GetPlayerName(targetPlayer) & " has joined the party.", Pink
        ' clear the invitation
        TempPlayer(targetPlayer).partyInvite = 0
        ' add them to the party
        TempPlayer(index).inParty = partyNum
        TempPlayer(targetPlayer).inParty = partyNum
        Exit Sub
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Party_InviteAccept", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_InviteDecline(ByVal index As Long, ByVal targetPlayer As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    PlayerMsg index, GetPlayerName(targetPlayer) & " has declined to join the party.", BrightRed
    PlayerMsg targetPlayer, "You declined to join the party.", BrightRed
    ' clear the invitation
    TempPlayer(targetPlayer).partyInvite = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Party_InviteDecline", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_CountMembers(ByVal partyNum As Long)
    Dim i As Long, highIndex As Long, x As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' find the high index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(partyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    ' count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' we' ve got a blank member
        If Party(partyNum).Member(i) = 0 Then
            ' is it lower than the high index?
            If i < highIndex Then
                ' move everyone down a slot
                For x = i To MAX_PARTY_MEMBERS - 1
                    Party(partyNum).Member(x) = Party(partyNum).Member(x + 1)
                    Party(partyNum).Member(x + 1) = 0
                Next x
            Else
                ' not lower - highindex is count
                Party(partyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        ' check if we' ve reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(partyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    ' if we' re here it means that we need to re-count again
    Party_CountMembers partyNum

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Party_CountMembers", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Party_ShareExp(ByVal partyNum As Long, ByVal EXP As Long, ByVal index As Long)
    Dim expShare As Long, leftOver As Long, i As Long, tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Party(partyNum).MemberCount <= 0 Then Exit Sub
    
    ' check if it' s worth sharing
    If Not EXP >= Party(partyNum).MemberCount Then
        ' no party - keep exp for self
        GivePlayerEXP index, EXP
        Exit Sub
    End If
    
    ' find out the equal share
    expShare = EXP \ Party(partyNum).MemberCount
    leftOver = EXP Mod Party(partyNum).MemberCount
    
    ' loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(partyNum).Member(i)
        ' existing member?Kn
        If tmpIndex > 0 Then
            ' playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' give them their share
                GivePlayerEXP tmpIndex, expShare
            End If
        End If
    Next
    
    ' give the remainder to a random member
    tmpIndex = Party(partyNum).Member(RAND(1, Party(partyNum).MemberCount))
    ' give the exp
    GivePlayerEXP tmpIndex, leftOver

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Party_ShareExp", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GivePlayerEXP(ByVal index As Long, ByVal EXP As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' give the exp
    Call SetPlayerExp(index, GetPlayerExp(index) + EXP)
    SendEXP index
    SendActionMsg GetPlayerMap(index), "+" & EXP & " EXP", White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    ' check if we' ve leveled
    CheckPlayerLevelUp index

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GivePlayerEXP", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function IsInventoryFull(ByVal tradeTarget As Long, ByVal index As Long) As Boolean
    Dim InvEmpty As Long, TradeFull As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(tradeTarget, i) > 0 And GetPlayerInvItemNum(tradeTarget, i) <= MAX_ITEMS Then
            InvEmpty = InvEmpty + 1
        End If
    Next
        
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num > 0 And TempPlayer(index).TradeOffer(i).Num <= MAX_ITEMS Then
            TradeFull = TradeFull + 1
        End If
    Next
        
    If TradeFull > (MAX_INV - InvEmpty) Then
        IsInventoryFull = True
        Exit Function
    End If
    
    IsInventoryFull = False

    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsInventoryFull", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheDoors(ByVal MapNum As Long)
    Dim x As Long, y As Long, Door_Count As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Door_Count = 0

    For x = 0 To Map(MapNum).MaxX
        For y = 0 To Map(MapNum).MaxY
            If Map(MapNum).Tile(x, y).Type = TILE_TYPE_DOOR Then
                Door_Count = Door_Count + 1
                ReDim Preserve DoorCache(MapNum).Data(Door_Count)
                DoorCache(MapNum).Data(Door_Count).x = x
                DoorCache(MapNum).Data(Door_Count).y = y
            End If

        Next
    Next

    DoorCache(MapNum).Count = Door_Count

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheDoors", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub OpenDoor(ByVal index As Integer, ByVal x As Integer, ByVal y As Integer, Optional ByVal Key As Integer = 0)
    Dim i As Integer
    Dim DoorNum As Integer, MapDoor As Integer
    Dim MapNum As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Player localization
    MapNum = GetPlayerMap(index)
    
    ' Prevent subscript out range
    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_DOOR Then Exit Sub
    
    ' Door num
    DoorNum = Map(MapNum).Tile(x, y).Data1
    
    ' Map door num
    For i = 1 To DoorCache(MapNum).Count
        If DoorCache(MapNum).Data(i).x = x Then
            If DoorCache(MapNum).Data(i).y = y Then
                MapDoor = i
                Exit For
            End If
        End If
    Next
    
    ' Prevent subscript out range
    If MapDoor <= 0 Then Exit Sub

    ' Check if player have key
    If Door(DoorNum).OpenWith > 0 Then
        ' Prevent subscript out range
        If Key <= 0 Then Exit Sub
            
        If Door(DoorNum).OpenWith <> Key Then ' closed
            Call PlayerMsg(index, "The door is not open.", Red)
            Exit Sub
        End If
    End If
    
    ' open the door
    If DoorCache(MapNum).Data(MapDoor).State = 0 Then ' closed
        ' open
        DoorCache(MapNum).Data(MapDoor).State = 1 ' opening
        DoorCache(MapNum).Data(MapDoor).RespawnTime = GetTickCount
        Call SendDoorCacheToMap(MapNum, MapDoor)
    
        ' Invite the sound and animation
        Call SendMapSound(index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seDoor, DoorNum)
        Call SendAnimation(MapNum, Door(DoorNum).Animation, GetPlayerX(index), GetPlayerY(index))
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "OpenDoor", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckDoor(ByVal index As Integer, ByVal x As Byte, ByVal y As Byte)
    Dim i As Integer
    Dim MapNum As Integer
    Dim DoorNum As Integer, MapDoor As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Player map
    MapNum = GetPlayerMap(index)
    
    ' Prevent subscript out range
    If Map(MapNum).Tile(x, y).Type <> TILE_TYPE_DOOR Then Exit Sub
    
    ' Door num
    DoorNum = Map(MapNum).Tile(x, y).Data1
    
    ' Map door num
    For i = 1 To DoorCache(MapNum).Count
        If DoorCache(MapNum).Data(i).x = x Then
            If DoorCache(MapNum).Data(i).y = y Then
                MapDoor = i
                Exit For
            End If
        End If
    Next
    
    ' Prevent subscript out range
    If MapDoor <= 0 Then Exit Sub
    
    ' Opening?
    If DoorCache(MapNum).Data(MapDoor).State <> 1 Then Exit Sub
    
    ' Check if player have the requirements
    If Door(DoorNum).LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You do not meet the level requirement to enter in door.", Red)
        Exit Sub
    End If
    
    For i = 1 To Stats.Stat_Count - 1
        If Door(DoorNum).Stat_Req(i) > GetPlayerStat(index, i) Then
            Call PlayerMsg(index, "You do not meet the stat requirement to enter in door!", Red)
            Exit Sub
        End If
    Next
    
    ' Checks if the door will teleport the player
    If Door(DoorNum).Map > 0 Then
        Call PlayerWarp(index, Door(DoorNum).Map, Door(DoorNum).x, Door(DoorNum).y)
        Call SetPlayerDir(index, Door(DoorNum).Dir)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckDoor", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetStat(ByVal Stat As Stats) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with stat name
    Select Case Stat
        Case Stats.Strength
            GetStat = "Strength"
        Case Stats.Endurance
            GetStat = "Endurance"
        Case Stats.Intelligence
            GetStat = "Intelligence"
        Case Stats.Agility
            GetStat = "Agility"
        Case Stats.Willpower
            GetStat = "Willpower"
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetStat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetVital(ByVal Vital As Vitals) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with vital name
    Select Case Vital
        Case Vitals.HP
            GetVital = "HP"
        Case Vitals.MP
            GetVital = "MP"
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetVital", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetQuestTypeTwo(ByVal Quest_Type As Byte) As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return with QuestType name
    Select Case Quest_Type
        Case QUEST_TYPE_NONE
            GetQuestTypeTwo = "None"
        Case QUEST_TYPE_KILLNPC
            GetQuestTypeTwo = "Kill "
        Case QUEST_TYPE_KILLPLAYER
            GetQuestTypeTwo = "Kill player"
        Case QUEST_TYPE_GOTOMAP
            GetQuestTypeTwo = "Go to map "
        Case QUEST_TYPE_TALKNPC
            GetQuestTypeTwo = "Talk to "
        Case QUEST_TYPE_COLLECTITEMS
            GetQuestTypeTwo = "Collect "
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetQuestTypeTwo", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub QuestAccept(ByVal index As Integer)
    Dim QuestNum As Integer, QuestSlot As Byte, QuestSlotFind As Byte, ActualTask As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Declaration
    QuestNum = TempPlayer(index).QuestInvite
    
    ' Prevent subscript out range
    If QuestNum = 0 Then Exit Sub

    ' Declarations
    QuestSlot = FindOpenQuestSlot(index)
    QuestSlotFind = FindQuestSlot(index, QuestNum)
    
    ' Checks that have already completed this quest
    If QuestSlotFind > 0 Then
        If Quest(QuestNum).Retry = True And GetPlayerQuestStatus(index, QuestSlotFind) = QUEST_STATUS_END Then
            QuestSlot = QuestSlotFind
        End If
    End If

    ' Prevent subscript out range
    If QuestSlot = 0 Then Exit Sub

    ' Start quest
    Call SetPlayerQuestNum(index, QuestSlot, QuestNum)
    Call SetPlayerQuestStatus(index, QuestSlot, QUEST_STATUS_STARTING)
    Call SetPlayerQuestPart(index, QuestSlot, 1)
    
    ' Declaration
    ActualTask = GetPlayerQuestPart(index, QuestSlot)

    With Quest(QuestNum).Task(ActualTask)
        ' Reset the status of the old quest
        If .Type = QUEST_TYPE_KILLNPC Then Call SetPlayerKillNpcs(index, .Num, 0)
        If .Type = QUEST_TYPE_KILLPLAYER Then Call SetPlayerKillPlayers(index, 0)

        ' Message
        If Len(Trim$(.Message(1))) > 0 Then Call PlayerMsg(index, Trim$(.Message(1)), Yellow)
    End With
    
    ' Check now completed the quest
    Select Case Quest(QuestNum).Task(ActualTask).Type
            ' Go to map
        Case QUEST_TYPE_GOTOMAP
            Call CheckCompleteQuest(index, GetPlayerMap(index))
            ' Others
        Case QUEST_TYPE_KILLNPC, QUEST_TYPE_KILLPLAYER, QUEST_TYPE_COLLECTITEMS
            Call CheckCompleteQuest(index, , True)
    End Select

    ' Update player
    Call SendPlayerData(index)
    TempPlayer(index).QuestInvite = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "QuestAccept", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckQuest(ByVal index As Long, ByVal NpcNum As Integer, ByVal Slot As Byte)
    Dim QuestNum As Integer, QuestSlot As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Slot <= 0 Then Exit Sub
    
    ' Declaration
    QuestNum = Npc(NpcNum).Quest(Slot)
    
    ' Prevent subscript out range
    If Not IsPlaying(index) Or QuestNum <= 0 Then Exit Sub
    
    ' Declaration
    QuestSlot = FindQuestSlot(index, QuestNum)

    ' Checks can start the quest
    If QuestSlot > 0 Then
        ' Checks can complete the quest
        If GetPlayerQuestStatus(index, QuestSlot) = QUEST_STATUS_COMPLETE Or GetPlayerQuestStatus(index, QuestSlot) = QUEST_STATUS_STARTING Then
            Call CheckEndQuest(index, QuestNum)
            Exit Sub
        End If
        
        ' Checks can end the quest, if yes checking whether it can redo
        If GetPlayerQuestStatus(index, QuestSlot) = QUEST_STATUS_END And Quest(QuestNum).Retry = False Then Exit Sub
    End If
    
    ' Check if player can start quest
    Call CheckStartQuest(index, QuestNum)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckQuest", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckStartQuest(ByVal index As Long, ByVal QuestNum As Integer)
    Dim i As Byte
    Dim QuestSlot As Byte, QuestSlotFind As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Or QuestNum <= 0 Then Exit Sub
    
    ' Declarations
    QuestSlot = FindOpenQuestSlot(index)
    QuestSlotFind = FindQuestSlot(index, QuestNum)
    
    ' Check if quest is complete
    If QuestSlotFind > 0 Then
        If GetPlayerQuestStatus(index, QuestSlotFind) = QUEST_STATUS_END Then
            QuestSlot = QuestSlotFind
            QuestNum = GetPlayerQuestNum(index, QuestSlot)
        End If
    End If

    ' Prevent subscript out range
    If QuestSlot = 0 Or QuestNum = 0 Then Exit Sub

    ' Can start quest?
    With Quest(QuestNum)
        ' Level required
        If GetPlayerLevel(index) < .LevelReq Then
            Call PlayerMsg(index, "You need the level " & .LevelReq & ", or more, to start this quest!", Red)
            Exit Sub
        End If
        
        ' Stats required
        For i = 1 To Stats.Stat_Count - 1
            If GetPlayerStat(index, i) < .StatReq(i) Then
                Call PlayerMsg(index, "You do not have the stats needed to start this quest!", Red)
                Exit Sub
            End If
        Next
        
        ' Quest required
        If .QuestReq > 0 Then
            ' Declaration
            QuestSlotFind = FindQuestSlot(index, Quest(QuestNum).QuestReq)

            If QuestSlotFind > 0 Then
                If GetPlayerQuestStatus(index, QuestSlotFind) <> QUEST_STATUS_END Then
                    Call PlayerMsg(index, "You need the quest " & Trim$(Quest(.QuestReq).Name) & " to start this quest!", Red)
                    Exit Sub
                End If
            End If
        End If

        ' Class required
        If .ClassReq > 0 Then
            If GetPlayerClass(index) <> .ClassReq Then
                Call PlayerMsg(index, "You need the class " & Trim$(Class(.ClassReq).Name) & " to start this quest!", Red)
                Exit Sub
            End If
        End If
        
        ' Sprite required
        If .SpriteReq > 0 Then
            If GetPlayerSprite(index) <> .SpriteReq Then
                Call PlayerMsg(index, "You do not have the sprite needed to start this quest!", Red)
                Exit Sub
            End If
        End If

        ' Send the invitation of the quest
        TempPlayer(index).QuestInvite = QuestNum
        Call SendDialogue(index, Trim$(.Name), Trim$(.Description), DIALOGUE_TYPE_QUEST, YES)
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckStartQuest", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckCompleteQuest(ByVal index As Long, Optional ByVal Num As Integer = 0, Optional ByVal Message As Boolean = False)
    Dim i As Integer
    Dim QuestNum As Integer, QuestStatus As Byte
    Dim ActualTask As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Checks if the player completed a quest
    For i = 1 To MAX_PLAYER_QUESTS
        ' Declarations
        QuestNum = GetPlayerQuestNum(index, i)
        QuestStatus = GetPlayerQuestStatus(index, i)
        ActualTask = GetPlayerQuestPart(index, i)
    
        ' Prevent subscript out range
        If QuestNum > 0 Then
            If QuestStatus = QUEST_STATUS_STARTING Or QuestStatus = QUEST_STATUS_COMPLETE Then
                ' Checks if the player managed to complete the quest
                With Quest(QuestNum).Task(ActualTask)
                    ' Killing npcs
                    If .Type = QUEST_TYPE_KILLNPC And .Num > 0 Then
                        If GetPlayerKillNpcs(index, .Num) > .Value Then: GoTo Continue
                        If Message Then Call PlayerMsg(index, "[" & Trim$(Quest(QuestNum).Name) & "] " & Trim$(Npc(.Num).Name) & " - " & GetPlayerKillNpcs(index, .Num) & "/" & .Value, White)
                        If GetPlayerKillNpcs(index, .Num) < .Value Then Call SetPlayerQuestStatus(index, i, QUEST_STATUS_STARTING): GoTo Continue
                        ' Killing Players
                    ElseIf .Type = QUEST_TYPE_KILLPLAYER And .Value > 0 Then
                        If GetPlayerKillPlayers(index) > .Value Then GoTo Continue
                        If Message Then Call PlayerMsg(index, "[" & Trim$(Quest(QuestNum).Name) & "] " & GetPlayerKillPlayers(index) & "/" & .Value, White)
                        If GetPlayerKillPlayers(index) < .Value Then Call SetPlayerQuestStatus(index, i, QUEST_STATUS_STARTING): GoTo Continue
                        ' Go to a map
                    ElseIf .Type = QUEST_TYPE_GOTOMAP And .Num > 0 Then
                        If Num <> .Num Then GoTo Continue
                        ' Conversation with npcs
                    ElseIf .Type = QUEST_TYPE_TALKNPC And .Num > 0 Then
                        If Num <> .Num Then GoTo Continue
                        ' Collect items
                    ElseIf .Type = QUEST_TYPE_COLLECTITEMS And .Num > 0 Then
                        If HasItem(index, .Num) > .Value Then GoTo Continue
                        If Message Then Call PlayerMsg(index, "[" & Trim$(Quest(QuestNum).Name) & "] " & Trim$(Item(.Num).Name) & " - " & HasItem(index, .Num) & "/" & .Value, White)
                        If HasItem(index, .Num) < .Value Then Call SetPlayerQuestStatus(index, i, QUEST_STATUS_STARTING): GoTo Continue
                    End If
                    
                    ' Checks if the mission is completed
                    If QuestStatus = QUEST_STATUS_COMPLETE Then GoTo Continue
                    
                    ' Complete the quest
                    Call SetPlayerQuestStatus(index, i, QUEST_STATUS_COMPLETE)
                    Call PlayerMsg(index, "The quest " & Trim$(Quest(QuestNum).Name) & " was completed.", White)
                    
                    ' Checks if the mission is completed instantly
                    If .Instant = True Then Call CheckEndQuest(index, QuestNum)
                End With
Continue:
            End If
        End If
    Next

    ' Update player
    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCompleteQuest", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CheckEndQuest(ByVal index As Long, ByVal QuestNum As Integer)
    Dim QuestSlot As Byte, ActualTask As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Or QuestNum <= 0 Then Exit Sub

    ' Declarations
    QuestSlot = FindQuestSlot(index, QuestNum)
    ActualTask = GetPlayerQuestPart(index, QuestSlot)
    
    With Quest(QuestNum).Task(ActualTask)
        ' Check if the mission is in progress
        If GetPlayerQuestStatus(index, QuestSlot) <> QUEST_STATUS_COMPLETE Then
            If Len(Trim$(.Message(2))) > 0 Then Call PlayerMsg(index, Trim$(.Message(2)), Red)
            Exit Sub
        End If
        
        SetPlayerQuestStatus index, QuestSlot, QUEST_STATUS_STARTING
        
        ' Reset the status of quest
        If .Type = QUEST_TYPE_NONE Then Exit Sub
        If .Type = QUEST_TYPE_KILLNPC And .Num > 0 Then Call SetPlayerKillNpcs(index, .Num, 0)
        If .Type = QUEST_TYPE_KILLPLAYER And .Value > 0 Then Call SetPlayerKillPlayers(index, 0)
        If .Type = QUEST_TYPE_COLLECTITEMS And .Num > 0 Then Call TakeInvItem(index, .Num, .Value, False)
                
        ' Message
        If Len(Trim$(.Message(3))) > 0 Then Call PlayerMsg(index, Trim$(.Message(3)), Yellow)
        
        ' Check if the player has fulfilled all tasks if yes finish the mission
        If ActualTask >= QuestMaxTasks(QuestNum) Then
            Call EndQuest(index, QuestNum)
            Exit Sub
        End If
        
        ' Continue the quest
        Call SetPlayerQuestPart(index, QuestSlot, ActualTask + 1)
    End With

    ' Update player
    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckEndQuest", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EndQuest(ByVal index As Long, ByVal QuestNum As Integer)
    Dim i As Byte
    Dim QuestSlot As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not IsPlaying(index) Or QuestNum <= 0 Then Exit Sub

    ' Declarations
    QuestSlot = FindQuestSlot(index, QuestNum)

    With Quest(QuestNum)
        ' End quest
        Call PlayerMsg(index, "You finished the quest " & Trim$(Quest(QuestNum).Name), White)
        Call SetPlayerQuestStatus(index, QuestSlot, QUEST_STATUS_END)
        Call SetPlayerQuestPart(index, QuestSlot, 0)

        ' Rewards
        Call SetPlayerLevel(index, GetPlayerLevel(index) + .LevelRew)
        Call SetPlayerExp(index, GetPlayerExp(index) + .ExpRew)
        Call CheckPlayerLevelUp(index)
        
        For i = 1 To Stats.Stat_Count - 1
            Call SetPlayerStat(index, i, GetPlayerStat(index, i) + .StatRew(i))
        Next
        
        For i = 1 To Vitals.Vital_Count - 1
            Call SetPlayerVital(index, i, GetPlayerVital(index, i) + .VitalRew(i))
        Next
        
        If .ClassRew > 0 Then Call SetPlayerClass(index, .ClassRew)
        If .SpriteRew > 0 Then Call SetPlayerSprite(index, .SpriteRew)
    End With

    ' Update player
    Call SendPlayerData(index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EndQuest", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub PlayerSwitchTitleSlots(ByVal index As Long, ByVal oldSlot As Long, ByVal newSlot As Long)
    Dim OldNum As Long
    Dim NewNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If oldSlot = 0 Or newSlot = 0 Then Exit Sub

    ' Titles slot
    OldNum = GetPlayerTitle(index, oldSlot)
    NewNum = GetPlayerTitle(index, newSlot)
    
    ' Move titles
    SetPlayerTitle index, oldSlot, NewNum
    SetPlayerTitle index, newSlot, OldNum
    
    ' Update player
    SendPlayerTitles index
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerSwitchTitleSlots", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GetPlayerNextTile(ByVal index As Long, ByVal Dir As Byte, ByRef x As Long, ByRef y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out range
    If Not IsPlaying(index) Then Exit Sub
    
    ' Return values
    Select Case Dir
        Case DIR_UP
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1
        Case DIR_DOWN
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1
        Case DIR_LEFT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)
        Case DIR_RIGHT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)
        Case DIR_UP_RIGHT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) + 1
        Case DIR_UP_LEFT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) + 1
        Case DIR_DOWN_RIGHT
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index) - 1
        Case DIR_DOWN_LEFT
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index) - 1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetPlayerNextTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub GetNpcNextTile(ByVal MapNum As Long, ByVal MapNpcNum As Long, ByVal Dir As Byte, ByRef x As Long, ByRef y As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out range
    If MapNpc(MapNum).Npc(MapNpcNum).Num <= 0 Or MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Sub
    
    ' Return values
    Select Case Dir
        Case DIR_UP
            x = MapNpc(MapNum).Npc(MapNpcNum).x
            y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
        Case DIR_DOWN
            x = MapNpc(MapNum).Npc(MapNpcNum).x
            y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
        Case DIR_LEFT
            x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
            y = MapNpc(MapNum).Npc(MapNpcNum).y
        Case DIR_RIGHT
            x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
            y = MapNpc(MapNum).Npc(MapNpcNum).y
        Case DIR_UP_RIGHT
            x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
            y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
        Case DIR_UP_LEFT
            x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
            y = MapNpc(MapNum).Npc(MapNpcNum).y + 1
        Case DIR_DOWN_RIGHT
            x = MapNpc(MapNum).Npc(MapNpcNum).x - 1
            y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
        Case DIR_DOWN_LEFT
            x = MapNpc(MapNum).Npc(MapNpcNum).x + 1
            y = MapNpc(MapNum).Npc(MapNpcNum).y - 1
    End Select
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "GetNpcNextTile", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub CreateChat(ByVal Name As String, ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Find chats with similar name
    If FindChatRoom(Name) > 0 Then
        PlayerMsg index, "Chat room with given name already exists.", Red
        Exit Sub
    End If

    ' Find a room empty
    i = FindChatRoom

    ' Prevent subscript out range
    If i > 0 Then
        PlayerMsg index, "No more empty rooms.", Red
        Exit Sub
    End If

    ' Add room
    ChatRoom(i).index = i
    ChatRoom(i).Name = Name
    ChatRoom(i).Members = 1
    
    ' Adds the player to room
    TempPlayer(index).roomIndex = i
    PlayerMsg index, "Your chat room '" & Name & "' has been created.", Blue

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateChat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function FindChatRoom(Optional ByVal Name As String = vbNullString) As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_ROOMS
        If ChatRoom(i).Name = Name Then
            FindChatRoom = i
            Exit Function
        End If
    Next

    ' Error handler
    Exit Function
errorhandler:
    HandleError "FindChatRoom", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub JoinChat(ByVal Name As String, ByVal index As Long)
    Dim i As Long
    Dim ChatSlot As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Find the chat room
    ChatSlot = FindChatRoom(Name)
    
    ' Checks if this chat room exists
    If Not FindChatRoom(Name) Then
        PlayerMsg index, "The chat room '" & Name & "' doesn't exist.", AlertColor
        Exit Sub
    End If

    ' Adds the player in the chat room
    ChatRoom(ChatSlot).Members = ChatRoom(ChatSlot).Members + 1
    TempPlayer(index).roomIndex = i
    PlayerMsg index, "Joined chat room '" & Name & "' successfully.", Blue

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "JoinChat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub RemoveChat(ByVal index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With ChatRoom(index)
        .index = 0
        .Members = 0
        .Name = vbNullString
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RemoveChat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LeaveChat(ByVal index As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' If the player leaves the room and she becomes empty delete it
    If ChatRoom(TempPlayer(index).roomIndex).Members - 1 = 0 Then
        RemoveChat (TempPlayer(index).roomIndex)
    Else
        ChatRoom(TempPlayer(index).roomIndex).Members = ChatRoom(TempPlayer(index).roomIndex).Members - 1
    End If
    
    ' Removes the player's room
    PlayerMsg index, "Left chatroom successfully.", Blue
    ChatRoomMsg index, GetPlayerName(index) & " left the chatroom.", White
    TempPlayer(index).roomIndex = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LeaveChat", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

