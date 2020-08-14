Attribute VB_Name = "modServerLoop"
Option Explicit

' halts thread of execution
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Sub ServerLoop()
    Dim i As Long, x As Long
    Dim Tick As Long, TickCPS As Long, CPS As Long
    Dim tmr500 As Long, tmr1000 As Long
    Dim LastUpdateSavePlayers, LastUpdatePlayers As Long
    Dim LastUpdateMapSpawnItems As Long, LastUpdateMapLogic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ServerOnline = True

    Do While ServerOnline
        Tick = GetTickCount
  
        ' Update players
        If Tick > LastUpdatePlayers Then
            Call UpdatePlayers
            LastUpdatePlayers = GetTickCount + 30
        End If

        If Tick > tmr1000 Then
            ' Check for disconnections every second
            For i = 1 To MAX_PLAYERS
                If frmServer.Socket(i).State > sckConnected Then
                    Call CloseSocket(i)
                End If
            Next
            
            ' Verifies if the server is shutting down
            If isShuttingDown Then
                Call HandleShutdown
            End If
            
            tmr1000 = GetTickCount + 1000
        End If
        
        ' Checks to save players every 5 minutes - Can be tweaked
        If Tick > LastUpdateSavePlayers Then
            UpdateSavePlayers
            LastUpdateSavePlayers = GetTickCount + 300000
        End If
        
        ' Checks to spawn map items every 300 seconds - Can be tweaked
        If Tick > LastUpdateMapSpawnItems Then
            UpdateMapSpawnItems
            LastUpdateMapSpawnItems = GetTickCount + 3000000
        End If
        
        ' update map logic
        If Tick > LastUpdateMapLogic Then
            UpdateMapLogic
            LastUpdateMapLogic = GetTickCount + 500
        End If
    
        ' Make sure we reset the timer for npc hp regeneration
        If Tick > GiveNPCHPTimer + 10000 Then
            GiveNPCHPTimer = GetTickCount
        End If

        ' Lock fps
        If Not CPSUnlock Then Sleep 1
        DoEvents
        
        ' Calculate CPS
        If TickCPS < Tick Then
            GameCPS = CPS
            TickCPS = Tick + 1000
            CPS = 0
        Else
            CPS = CPS + 1
        End If
        
        ' Set server CPS on label
        frmServer.lblCPS.Caption = "CPS: " & Format$(GameCPS, "#,###,###,###")
    Loop

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ServerLoop", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateMapSpawnItems()
    Dim x As Long
    Dim i As Integer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAPS
        ' Clear out unnecessary junk
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, i)
        Next
    
        ' Spawn the items
        Call SpawnMapItems(i)
        Call SendMapItemsToAll(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateMapSpawnItems", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateMapLogic()
    Dim i As Long, x As Long, n As Long
    Dim DistanceX As Long, DistanceY As Long, NpcNum As Long
    Dim target As Long, targetType As Byte, DidWalk As Boolean, Resource_index As Long, Door_index As Long
    Dim TargetX As Long, TargetY As Long, target_verify As Boolean
    Dim MapNum As Integer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For MapNum = 1 To MAX_MAPS
        ' items appearing to everyone
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(MapNum, i).Num > 0 Then
                ' despawn item?
                If MapItem(MapNum, i).canDespawn Then
                    If MapItem(MapNum, i).despawnTimer < GetTickCount Then
                        ' despawn it
                        ClearMapItem i, MapNum
                        ' send updates to everyone
                        SendMapItemsToAll MapNum
                    End If
                End If
            End If
        Next
                        
        For x = 1 To MAX_MAP_NPCS
            NpcNum = MapNpc(MapNum).Npc(x).Num
    
            If Map(MapNum).Npc(x) > 0 Then
                ' This is used for handle dot and hot npc
                If NpcNum > 0 Then
                    For i = 1 To MAX_DOTS
                        HandleDoT_Npc MapNum, x, i
                        HandleHoT_Npc MapNum, x, i
                    Next
                End If
                    
                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(MapNum).Npc(x).Num > 0 Then
                    ' If the npc is a attack on sight, search for a player on the map
                    If Npc(NpcNum).Behaviour = nAttackOnSight Or Npc(NpcNum).Behaviour = nGuard Then
                        ' make sure it' s not stunned
                        If Not MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                            For i = 1 To Player_HighIndex
                                If IsPlaying(i) Then
                                    If GetPlayerMap(i) = MapNum And MapNpc(MapNum).Npc(x).target = 0 And GetPlayerAccess(i) <= ADMIN_MONITOR Then
                                        n = Npc(NpcNum).Range
                                        DistanceX = MapNpc(MapNum).Npc(x).x - GetPlayerX(i)
                                        DistanceY = MapNpc(MapNum).Npc(x).y - GetPlayerY(i)
            
                                        ' Make sure we get a positive value
                                        If DistanceX < 0 Then DistanceX = DistanceX * -1
                                        If DistanceY < 0 Then DistanceY = DistanceY * -1
            
                                        ' Are they in range?  if so GET' M!
                                        If DistanceX <= n And DistanceY <= n Then
                                            If Len(Trim$(Npc(NpcNum).AttackSay)) > 0 Then
                                                Call PlayerMsg(i, Trim$(Npc(NpcNum).Name) & " says: " & Trim$(Npc(NpcNum).AttackSay), SayColor)
                                            End If
                                            MapNpc(MapNum).Npc(x).targetType = 1 ' player
                                            MapNpc(MapNum).Npc(x).target = i
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                End If
                        
                target_verify = False
        
                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(MapNum).Npc(x).Num > 0 Then
                    If MapNpc(MapNum).Npc(x).StunDuration > 0 Then
                        ' check if we can unstun them
                        If GetTickCount > MapNpc(MapNum).Npc(x).StunTimer + (MapNpc(MapNum).Npc(x).StunDuration * 1000) Then
                            MapNpc(MapNum).Npc(x).StunDuration = 0
                            MapNpc(MapNum).Npc(x).StunTimer = 0
                        End If
                    Else
                                  
                        target = MapNpc(MapNum).Npc(x).target
                        targetType = MapNpc(MapNum).Npc(x).targetType
            
                        ' Check to see if its time for the npc to walk
                        If Npc(NpcNum).Behaviour <> nShopKeeper Then
                                
                            If targetType = 1 Then ' player
            
                                ' Check to see if we are following a player or not
                                If target > 0 Then
                
                                    ' Check if the player is even playing, if so follow' m
                                    If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = GetPlayerY(target)
                                        TargetX = GetPlayerX(target)
                                    Else
                                        MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).target = 0
                                    End If
                                End If
                                    
                            ElseIf targetType = 2 Then ' npc
                                        
                                If target > 0 Then
                                            
                                    If MapNpc(MapNum).Npc(target).Num > 0 Then
                                        DidWalk = False
                                        target_verify = True
                                        TargetY = MapNpc(MapNum).Npc(target).y
                                        TargetX = MapNpc(MapNum).Npc(target).x
                                    Else
                                        MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                                        MapNpc(MapNum).Npc(x).target = 0
                                    End If
                                End If
                            End If
                                    
                            If target_verify Then
                                ' Up Left
                                If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                    If MapNpc(MapNum).Npc(x).x > TargetX Then
                                        Call NpcMove(MapNum, x, DIR_UP_LEFT, MOVING_WALKING)
                                        DidWalk = True
                                    End If
                                End If
                                                                                   
                                ' Up right
                                If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                    If MapNpc(MapNum).Npc(x).x < TargetX Then
                                        Call NpcMove(MapNum, x, DIR_UP_RIGHT, MOVING_WALKING)
                                        DidWalk = True
                                    End If
                                End If
                                                                                   
                                ' Down Left
                                If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                    If MapNpc(MapNum).Npc(x).x > TargetX Then
                                        Call NpcMove(MapNum, x, DIR_DOWN_LEFT, MOVING_WALKING)
                                        DidWalk = True
                                    End If
                                End If
                                                                                   
                                ' Down Right
                                If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                    If MapNpc(MapNum).Npc(x).x < TargetX Then
                                        Call NpcMove(MapNum, x, DIR_DOWN_RIGHT, MOVING_WALKING)
                                        DidWalk = True
                                    End If
                                End If
           
                                ' Up
                                If MapNpc(MapNum).Npc(x).y > TargetY And Not DidWalk Then
                                    Call NpcMove(MapNum, x, DIR_UP, MOVING_WALKING)
                                    DidWalk = True
                                End If
           
                                ' Down
                                If MapNpc(MapNum).Npc(x).y < TargetY And Not DidWalk Then
                                    Call NpcMove(MapNum, x, DIR_DOWN, MOVING_WALKING)
                                    DidWalk = True
                                End If
           
                                ' Left
                                If MapNpc(MapNum).Npc(x).x > TargetX And Not DidWalk Then
                                    Call NpcMove(MapNum, x, DIR_LEFT, MOVING_WALKING)
                                    DidWalk = True
                                End If
           
                                ' Right
                                If MapNpc(MapNum).Npc(x).x < TargetX And Not DidWalk Then
                                    Call NpcMove(MapNum, x, DIR_RIGHT, MOVING_WALKING)
                                    DidWalk = True
                                End If
            
                                ' Check if we can' t move and if Target is behind Soundething and if we can just switch dirs
                                If Not DidWalk Then
                                    If MapNpc(MapNum).Npc(x).x - 1 = TargetX And MapNpc(MapNum).Npc(x).y = TargetY Then
                                        If MapNpc(MapNum).Npc(x).Dir <> DIR_LEFT Then
                                            Call NpcDir(MapNum, x, DIR_LEFT)
                                        End If
            
                                        DidWalk = True
                                    End If
            
                                    If MapNpc(MapNum).Npc(x).x + 1 = TargetX And MapNpc(MapNum).Npc(x).y = TargetY Then
                                        If MapNpc(MapNum).Npc(x).Dir <> DIR_RIGHT Then
                                            Call NpcDir(MapNum, x, DIR_RIGHT)
                                        End If
            
                                        DidWalk = True
                                    End If
            
                                    If MapNpc(MapNum).Npc(x).x = TargetX And MapNpc(MapNum).Npc(x).y - 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(x).Dir <> DIR_UP Then
                                            Call NpcDir(MapNum, x, DIR_UP)
                                        End If
            
                                        DidWalk = True
                                    End If
            
                                    If MapNpc(MapNum).Npc(x).x = TargetX And MapNpc(MapNum).Npc(x).y + 1 = TargetY Then
                                        If MapNpc(MapNum).Npc(x).Dir <> DIR_DOWN Then
                                            Call NpcDir(MapNum, x, DIR_DOWN)
                                        End If
            
                                        DidWalk = True
                                    End If
            
                                    ' We could not move so Target must be behind Soundething, walk randomly.
                                    If Not DidWalk Then
                                        i = Int(Rnd * 2)
            
                                        If i = 1 Then
                                            Call NpcMove(MapNum, x, Int(Rnd * 4), MOVING_WALKING)
                                        End If
                                    End If
                                End If
            
                            Else
                                i = Int(Rnd * 4)
                                
                                If i = 1 Then
                                    Call NpcMove(MapNum, x, Int(Rnd * 4), MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If
        
                ' /////////////////////////////////////////////
                ' // This is used for npcs to attack targets //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If MapNpc(MapNum).Npc(x).Num > 0 Then
                    target = MapNpc(MapNum).Npc(x).target
                    targetType = MapNpc(MapNum).Npc(x).targetType
        
                    ' Check if the npc can attack the targeted player player
                    If target > 0 Then
                        If targetType = 1 Then ' player
                            ' Is the target playing and on the same map?
                            If IsPlaying(target) And GetPlayerMap(target) = MapNum Then
                                TryNpcAttackPlayer x, target
                            Else
                                ' Player left map or game, set target to 0
                                MapNpc(MapNum).Npc(x).target = 0
                                MapNpc(MapNum).Npc(x).targetType = 0 ' clear
                            End If
                        End If
                    End If
                End If
        
                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC' s HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen Sounde of the npc' s hp
                If Not MapNpc(MapNum).Npc(x).stopRegen Then
                    If MapNpc(MapNum).Npc(x).Num > 0 And GetTickCount > GiveNPCHPTimer + 10000 Then
                        If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > 0 Then
                            MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = MapNpc(MapNum).Npc(x).Vital(Vitals.HP) + GetNpcVitalRegen(NpcNum, Vitals.HP)
            
                            ' Check if they have more then they should and if so just set it to max
                            If MapNpc(MapNum).Npc(x).Vital(Vitals.HP) > GetNpcMaxVital(NpcNum, Vitals.HP) Then
                                MapNpc(MapNum).Npc(x).Vital(Vitals.HP) = GetNpcMaxVital(NpcNum, Vitals.HP)
                            End If
                        End If
                    End If
                End If
         
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNpc(MapNum).Npc(x).Num = 0 Then
                    If GetTickCount > MapNpc(MapNum).Npc(x).SpawnWait + (Npc(Map(MapNum).Npc(x)).SpawnSecs * 1000) Then
                        Call SpawnNpc(x, MapNum)
                    End If
                End If
            End If
        Next
    
        ' Respawning Resources
        If ResourceCache(MapNum).Resource_Count > 0 Then
            For i = 0 To ResourceCache(MapNum).Resource_Count
                ' Map resource
                Resource_index = Map(MapNum).Tile(ResourceCache(MapNum).ResourceData(i).x, ResourceCache(MapNum).ResourceData(i).y).Data1
        
                ' Prevent subscript out range
                If Resource_index > 0 Then
                    If ResourceCache(MapNum).ResourceData(i).ResourceState = 1 Or ResourceCache(MapNum).ResourceData(i).cur_health < 1 Then  ' dead or fucked up
                        If ResourceCache(MapNum).ResourceData(i).ResourceTimer + (Resource(Resource_index).RespawnTime * 1000) < GetTickCount Then
                            ResourceCache(MapNum).ResourceData(i).ResourceTimer = GetTickCount
                            ResourceCache(MapNum).ResourceData(i).ResourceState = 0 ' normal
                            ' re-set health to resource root
                            ResourceCache(MapNum).ResourceData(i).cur_health = Resource(Resource_index).Health
                            SendResourceCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If
        
        ' Respawning Doors
        If DoorCache(MapNum).Count > 0 Then
            For i = 0 To DoorCache(MapNum).Count
                ' Map door
                Door_index = Map(MapNum).Tile(DoorCache(MapNum).Data(i).x, DoorCache(MapNum).Data(i).y).Data1
        
                ' Prevent subscript out range
                If Door_index > 0 Then
                    If DoorCache(MapNum).Data(i).State = 1 Then ' opening
                        If DoorCache(MapNum).Data(i).RespawnTime + (Door(Door_index).Respawn * 1000) < GetTickCount Then
                            DoorCache(MapNum).Data(i).RespawnTime = GetTickCount
                            DoorCache(MapNum).Data(i).State = 0 ' closed
                            SendDoorCacheToMap MapNum, i
                        End If
                    End If
                End If
            Next
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateMapLogic", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdatePlayers()
    Dim i As Long, x As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            ' check if they' ve completed casting, and if so set the actual Spell going
            If TempPlayer(i).SpellBuffer.Spell > 0 Then
                If GetTickCount > TempPlayer(i).SpellBuffer.Timer + (Spell(Player(i).Char(TempPlayer(i).Char).Spell(TempPlayer(i).SpellBuffer.Spell)).CastTime * 1000) Then
                    CastSpell i, TempPlayer(i).SpellBuffer.Spell, TempPlayer(i).SpellBuffer.target, TempPlayer(i).SpellBuffer.tType
                    TempPlayer(i).SpellBuffer.Spell = 0
                    TempPlayer(i).SpellBuffer.Timer = 0
                    TempPlayer(i).SpellBuffer.target = 0
                    TempPlayer(i).SpellBuffer.tType = 0
                End If
            End If
                    
            ' check if need to turn off stunned
            If TempPlayer(i).StunDuration > 0 Then
                If GetTickCount > TempPlayer(i).StunTimer + (TempPlayer(i).StunDuration * 1000) Then
                    TempPlayer(i).StunDuration = 0
                    TempPlayer(i).StunTimer = 0
                    SendStunned i
                End If
            End If
                    
            ' check regen timer
            If TempPlayer(i).stopRegen Then
                If TempPlayer(i).stopRegenTimer + 5000 < GetTickCount Then
                    TempPlayer(i).stopRegen = False
                    TempPlayer(i).stopRegenTimer = 0
                End If
            End If
                    
            ' HoT and DoT logic
            For x = 1 To MAX_DOTS
                HandleDoT_Player i, x
                HandleHoT_Player i, x
            Next
                    
            ' Checks to update player vitals every 5 seconds - Can be tweaked
            If GetTickCount > LastUpdatePlayerVitals Then
                If Not TempPlayer(i).stopRegen Then
                    If GetPlayerVital(i, Vitals.HP) <> GetPlayerMaxVital(i, Vitals.HP) Then
                        Call SetPlayerVital(i, Vitals.HP, GetPlayerVital(i, Vitals.HP) + GetPlayerVitalRegen(i, Vitals.HP))
                        Call SendVital(i, Vitals.HP)
                        ' send vitals to party if in one
                        If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                    End If
                        
                    If GetPlayerVital(i, Vitals.MP) <> GetPlayerMaxVital(i, Vitals.MP) Then
                        Call SetPlayerVital(i, Vitals.MP, GetPlayerVital(i, Vitals.MP) + GetPlayerVitalRegen(i, Vitals.MP))
                        Call SendVital(i, Vitals.MP)
                        ' send vitals to party if in one
                        If TempPlayer(i).inParty > 0 Then SendPartyVitals TempPlayer(i).inParty, i
                    End If
                End If
            
                LastUpdatePlayerVitals = GetTickCount + 5000
            End If
        End If
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdatePlayers", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub UpdateSavePlayers()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Do not run if it has no players online
    If TotalOnlinePlayers <= 0 Then Exit Sub
    
    ' Add log
    Call TextAdd("Saving all online players...")

    ' Save all players
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateSavePlayers", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleShutdown()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Resets the time
    If Secs <= 0 Then Secs = 30
    
    ' Send a global message if the time to shut down the server is equal or less than five
    If Secs <= 5 Then
        Call GlobalMsg("Server Shutdown in " & Secs & " seconds.", BrightBlue)
        Call TextAdd("Automated Server Shutdown in " & Secs & " seconds.")
    End If
    Secs = Secs - 1

    ' Destroy server
    If Secs <= 0 Then
        Call GlobalMsg("Server Shutdown.", BrightRed)
        Call DestroyServer
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleShutdown", "modServerLoop", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
