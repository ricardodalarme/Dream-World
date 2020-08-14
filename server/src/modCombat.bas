Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################

Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim x As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    ' Original player vital
    Select Case Vital
        Case Vitals.HP
            x = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Stats.Endurance) / 2)) * 15 + 100
        Case Vitals.MP
            x = ((GetPlayerLevel(index) / 2) + (GetPlayerStat(index, Stats.Intelligence) / 2)) * 5 + 25
    End Select
    
    ' Vital of titles
    x = GetTitleVital(index, Vital, x)

    ' Return function value
    GetPlayerMaxVital = x

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerMaxVital", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out of range
    If index > MAX_PLAYERS Or TempPlayer(index).Char = 0 Then Exit Function

    ' Return function value
    Select Case Vital
        Case Vitals.HP
            GetPlayerVitalRegen = (GetPlayerStat(index, Stats.Willpower) * 0.8) + 7
        Case Vitals.MP
            GetPlayerVitalRegen = (GetPlayerStat(index, Stats.Willpower) / 4) + 11.5
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerVitalRegen", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function

    ' Return function value
    Select Case Vital
        Case Vitals.HP
            GetNpcMaxVital = Npc(NpcNum).HP
        Case Vitals.MP
            GetNpcMaxVital = 30 + (Npc(NpcNum).Stat(Stats.Intelligence) * 10) + 2
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetNpcMaxVital", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function

    ' Return function value
    Select Case Vital
        Case Vitals.HP
            GetNpcVitalRegen = (Npc(NpcNum).Stat(Stats.Willpower) * 0.8) + 6
        Case Vitals.MP
            GetNpcVitalRegen = (Npc(NpcNum).Stat(Stats.Willpower) / 4) + 12.5
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetNpcVitalRegen", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Function GetNpcDamage(ByVal NpcNum As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function
    
    ' Return function value
    GetNpcDamage = 0.085 * 5 * Npc(NpcNum).Stat(Stats.Strength) * Npc(NpcNum).Damage + (Npc(NpcNum).Level / 5)

    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetNpcDamage", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ###############################
' ##      Luck-based raes     ##
' ###############################
Public Function CanPlayerBlock(ByVal index As Long) As Boolean
    Dim Rate As Long, RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Then Exit Function
    
    ' Random
    Rate = GetPlayerStat(index, Stats.Endurance) / 13.82
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanPlayerBlock = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanPlayerBlock", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanPlayerCrit(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Then Exit Function
    
    ' Random
    Rate = (GetPlayerStat(index, Stats.Strength) \ 3.3) + (GetPlayerLevel(index) \ 2.8)
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanPlayerCrit = True
    End If
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanPlayerCrit", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Then Exit Function
    
    ' Random
    Rate = GetPlayerStat(index, Stats.Agility) / 33.3
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanPlayerDodge = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanPlayerDodge", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanPlayerParry(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(index) Then Exit Function
    
    ' Random
    Rate = GetPlayerStat(index, Stats.Strength) * 0.25
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanPlayerParry = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanPlayerParry", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanNpcBlock(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function
    
    ' Random
    Rate = Npc(NpcNum).Stat(Endurance) / 13.82
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanNpcBlock = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanNpcBlock", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanNpcCrit(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function
    
    ' Random
    Rate = (Npc(NpcNum).Stat(Stats.Strength) \ 3.3) + (Npc(NpcNum).Level \ 2.8)
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanNpcCrit = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanNpcCrit", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanNpcDodge(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function
    
    ' Random
    Rate = Npc(NpcNum).Stat(Stats.Agility) / 33.3
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanNpcDodge = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanNpcDodge", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function CanNpcParry(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RndNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function
    
    ' Random
    Rate = Npc(NpcNum).Stat(Stats.Strength) * 0.25
    RndNum = RAND(1, 100)
    
    ' Return function value
    If RndNum <= Rate Then
        CanNpcParry = True
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanNpcParry", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function GetPlayerDamageBase(ByVal index As Long, ByVal Base As Stats, ByVal Value As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    GetPlayerDamageBase = GetPlayerStat(index, Base) * 1.5 + GetPlayerLevel(index) / 6.2 + Value
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "GetPlayerDamageBase", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################

Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNpcNum As Long)
    Dim Damage As Long, BlockAmount As Long
    Dim NpcNum As Long
    Dim MapNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Can we attack the npc?
    If Not CanPlayerAttackNpc(index, MapNpcNum) Then Exit Sub
    
    ' Get npc data
    MapNum = GetPlayerMap(index)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Check if NPC can avoid the attack
    If CanNpcDodge(NpcNum) Then
        SendActionMsg MapNum, "Dodge!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        Exit Sub
    ElseIf CanNpcParry(NpcNum) Then
        SendActionMsg MapNum, "Parry!", Pink, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        Exit Sub
    End If

    ' Get the damage we can do
    Damage = GetPlayerDamage(index)
        
    ' if the npc blocks, take away the block amount
    BlockAmount = CanNpcBlock(MapNpcNum)
    Damage = Damage - BlockAmount

    ' randomise from 1 to max hit
    Damage = RAND(1, Damage)
        
    ' * 1.5 if it' s a crit!
    If CanPlayerCrit(index) Then
        Damage = Damage * 1.5
        SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    End If
            
    If Damage > 0 Then
        Call PlayerAttackNpc(index, MapNpcNum, Damage)
    Else
        Call PlayerMsg(index, "Your attack does nothing.", BrightRed)
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TryPlayerAttackNpc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function CanPlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim NpcX As Long, NpcY As Long
    Dim AttackSpeed As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(attacker) Then Exit Function
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then Exit Function
    If MapNpc(GetPlayerMap(attacker)).Npc(MapNpcNum).Num <= 0 Then Exit Function

    ' Get npc data
    MapNum = GetPlayerMap(attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Make sure the npc isn' t already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then Exit Function

    ' exit out early
    If IsSpell Then
        If Npc(NpcNum).Behaviour <> nFriendly And Npc(NpcNum).Behaviour <> nShopKeeper And Npc(NpcNum).Behaviour <> nQuest Then
            CanPlayerAttackNpc = True
            Exit Function
        End If
    End If

    ' attack speed from weapon
    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        AttackSpeed = Item(GetPlayerEquipment(attacker, Weapon)).Speed
    Else
        AttackSpeed = 1000
    End If

    ' Checks if the attack is valid
    If GetTickCount <= TempPlayer(attacker).AttackTimer + AttackSpeed Then Exit Function
    
    ' Check if at same coordinates
    Call GetNpcNextTile(MapNum, MapNpcNum, GetPlayerDir(attacker), NpcX, NpcY)

    ' Checks if the npc is in front of the attacker
    If NpcX <> GetPlayerX(attacker) And NpcY <> GetPlayerY(attacker) Then Exit Function
    
    If Npc(NpcNum).Behaviour <> nFriendly And Npc(NpcNum).Behaviour <> nShopKeeper And Npc(NpcNum).Behaviour <> nQuest Then
        CanPlayerAttackNpc = True
    Else
        ' Speaks of the npc
        If Len(Trim$(Npc(NpcNum).AttackSay)) > 0 Then PlayerMsg attacker, Trim$(Npc(NpcNum).Name) & ": " & Trim$(Npc(NpcNum).AttackSay), White

        ' Check if the player completed a quest
        CheckCompleteQuest attacker, NpcNum
                        
        ' Open the selector of quest
        Select Case Npc(NpcNum).Behaviour
            Case NpcBehaviour.nQuest
                SendQuestCommand attacker, 1, NpcNum
                TempPlayer(attacker).QuestSelect = NpcNum
            Case NpcBehaviour.nShopKeeper
                SendOpenShop attacker, Npc(NpcNum).ShopNum
                TempPlayer(attacker).InShop = Npc(NpcNum).ShopNum ' stops movement and the like
        End Select
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanPlayerAttackNpc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub PlayerAttackNpc(ByVal attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0, Optional ByVal overTime As Boolean = False)
    Dim EXP As Long
    Dim n As Long
    Dim i As Long
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(attacker) Then Exit Sub
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then Exit Sub
        
    ' Declarations
    MapNum = GetPlayerMap(attacker)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num

    ' Check for weapon
    If GetPlayerEquipment(attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(attacker, Weapon)
    Else
        n = 0
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount
        
    If Damage >= MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) Then
        SendActionMsg GetPlayerMap(attacker), "-" & MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP), BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y)
            End If
        End If

        ' Calculate exp to give attacker
        EXP = Npc(NpcNum).EXP

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If

        ' in party?
        If TempPlayer(attacker).inParty > 0 Then
            ' pass through party sharing function
            Party_ShareExp TempPlayer(attacker).inParty, EXP, attacker
        Else
            ' no party - keep exp for self
            GivePlayerEXP attacker, EXP
        End If
        
        ' Drop the goods if they get it
        For n = 1 To MAX_NPC_DROPS
            If Rnd <= Npc(NpcNum).DropChance(n) Then
                Call SpawnItem(Npc(NpcNum).DropItem(n), Npc(NpcNum).DropItemValue(n), MapNum, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y)
            End If
        Next

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).Npc(MapNpcNum).Num = 0
        MapNpc(MapNum).Npc(MapNpcNum).SpawnWait = GetTickCount
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = 0
        
        ' Kill npcs
        SetPlayerKillNpcs attacker, NpcNum, GetPlayerKillNpcs(attacker, NpcNum) + 1
        
        ' Check if the player completed a quest
        CheckCompleteQuest attacker, , True

        ' clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).Npc(MapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).Npc(MapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' send death to the map
        Set Buffer = New clsBuffer
        Buffer.WriteLong ServerPackets.SNpcDead
        Buffer.WriteLong MapNpcNum
        SendDataToMap MapNum, Buffer.ToArray()
        Set Buffer = Nothing
        
        ' Loop through entire map and purge NPC from targets
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = MapNum Then
                    If TempPlayer(i).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).target = MapNpcNum Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' NPC not dead, just do the damage
        MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) - Damage

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        SendBlood GetPlayerMap(attacker), MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound attacker, MapNpc(MapNum).Npc(MapNpcNum).x, MapNpc(MapNum).Npc(MapNpcNum).y, SoundEntity.seSpell, SpellNum
        
        ' send animation
        If n > 0 Then
            If Not overTime Then
                If SpellNum = 0 Then Call SendAnimation(MapNum, Item(GetPlayerEquipment(attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
            End If
        End If

        ' Set the NPC target to the player
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 1 ' player
        MapNpc(MapNum).Npc(MapNpcNum).target = attacker

        ' Now check for guard ai and if so have all onmap guards come after' m
        If Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Behaviour = nGuard Then
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(MapNum).Npc(i).Num = MapNpc(MapNum).Npc(MapNpcNum).Num Then
                    MapNpc(MapNum).Npc(i).target = attacker
                    MapNpc(MapNum).Npc(i).targetType = 1 ' player
                End If
            Next
        End If
        
        ' set the regen timer
        MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
        MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = GetTickCount
        
        ' if stunning Spell, stun the npc
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNpcNum, MapNum, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc MapNum, MapNpcNum, SpellNum, attacker
            End If
        End If
        
        SendMapNpcVitals MapNum, MapNpcNum
    End If

    ' Reset attack timer
    If SpellNum = 0 Then TempPlayer(attacker).AttackTimer = GetTickCount

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerAttackNpc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################

Public Sub TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
    Dim MapNum As Long, NpcNum As Long, BlockAmount As Long, Damage As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Can the npc attack the player?
    If CanNpcAttackPlayer(MapNpcNum, index) Then
        MapNum = GetPlayerMap(index)
        NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
        ' check if PLAYER can avoid the attack
        If CanPlayerDodge(index) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (Player(index).Char(TempPlayer(index).Char).x * 32), (Player(index).Char(TempPlayer(index).Char).y * 32)
            Exit Sub
        ElseIf CanPlayerParry(index) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (Player(index).Char(TempPlayer(index).Char).x * 32), (Player(index).Char(TempPlayer(index).Char).y * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetNpcDamage(NpcNum)
        
        ' if the player blocks, take away the block amount
        BlockAmount = CanPlayerBlock(index)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(index, Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' Checks if the attack was critical
        If CanNpcCrit(NpcNum) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (MapNpc(MapNum).Npc(MapNpcNum).x * 32), (MapNpc(MapNum).Npc(MapNpcNum).y * 32)
        End If

        If Damage > 0 Then
            Call NpcAttackPlayer(MapNpcNum, index, Damage)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TryNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
    Dim MapNum As Long
    Dim NpcNum As Long
    Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then Exit Function
    If MapNpc(GetPlayerMap(index)).Npc(MapNpcNum).Num <= 0 Then Exit Function

    ' Npc data
    MapNum = GetPlayerMap(index)
    NpcNum = MapNpc(MapNum).Npc(MapNpcNum).Num
    
    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function

    ' Make sure the npc isn' t already dead
    If MapNpc(MapNum).Npc(MapNpcNum).Vital(Vitals.HP) <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNpc(MapNum).Npc(MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then
        Exit Function
    End If

    ' Send this packet so they can see the npc attacking
    Set Buffer = New clsBuffer
    Buffer.WriteLong ServerPackets.SNpcAttack
    Buffer.WriteLong MapNpcNum
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
    
    ' Reset npc attack
    MapNpc(MapNum).Npc(MapNpcNum).AttackTimer = GetTickCount

    ' Check if at same coordinates
    If (GetPlayerY(index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
        CanNpcAttackPlayer = True
    Else
        If (GetPlayerY(index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) = MapNpc(MapNum).Npc(MapNpcNum).x) Then
            CanNpcAttackPlayer = True
        Else
            If (GetPlayerY(index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) = MapNpc(MapNum).Npc(MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNpc(MapNum).Npc(MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanNpcAttackPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim MapNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(victim) Then Exit Sub
    If MapNpc(GetPlayerMap(victim)).Npc(MapNpcNum).Num <= 0 Then Exit Sub
    
    MapNum = GetPlayerMap(victim)
    Name = Trim$(Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Name)
    
    If Damage <= 0 Then Exit Sub

    ' set the regen timer
    MapNpc(MapNum).Npc(MapNpcNum).stopRegen = True
    MapNpc(MapNum).Npc(MapNpcNum).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
        
        ' kill player
        KillPlayer victim
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & Name, BrightRed)

        ' Set NPC target to 0
        MapNpc(MapNum).Npc(MapNpcNum).target = 0
        MapNpc(MapNum).Npc(MapNpcNum).targetType = 0
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        Call SendAnimation(MapNum, Npc(MapNpc(GetPlayerMap(victim)).Npc(MapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, victim)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seNpc, MapNpc(MapNum).Npc(MapNpcNum).Num
        
        ' Say damage
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NpcAttackPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################

Public Sub TryPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long)
    Dim BlockAmount As Long
    Dim MapNum As Long
    Dim Damage As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Can we attack the npc?
    If CanPlayerAttackPlayer(attacker, victim) Then
        ' Map on which the players are
        MapNum = GetPlayerMap(attacker)
    
        ' check if NPC can avoid the attack
        If CanPlayerDodge(victim) Then
            SendActionMsg MapNum, "Dodge!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        ElseIf CanPlayerParry(victim) Then
            SendActionMsg MapNum, "Parry!", Pink, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
            Exit Sub
        End If

        ' Get the damage we can do
        Damage = GetPlayerDamage(attacker)
        
        ' if the npc blocks, take away the block amount
        BlockAmount = CanPlayerBlock(victim)
        Damage = Damage - BlockAmount
        
        ' take away armour
        Damage = Damage - RAND(1, (GetPlayerStat(victim, Stats.Agility) * 2))
        
        ' randomise for up to 10% lower than max hit
        Damage = RAND(1, Damage)
        
        ' Checks if the attack was critical
        If CanPlayerCrit(attacker) Then
            Damage = Damage * 1.5
            SendActionMsg MapNum, "Critical!", BrightCyan, 1, (GetPlayerX(attacker) * 32), (GetPlayerY(attacker) * 32)
        End If

        If Damage > 0 Then
            Call PlayerAttackPlayer(attacker, victim, Damage)
        Else
            Call PlayerMsg(attacker, "Your attack does nothing.", BrightRed)
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TryPlayerAttackPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CanPlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsSpell Then
        ' Check attack timer
        If GetPlayerEquipment(attacker, Weapon) > 0 Then
            If GetTickCount < TempPlayer(attacker).AttackTimer + Item(GetPlayerEquipment(attacker, Weapon)).Speed Then Exit Function
        Else
            If GetTickCount < TempPlayer(attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If

    ' Prevent subscript out of range
    If Not IsPlaying(victim) Then Exit Function

    ' Make sure they are on the same map
    If Not GetPlayerMap(attacker) = GetPlayerMap(victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(victim).GettingMap = YES Then Exit Function

    If Not IsSpell Then
        ' Check if at same coordinates
        Select Case GetPlayerDir(attacker)
            Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
   
                If Not ((GetPlayerY(victim) + 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
   
                If Not ((GetPlayerY(victim) - 1 = GetPlayerY(attacker)) And (GetPlayerX(victim) = GetPlayerX(attacker))) Then Exit Function
            Case DIR_LEFT
   
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) + 1 = GetPlayerX(attacker))) Then Exit Function
            Case DIR_RIGHT
   
                If Not ((GetPlayerY(victim) = GetPlayerY(attacker)) And (GetPlayerX(victim) - 1 = GetPlayerX(attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If

    ' Check if map is attackable
    If Not Map(GetPlayerMap(attacker)).Moral = MAP_MORAL_NONE Then
        If GetPlayerPK(victim) = NO Then
            Call PlayerMsg(attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerVital(victim, Vitals.HP) <= 0 Then Exit Function

    ' Check to make sure that they dont have access
    If GetPlayerAccess(attacker) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "Admins cannot attack other players.", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn' t an admin
    If GetPlayerAccess(victim) > ADMIN_MONITOR Then
        Call PlayerMsg(attacker, "You cannot attack " & GetPlayerName(victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(attacker) < 10 Then
        Call PlayerMsg(attacker, "You are below level 10, you cannot attack another player yet!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(victim) < 10 Then
        Call PlayerMsg(attacker, GetPlayerName(victim) & " is below level 10, you cannot attack this player yet!", BrightRed)
        Exit Function
    End If

    CanPlayerAttackPlayer = True

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CanPlayerAttackPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub PlayerAttackPlayer(ByVal attacker As Long, ByVal victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long = 0)
    Dim EXP As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Not IsPlaying(attacker) Or IsPlaying(victim) = False Or Damage < 0 Then
        Exit Sub
    End If
    
    ' set the regen timer
    TempPlayer(attacker).stopRegen = True
    TempPlayer(attacker).stopRegenTimer = GetTickCount

    If Damage >= GetPlayerVital(victim, Vitals.HP) Then
        SendActionMsg GetPlayerMap(victim), "-" & GetPlayerVital(victim, Vitals.HP), BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(victim) & " has been killed by " & GetPlayerName(attacker), BrightRed)
        ' Calculate exp to give attacker
        EXP = (GetPlayerExp(victim) \ 10)

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 0
        End If

        If EXP = 0 Then
            Call PlayerMsg(victim, "You lost no exp.", BrightRed)
            Call PlayerMsg(attacker, "You received no exp.", BrightBlue)
        Else
            Call SetPlayerExp(victim, GetPlayerExp(victim) - EXP)
            SendEXP victim
            Call PlayerMsg(victim, "You lost " & EXP & " exp.", BrightRed)
            
            ' check if we' re in a party
            If TempPlayer(attacker).inParty > 0 Then
                ' pass through party exp share function
                Party_ShareExp TempPlayer(attacker).inParty, EXP, attacker
            Else
                ' not in party, get exp for self
                GivePlayerEXP attacker, EXP
            End If
        End If
        
        ' purge target info of anyone who targetted dead guy
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And IsConnected(i) Then
                If GetPlayerMap(i) = GetPlayerMap(attacker) Then
                    If TempPlayer(i).target = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).target = victim Then
                            TempPlayer(i).target = 0
                            TempPlayer(i).targetType = TARGET_TYPE_NONE
                            SendTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(victim) = NO Then
            If GetPlayerPK(attacker) = NO Then
                Call SetPlayerPK(attacker, YES)
                Call SendPlayerData(attacker)
                Call GlobalMsg(GetPlayerName(attacker) & " has been deemed a Player Killer!!!", BrightRed)
            End If

        Else
            Call GlobalMsg(GetPlayerName(victim) & " has paid the price for being a Player Killer!!!", BrightRed)
        End If

        Call OnDeath(victim)
        
        ' Kill players
        Call SetPlayerKillPlayers(attacker, GetPlayerKillPlayers(attacker) + 1)
        
        ' Check if the player completed a quest
        Call CheckCompleteQuest(attacker, , True)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(victim, Vitals.HP, GetPlayerVital(victim, Vitals.HP) - Damage)
        Call SendVital(victim, Vitals.HP)
        
        ' send vitals to party if in one
        If TempPlayer(victim).inParty > 0 Then SendPartyVitals TempPlayer(victim).inParty, victim
        
        ' send the sound
        If SpellNum > 0 Then SendMapSound victim, GetPlayerX(victim), GetPlayerY(victim), SoundEntity.seSpell, SpellNum
        
        SendActionMsg GetPlayerMap(victim), "-" & Damage, BrightRed, 1, (GetPlayerX(victim) * 32), (GetPlayerY(victim) * 32)
        SendBlood GetPlayerMap(victim), GetPlayerX(victim), GetPlayerY(victim)
        
        ' set the regen timer
        TempPlayer(victim).stopRegen = True
        TempPlayer(victim).stopRegenTimer = GetTickCount
        
        ' if a stunning Spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player victim, SpellNum, attacker
            End If
        End If
    End If

    ' Reset attack timer
    If SpellNum = 0 Then TempPlayer(attacker).AttackTimer = GetTickCount

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PlayerAttackPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ############
' ## Spells ##
' ############

Public Sub BufferSpell(ByVal index As Long, ByVal Spellslot As Long)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Long
    Dim SpellCastType As Long
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    Dim targetType As Byte
    Dim target As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Spellslot <= 0 Or Spellslot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(index, Spellslot)
    MapNum = GetPlayerMap(index)
    
    If SpellNum <= 0 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    ' Make sure player has the Spell
    If Not HasSpell(index, SpellNum) Then Exit Sub
    
    ' see if cooldown has finished
    If TempPlayer(index).SpellCD(Spellslot) > GetTickCount Then
        PlayerMsg index, "Spell hasn' t cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this Spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this Spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this Spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of Spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    targetType = TempPlayer(index).targetType
    target = TempPlayer(index).target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' check if have target
            If Not target > 0 Then
                PlayerMsg index, "You do not have a target.", BrightRed
            End If
            If targetType = TARGET_TYPE_PLAYER Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(target), GetPlayerY(target)) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                Else
                    ' go through Spell types
                    If Spell(SpellNum).Type <> sDamageHP And Spell(SpellNum).Type <> sDamageMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' if have target, check in range
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(MapNum).Npc(target).x, MapNpc(MapNum).Npc(target).y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    HasBuffered = False
                Else
                    ' go through Spell types
                    If Spell(SpellNum).Type <> sDamageHP And Spell(SpellNum).Type <> sDamageMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        TempPlayer(index).SpellBuffer.Spell = Spellslot
        TempPlayer(index).SpellBuffer.Timer = GetTickCount
        TempPlayer(index).SpellBuffer.target = TempPlayer(index).target
        TempPlayer(index).SpellBuffer.tType = TempPlayer(index).targetType
        Exit Sub
    Else
        SendClearSpellBuffer index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BufferSpell", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal Spellslot As Long, ByVal target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim MapNum As Long
    Dim Vital As Long, VitalType As Byte, increment As Boolean
    Dim i As Long, n As Long
    Dim AoE As Long, Range As Long
    Dim x As Long, y As Long
    Dim SpellCastType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If Spellslot <= 0 Or Spellslot > MAX_PLAYER_SPELLS Then Exit Sub

    ' Data
    SpellNum = GetPlayerSpell(index, Spellslot)
    MapNum = GetPlayerMap(index)

    ' Make sure player has the Spell
    If Not HasSpell(index, SpellNum) Then Exit Sub

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    n = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If n > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & n & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    n = Spell(SpellNum).AccessReq
    
    ' make sure they have the right access
    If n > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be an administrator to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    n = Spell(SpellNum).ClassReq
    
    ' make sure the classreq > 0
    If n > 0 Then ' 0 = no req
        If n <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(n).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' find out what kind of Spell it is! self cast, target or AOE
    If Spell(SpellNum).Range > 0 Then
        ' ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoE Then
            SpellCastType = 0 ' self-cast
        Else
            SpellCastType = 1 ' self-cast AoE
        End If
    End If
    
    ' set the vital
    Vital = GetPlayerDamageBase(index, Spell(SpellNum).BaseStat, Spell(SpellNum).Vital)
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
    
    Select Case SpellCastType
        Case 0 ' self-cast target
            Select Case Spell(SpellNum).Type
                Case SpellType.sHealHP
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, SpellNum
                Case SpellType.sHealMP
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, SpellNum
                Case SpellType.sWarp
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SetPlayerDir index, Spell(SpellNum).Dir
                    PlayerWarp index, Spell(SpellNum).Map, Spell(SpellNum).x, Spell(SpellNum).y
                    SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
            End Select
        Case 1, 3 ' self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                x = GetPlayerX(index)
                y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    x = GetPlayerX(target)
                    y = GetPlayerY(target)
                Else
                    x = MapNpc(MapNum).Npc(target).x
                    y = MapNpc(MapNum).Npc(target).y
                End If
                
                If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                    PlayerMsg index, "Target not in range.", BrightRed
                    SendClearSpellBuffer index
                End If
            End If
            Select Case Spell(SpellNum).Type
                Case SpellType.sDamageHP
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, True) Then
                                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, Vital, SpellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(Vitals.HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    If CanPlayerAttackNpc(index, i, True) Then
                                        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                        PlayerAttackNpc index, i, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                Case SpellType.sHealHP, SpellType.sHealMP, SpellType.sDamageMP
                    If Spell(SpellNum).Type = SpellType.sHealHP Then
                        VitalType = Vitals.HP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SpellType.sHealMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SpellType.sDamageMP Then
                        VitalType = Vitals.MP
                        increment = False
                    End If
                                        
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If i <> index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If isInRange(AoE, x, y, GetPlayerX(i), GetPlayerY(i)) Then
                                        SpellPlayer_Effect VitalType, increment, i, Vital, SpellNum
                                    End If
                                End If
                            End If
                        End If
                    Next
                    For i = 1 To MAX_MAP_NPCS
                        If MapNpc(MapNum).Npc(i).Num > 0 Then
                            If MapNpc(MapNum).Npc(i).Vital(Vitals.HP) > 0 Then
                                If isInRange(AoE, x, y, MapNpc(MapNum).Npc(i).x, MapNpc(MapNum).Npc(i).y) Then
                                    SpellNpc_Effect VitalType, increment, i, Vital, SpellNum, MapNum
                                End If
                            End If
                        End If
                    Next
            End Select
        Case 2 ' targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                x = GetPlayerX(target)
                y = GetPlayerY(target)
            Else
                x = MapNpc(MapNum).Npc(target).x
                y = MapNpc(MapNum).Npc(target).y
            End If
                
            If Not isInRange(Range, GetPlayerX(index), GetPlayerY(index), x, y) Then
                PlayerMsg index, "Target not in range.", BrightRed
                SendClearSpellBuffer index
                Exit Sub
            End If
            
            Select Case Spell(SpellNum).Type
                Case SpellType.sDamageHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                PlayerAttackPlayer index, target, Vital, SpellNum
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, target, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                                PlayerAttackNpc index, target, Vital, SpellNum
                            End If
                        End If
                    End If
                    
                Case SpellType.sDamageMP, SpellType.sHealMP, SpellType.sHealHP
                    If Spell(SpellNum).Type = SpellType.sDamageMP Then
                        VitalType = Vitals.MP
                        increment = False
                    ElseIf Spell(SpellNum).Type = SpellType.sHealMP Then
                        VitalType = Vitals.MP
                        increment = True
                    ElseIf Spell(SpellNum).Type = SpellType.sHealHP Then
                        VitalType = Vitals.HP
                        increment = True
                    End If
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        SpellPlayer_Effect VitalType, increment, target, Vital, SpellNum
                    Else
                        SpellNpc_Effect VitalType, increment, target, Vital, SpellNum, MapNum
                    End If
            End Select
    End Select
    
    ' Update player vitals
    Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
    Call SendVital(index, Vitals.MP)
    
    ' send vitals to party if in one
    If TempPlayer(index).inParty > 0 Then SendPartyVitals TempPlayer(index).inParty, index
        
    ' Spell cooldown
    TempPlayer(index).SpellCD(Spellslot) = GetTickCount + (Spell(SpellNum).CDTime * 1000)
    Call SendCooldown(index, Spellslot)
    
    ' Say msg
    SendActionMsg MapNum, Trim$(Spell(SpellNum).Name) & "!", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CastSpell", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
    Dim sSymbol As String * 1
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' No need to make the rest
    If Damage <= 0 Then Exit Sub
    
    ' Message data
    If Vital = Vitals.HP Then Colour = BrightGreen
    If Vital = Vitals.MP Then Colour = BrightBlue

    If increment Then
        sSymbol = "+"
    Else
        sSymbol = "-"
    End If
    
    ' Say msg
    SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
    SendActionMsg GetPlayerMap(index), sSymbol & Damage, Colour, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
    ' send the sound
    SendMapSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
        
    ' Add effect
    If increment Then
        SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
        If Spell(SpellNum).Duration > 0 Then
            AddHoT_Player index, SpellNum
        End If
    ElseIf Not increment Then
        SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellPlayer_Effect", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Long)
    Dim sSymbol As String * 1
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' No need to make the rest
    If Damage <= 0 Then Exit Sub
    
    ' Message data
    If Vital = Vitals.HP Then Colour = BrightGreen
    If Vital = Vitals.MP Then Colour = BrightBlue

    If increment Then
        sSymbol = "+"
    Else
        sSymbol = "-"
    End If
    
    ' Say msg
    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
    SendActionMsg MapNum, sSymbol & Damage, Colour, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
        
    ' send the sound
    SendMapSound index, MapNpc(MapNum).Npc(index).x, MapNpc(MapNum).Npc(index).y, SoundEntity.seSpell, SpellNum
        
    ' Add effect
    If increment Then
        MapNpc(MapNum).Npc(index).Vital(Vital) = MapNpc(MapNum).Npc(index).Vital(Vital) + Damage
        If Spell(SpellNum).Duration > 0 Then
            AddHoT_Npc MapNum, index, SpellNum
        End If
    ElseIf Not increment Then
        MapNpc(MapNum).Npc(index).Vital(Vital) = MapNpc(MapNum).Npc(index).Vital(Vital) - Damage
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SpellNpc_Effect", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            ' ' Renew time
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
            End If
            
            ' Add dot
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
            End If
        End With
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddDoT_Player", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal SpellNum As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            ' Renew time
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
            End If
            
            ' Add hot
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
            End If
        End With
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddHoT_Player", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(index).DoT(i)
            ' Renew time
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .Caster = Caster
                .StartTime = GetTickCount
            End If
            
            ' Add dot
            If .Used = False Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Caster = Caster
                .Used = True
                .StartTime = GetTickCount
            End If
        End With
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddDoT_Npc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal SpellNum As Long)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).Npc(index).HoT(i)
            ' Renew time
            If .Spell = SpellNum Then
                .Timer = GetTickCount
                .StartTime = GetTickCount
            End If
            
            ' Add hot
            If Not .Used Then
                .Spell = SpellNum
                .Timer = GetTickCount
                .Used = True
                .StartTime = GetTickCount
            End If
        End With
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddHoT_Npc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TempPlayer(index).DoT(dotNum)
        ' Prevent subscript out range
        If Not .Used Or .Spell <= 0 Or .Spell > MAX_SPELLS Then
            Exit Sub
        End If
        
        ' time to tick?
        If GetTickCount <= .Timer + (Spell(.Spell).Interval * 1000) Then
            Exit Sub
        End If

        ' Add effect
        If CanPlayerAttackPlayer(.Caster, index, True) Then
            PlayerAttackPlayer .Caster, index, GetPlayerDamageBase(.Spell, Spell(.Spell).BaseStat, Spell(.Spell).Vital)
        End If
        .Timer = GetTickCount

        ' destroy DoT if finished
        If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End If
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoT_Player", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With TempPlayer(index).HoT(hotNum)
        ' Prevent subscript out range
        If Not .Used Or .Spell <= 0 Or .Spell > MAX_SPELLS Then
            Exit Sub
        End If
        
        ' time to tick?
        If GetTickCount <= .Timer + (Spell(.Spell).Interval * 1000) Then
            Exit Sub
        End If

        ' Say msg
        SendActionMsg GetPlayerMap(index), "+" & GetPlayerDamageBase(.Spell, Spell(.Spell).BaseStat, Spell(.Spell).Vital), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' Add effect
        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + GetPlayerDamageBase(.Spell, Spell(.Spell).BaseStat, Spell(.Spell).Vital))
        .Timer = GetTickCount

        ' destroy hoT if finished
        If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End If
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHoT_Player", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal dotNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With MapNpc(MapNum).Npc(index).DoT(dotNum)
        ' Prevent subscript out range
        If Not .Used Or .Spell <= 0 Or .Spell > MAX_SPELLS Then
            Exit Sub
        End If
        
        ' time to tick?
        If GetTickCount <= .Timer + (Spell(.Spell).Interval * 1000) Then
            Exit Sub
        End If

        ' Add effect
        If CanPlayerAttackNpc(.Caster, index, True) Then
            PlayerAttackNpc .Caster, index, GetPlayerDamageBase(.Spell, Spell(.Spell).BaseStat, Spell(.Spell).Vital), , True
        End If
        .Timer = GetTickCount

        ' destroy DoT if finished
        If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End If
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleDoT_Npc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Long, ByVal index As Long, ByVal hotNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With MapNpc(MapNum).Npc(index).HoT(hotNum)
        ' Prevent subscript out range
        If Not .Used Or .Spell <= 0 Or .Spell > MAX_SPELLS Then
            Exit Sub
        End If
        
        ' time to tick?
        If GetTickCount <= .Timer + (Spell(.Spell).Interval * 1000) Then
            Exit Sub
        End If
        
        ' Say msg
        SendActionMsg MapNum, "+" & GetPlayerDamageBase(.Spell, Spell(.Spell).BaseStat, Spell(.Spell).Vital), BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).Npc(index).x * 32, MapNpc(MapNum).Npc(index).y * 32
        
        ' Add the effect
        MapNpc(MapNum).Npc(index).Vital(Vitals.HP) = MapNpc(MapNum).Npc(index).Vital(Vitals.HP) + Spell(.Spell).Vital
        .Timer = GetTickCount

        ' destroy hoT if finished
        If GetTickCount - .StartTime >= (Spell(.Spell).Duration * 1000) Then
            .Used = False
            .Spell = 0
            .Timer = GetTickCount
            .Caster = 0
            .StartTime = 0
        End If
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleHoT_Npc", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal SpellNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check if it' s a stunning Spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        TempPlayer(index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(index).StunTimer = GetTickCount
        ' send it to the index
        SendStunned index
        ' tell him he' s stunned
        PlayerMsg index, "You have been stunned.", BrightRed
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "StunPlayer", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal MapNum As Long, ByVal SpellNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' check if it' s a stunning Spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' set the values on index
        MapNpc(MapNum).Npc(index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).Npc(index).StunTimer = GetTickCount
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "StunNPC", "modCombat", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
