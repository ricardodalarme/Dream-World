Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
    Dim FileName As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\logs\errors.txt"
    Open FileName For Append As #1
    Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
    Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
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

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not RAW Then
        If LenB(Dir$(App.Path & "\" & FileName)) > 0 Then
            FileExist = True
        End If
    Else
        If LenB(Dir$(FileName)) > 0 Then
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

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
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

    Call WritePrivateProfileString$(Header, Var, Value, File)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "PutVar", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveOptions()
    Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    Call PutVar(FileName, "Options", "Game_Name", Trim$(Options.Game_Name))
    Call PutVar(FileName, "Options", "Username", Trim$(Options.Username))
    Call PutVar(FileName, "Options", "Password", Trim$(Options.Password))
    Call PutVar(FileName, "Options", "SavePass", Str$(Options.SavePass))
    Call PutVar(FileName, "Options", "IP", Options.IP)
    Call PutVar(FileName, "Options", "Port", Str$(Options.Port))
    Call PutVar(FileName, "Options", "Music", Str$(Options.Music))
    Call PutVar(FileName, "Options", "Sound", Str$(Options.Sound))
    Call PutVar(FileName, "Options", "Debug", Str$(Options.Debug))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadOptions()
    Dim FileName As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\Data Files\config.ini"
    
    If Not FileExist(FileName, True) Then
        Options.Game_Name = "Eclipse Origins"
        Options.Password = vbNullString
        Options.SavePass = 0
        Options.Username = vbNullString
        Options.IP = "127.0.0.1"
        Options.Port = 7001
        Options.Music = 1
        Options.Sound = 1
        Options.Debug = 0
        SaveOptions
    Else
        Options.Game_Name = GetVar(FileName, "Options", "Game_Name")
        Options.Username = GetVar(FileName, "Options", "Username")
        Options.Password = GetVar(FileName, "Options", "Password")
        Options.SavePass = Val(GetVar(FileName, "Options", "SavePass"))
        Options.IP = GetVar(FileName, "Options", "IP")
        Options.Port = Val(GetVar(FileName, "Options", "Port"))
        Options.Music = GetVar(FileName, "Options", "Music")
        Options.Sound = GetVar(FileName, "Options", "Sound")
        Options.Debug = GetVar(FileName, "Options", "Debug")
    End If
    
    ' show in GUI
    If Options.Music = 0 Then
        frmMain.optMOff.Value = True
    Else
        frmMain.optMOn.Value = True
    End If
    
    If Options.Sound = 0 Then
        frmMain.optSOff.Value = True
    Else
        frmMain.optSOn.Value = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadOptions", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long
    Dim x As Long
    Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map.Name
    Put #f, , Map.Music
    Put #f, , Map.Revision
    Put #f, , Map.Moral
    Put #f, , Map.Panorama
    Put #f, , Map.Red
    Put #f, , Map.Green
    Put #f, , Map.Blue
    Put #f, , Map.Alpha
    Put #f, , Map.Fog
    Put #f, , Map.FogSpeed
    Put #f, , Map.FogOpacity
    Put #f, , Map.Up
    Put #f, , Map.Down
    Put #f, , Map.Left
    Put #f, , Map.Right
    Put #f, , Map.UpLeft
    Put #f, , Map.UpRight
    Put #f, , Map.DownLeft
    Put #f, , Map.DownRight
    Put #f, , Map.BootMap
    Put #f, , Map.BootX
    Put #f, , Map.BootY
    Put #f, , Map.MaxX
    Put #f, , Map.MaxY
    
    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            Put #f, , Map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #f, , Map.Npc(x)
    Next

    Close #f
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SaveMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long
    Dim x As Long
    Dim y As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
    ClearMap
    f = FreeFile
    Open FileName For Binary As #f
    Get #f, , Map.Name
    Get #f, , Map.Music
    Get #f, , Map.Revision
    Get #f, , Map.Moral
    Get #f, , Map.Panorama
    Get #f, , Map.Red
    Get #f, , Map.Green
    Get #f, , Map.Blue
    Get #f, , Map.Alpha
    Get #f, , Map.Fog
    Get #f, , Map.FogSpeed
    Get #f, , Map.FogOpacity
    Get #f, , Map.Up
    Get #f, , Map.Down
    Get #f, , Map.Left
    Get #f, , Map.Right
    Get #f, , Map.UpLeft
    Get #f, , Map.UpRight
    Get #f, , Map.DownLeft
    Get #f, , Map.DownRight
    Get #f, , Map.BootMap
    Get #f, , Map.BootX
    Get #f, , Map.BootY
    Get #f, , Map.MaxX
    Get #f, , Map.MaxY
        
    ' have to set the tile()
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    For x = 0 To Map.MaxX
        For y = 0 To Map.MaxY
            Get #f, , Map.Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Get #f, , Map.Npc(x)
    Next

    Close #f

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckDoors()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumDoors = 1
    
    Do While FileExist(GFX_PATH & "doors\" & NumDoors & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Door(0 To NumDoors)
        
        ' Set the surface
        Tex_Door(NumDoors) = CacheTexture("doors\" & NumDoors)
        NumDoors = NumDoors + 1
    Loop
    
    NumDoors = NumDoors - 1
     
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckDoors", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTilesets()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumTilesets = 1
    
    Do While FileExist(GFX_PATH & "Tilesets\" & NumTilesets & GFX_EXT)
        Call CacheTextures("Tilesets\", Tex_Tileset(), NumTilesets)
    Loop
    
    NumTilesets = NumTilesets - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTilesets", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckCharacters()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumCharacters = 1
    
    Do While FileExist(GFX_PATH & "characters\" & NumCharacters & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Character(0 To NumCharacters)
        
        ' Set the surface
        Tex_Character(NumCharacters) = CacheTexture("characters\" & NumCharacters)
        NumCharacters = NumCharacters + 1
    Loop
    
    NumCharacters = NumCharacters - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckCharacters", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPaperdolls()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumPaperdolls = 1
    
    Do While FileExist(GFX_PATH & "paperdolls\" & NumPaperdolls & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Paperdoll(0 To NumPaperdolls)
        
        ' Set the surface
        Tex_Paperdoll(NumPaperdolls) = CacheTexture("paperdolls\" & NumPaperdolls)
        NumPaperdolls = NumPaperdolls + 1
    Loop
    
    NumPaperdolls = NumPaperdolls - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPaperdolls", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckAnimations()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumAnimations = 1
    
    Do While FileExist(GFX_PATH & "animations\" & NumAnimations & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Animation(0 To NumAnimations)
        
        ' Set the surface
        Tex_Animation(NumAnimations) = CacheTexture("animations\" & NumAnimations)
        NumAnimations = NumAnimations + 1
    Loop
    
    NumAnimations = NumAnimations - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckAnimations", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckItems()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumItems = 1
    
    Do While FileExist(GFX_PATH & "items\" & NumItems & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Item(0 To NumItems)
        
        ' Set the surface
        Tex_Item(NumItems) = CacheTexture("items\" & NumItems)
        NumItems = NumItems + 1
    Loop
    
    NumItems = NumItems - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckSpells()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumSpells = 1
    
    Do While FileExist(GFX_PATH & "spells\" & NumSpells & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Spell(0 To NumSpells)
        
        ' Set the surface
        Tex_Spell(NumSpells) = CacheTexture("spells\" & NumSpells)
        NumSpells = NumSpells + 1
    Loop
    
    NumSpells = NumSpells - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckSpells", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckResources()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumResources = 1
    
    Do While FileExist(GFX_PATH & "resources\" & NumResources & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Resource(0 To NumResources)
        
        ' Set the surface
        Tex_Resource(NumResources) = CacheTexture("resources\" & NumResources)
        NumResources = NumResources + 1
    Loop
    
    NumResources = NumResources - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckResources", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFaces()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumFaces = 1
    
    Do While FileExist(GFX_PATH & "Faces\" & NumFaces & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Face(0 To NumFaces)
        
        ' Set the surface
        Tex_Face(NumFaces) = CacheTexture("Faces\" & NumFaces)
        NumFaces = NumFaces + 1
    Loop
    
    NumFaces = NumFaces - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFaces", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckPanoramas()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumPanoramas = 1
    
    Do While FileExist(GFX_PATH & "Panoramas\" & NumPanoramas & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Panorama(0 To NumPanoramas)
        
        ' Set the surface
        Tex_Panorama(NumPanoramas) = CacheTexture("Panoramas\" & NumPanoramas)
        NumPanoramas = NumPanoramas + 1
    Loop
    
    NumPanoramas = NumPanoramas - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckPanorama", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckFogs()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumFogs = 1
    
    Do While FileExist(GFX_PATH & "fogs\" & NumFogs & GFX_EXT)
        Call CacheTextures("fogs\", Tex_Fog(), NumFogs)
    Loop
    
    NumFogs = NumFogs - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckFog", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckTitles()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    NumTitles = 1
    
    Do While FileExist(GFX_PATH & "titles\" & NumTitles & GFX_EXT)
        ' Redim the surface
        ReDim Preserve Tex_Title(0 To NumTitles)
        
        ' Set the surface
        Tex_Title(NumTitles) = CacheTexture("titles\" & NumTitles)
        NumTitles = NumTitles + 1
    Loop
    
    NumTitles = NumTitles - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizePlayer(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Player(Index).Stat(Stats.Stat_Count - 1)
    ReDim Player(Index).Vital(Vitals.Vital_Count - 1)
    ReDim Player(Index).MaxVital(Vitals.Vital_Count - 1)
    ReDim Player(Index).Equipment(Equipment.Equipment_Count - 1)
    ReDim Player(Index).KillNpcs(MAX_NPCS)
    ReDim Player(Index).Quests(MAX_PLAYER_QUESTS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizePlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayers()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayers", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Player(Index)), LenB(Player(Index)))
    Call ResizePlayer(Index)
    Player(Index).Name = vbNullString
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearPlayer", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Item(Index).Add_Stat(Stats.Stat_Count - 1)
    ReDim Item(Index).Stat_Req(Stats.Stat_Count - 1)
    ReDim Item(Index).BagItem(MAX_BAG)
    ReDim Item(Index).BagValue(MAX_BAG)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Item(Index)), LenB(Item(Index)))
    Call ResizeItem(Index)
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString

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

Sub ClearAnimInstance(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(AnimInstance(Index)), LenB(AnimInstance(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearAnimInstance", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Animation(Index).Frames(1)
    ReDim Animation(Index).LoopCount(1)
    ReDim Animation(Index).looptime(1)
    ReDim Animation(Index).Sprite(1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeAnimation", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearAnimation(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Animation(Index)), LenB(Animation(Index)))
    Call ResizeAnimation(Index)
    Animation(Index).Name = vbNullString

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

Sub ResizeNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Npc(Index).DropChance(MAX_NPC_DROPS)
    ReDim Npc(Index).DropItem(MAX_NPC_DROPS)
    ReDim Npc(Index).DropItemValue(MAX_NPC_DROPS)
    ReDim Npc(Index).Stat(Stats.Stat_Count - 1)
    ReDim Npc(Index).Quest(MAX_NPC_QUESTS)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearNpc(ByVal Index As Long)
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Npc(Index)), LenB(Npc(Index)))
    Call ResizeNpc(Index)
    Npc(Index).Name = vbNullString
    Npc(Index).AttackSay = vbNullString
    Npc(Index).HP = 1
    
    For i = 1 To Stats.Stat_Count - 1
        Npc(Index).Stat(i) = 1
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

Sub ClearSpell(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Spell(Index)), LenB(Spell(Index)))
    Spell(Index).Name = vbNullString
    Spell(Index).Desc = vbNullString
    
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

Sub ClearShop(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Shop(Index)), LenB(Shop(Index)))
    ReDim Shop(Index).TradeItem(MAX_TRADES)
    Shop(Index).Name = vbNullString

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

Sub ClearResource(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Resource(Index)), LenB(Resource(Index)))
    Resource(Index).Name = vbNullString
    Resource(Index).SuccessMessage = vbNullString
    Resource(Index).EmptyMessage = vbNullString
    
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

Sub ClearMapItem(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItem", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Map), LenB(Map))
    Map.Name = vbNullString
    Map.MaxX = MAX_MAPX
    Map.MaxY = MAX_MAPY
    Map.FogOpacity = 255
    Map.Alpha = 255
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapItems()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapItems", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index)))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpc", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ClearMapNpcs()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearMapNpcs", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **********
' ** Door **
' **********
Sub ResizeDoor(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Door(Index).Stat_Req(Stats.Stat_Count - 1)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeDoor", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearDoor(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Door(Index)), LenB(Door(Index)))
    Call ResizeDoor(Index)
    Door(Index).Name = vbNullString
    
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

' ***************
' ** Tips **
' ***************
Public Sub CreateTipsINI()
    Dim FileName As String
    Dim File As String
    FileName = App.Path & "\data files\tips.ini"
    Max_Tips = 2

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not FileExist(FileName, True) Then
        File = FreeFile
        Open FileName For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxTips=" & Max_Tips
        Close File
    End If

    Exit Sub
errorhandler:
    HandleError "CreateTipsINI", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub LoadTips()
    Dim FileName As String
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\tips.ini"

    If CheckTips Then
        ReDim Tip(Max_Tips)
        Call SaveTips
    Else
        Max_Tips = Val(GetVar(FileName, "INIT", "MaxTips"))
        ReDim Tip(Max_Tips)
    End If

    Call ClearTips

    For i = 1 To Max_Tips
        Tip(i) = GetVar(FileName, "TIP" & i, "Text")
    Next
    
    Exit Sub
errorhandler:
    HandleError "LoadTips", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub SaveTips()
    Dim FileName As String
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\tips.ini"

    For i = 1 To Max_Tips
        Call PutVar(FileName, "TIP" & i, "Text", Trim$(Tip(i)))
    Next

    Exit Sub
errorhandler:
    HandleError "SaveTips", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function CheckTips() As Boolean
    Dim FileName As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    FileName = App.Path & "\data files\Tips.ini"

    If Not FileExist(FileName, True) Then
        Call CreateTipsINI
        CheckTips = True
    End If

    Exit Function
errorhandler:
    HandleError "CheckTips", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub ClearTips()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Max_Tips
        Tip(i) = vbNullString
    Next

    Exit Sub
errorhandler:
    HandleError "ClearTips", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub ResizeQuest(ByVal Index As Long)
    Dim i As Byte, x As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Quest(Index).VitalRew(Vitals.Vital_Count - 1)
    ReDim Quest(Index).StatRew(Stats.Stat_Count - 1)
    ReDim Quest(Index).StatReq(Stats.Stat_Count - 1)
    ReDim Quest(Index).Task(MAX_QUEST_TASKS)

    For i = 1 To MAX_QUEST_TASKS
        ReDim Quest(Index).Task(i).message(3)
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ResizeQuest", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearQuest(ByVal Index As Long)
    Dim i As Byte, x As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Quest(Index)), LenB(Quest(Index)))
    Call ResizeQuest(Index)
    Quest(Index).Name = vbNullString
    Quest(Index).Description = vbNullString
    
    For i = 1 To MAX_QUEST_TASKS
        For x = 1 To 3
            Quest(Index).Task(i).message(x) = vbNullString
        Next
    
        Quest(Index).Task(i).Num = 1
        Quest(Index).Task(i).Value = 1
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
Sub ResizeTitle(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ReDim Title(Index).VitalRew(Vitals.Vital_Count - 1)
    ReDim Title(Index).StatRew(Stats.Stat_Count - 1)
    ReDim Title(Index).StatReq(Stats.Stat_Count - 1)
    
    Exit Sub
errorhandler:
    HandleError "ResizeTitle", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Resize
    Exit Sub
End Sub

Sub ClearTitle(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call ZeroMemory(ByVal VarPtr(Title(Index)), LenB(Title(Index)))
    Call ResizeTitle(Index)
    Title(Index).Name = vbNullString
    
    ' Error handler
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

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ClearTitles", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

