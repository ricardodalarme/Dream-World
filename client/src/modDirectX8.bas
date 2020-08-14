Attribute VB_Name = "modDirectX8"
Option Explicit

' Master object
Public DX8 As DirectX8

' DirectX8 3D object
Public D3D8 As Direct3D8
Public D3DX As D3DX8

' DirectX8 device
Public D3DDevice As Direct3DDevice8

' DirectX8 window
Public D3DWindow As D3DPRESENT_PARAMETERS

' Render using pure 2D
Public D3DSprite As D3DXSprite
Public SpriteScaleVector As D3DVECTOR2

' Global texture
Public Texture() As TextureRec

' Textures
Public Tex_Item() As Long ' arrays
Public Tex_Spell() As Long
Public Tex_Character() As Long
Public Tex_Paperdoll() As Long
Public Tex_Tileset() As Long
Public Tex_Resource() As Long
Public Tex_Animation() As Long
Public Tex_Face() As Long
Public Tex_Panorama() As Long
Public Tex_Door() As Long
Public Tex_Title() As Long
Public Tex_Fog() As Long
Public Tex_Blood As Long ' singes
Public Tex_Misc As Long
Public Tex_Direction As Long
Public Tex_Target As Long
Public Tex_Bars As Long
Public Tex_Blank As Long
Public Tex_White As Long

' Number of graphic files
Public NumTextures As Long
Public NumTilesets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumSpells As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumFaces As Long
Public NumPanoramas As Long
Public NumDoors As Long
Public NumTitles As Long
Public NumFogs As Long

' Last texture loaded
Public CurrentTexture As Long

' Global Texture
Public Type TextureRec
    Tex As Direct3DTexture8
    Width As Long
    Height As Long
    FilePath As String
End Type

' Vertex
Public Type Vertex
    x As Single
    y As Single
    z As Single
    RHW As Single
    Color As Long
    TU As Single
    TV As Single
End Type

' Texture informations
Private Type ImageInfo
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

' ********************
' ** Initialization **
' ********************
Public Function InitDirectDraw() As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Init DirectX master object
    Set DX8 = New DirectX8
    
    ' Init DirectX 3D object
    Set D3DX = New D3DX8
    Set D3D8 = DX8.Direct3DCreate()
    
    ' Defines the window settings
    Call InitDirectWindow
    
    ' Test for set the DirectX8 device
    Call TryDeviceFlag
    
    ' Create pure 2d rendering objects
    Set D3DSprite = D3DX.CreateSprite(D3DDevice)
    
    With SpriteScaleVector
        .x = 1
        .y = 1
    End With

    ' Initialise texture effe
    InitD3DEffects

    ' Initialise the textures
    InitTextures
        
    ' Create the game font
    CreateFont

    ' Error handler
    Exit Function
errorhandler:
    HandleError "InitDirectDraw", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub InitD3DEffects()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' Now to tell directx which effects
    With D3DDevice
        ' Set directx vertex
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

        ' Alpha blender effects
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        
        ' Drawing effects
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitD3DEffects", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub InitDirectWindow()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    With D3DWindow
        ' Back buffer
        .BackBufferCount = 1
        .BackBufferFormat = D3DFMT_X8R8G8B8
        .BackBufferWidth = ScreenX
        .BackBufferHeight = ScreenY
        
        ' Efects
        .SwapEffect = D3DSWAPEFFECT_COPY

        ' The window
        .hDeviceWindow = frmMain.picScreen.hWnd
        .Windowed = True
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitD3DEffects", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub TryDeviceFlag()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
        
    ' Test for set the DirectX8 device
    If Not InitDirectXDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
        If Not InitDirectXDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not InitDirectXDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                If Not InitDirectXDevice(D3DCREATE_PUREDEVICE) Then
                    If Not InitDirectXDevice(D3DCREATE_FPU_PRESERVE) Then
                        MsgBox "Error initializing DirectX8."
                        DestroyGame
                    End If
                End If
            End If
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "TryDeviceFlag", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function InitDirectXDevice(Flag As CONST_D3DCREATEFLAGS) As Boolean
    ' If have error exit function
    On Error GoTo errorhandler

    ' Create DirectX8 device
    Set D3DDevice = D3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, Flag, D3DWindow)

    ' Return function value
    InitDirectXDevice = True
    
errorhandler:
    Exit Function
End Function

Private Sub InitTextures()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' load arrays textures
    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpells
    Call CheckFaces
    Call CheckPanoramas
    Call CheckDoors
    Call CheckTitles
    Call CheckFogs
    
    ' load singles textures
    Tex_Direction = CacheTexture("direction")
    Tex_Blood = CacheTexture("blood")
    Tex_Target = CacheTexture("target")
    Tex_Bars = CacheTexture("bars")
    Tex_Blank = CacheTexture("blank")
    Tex_White = CacheTexture("white")
    Tex_Misc = CacheTexture("misc")

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "InitTextures", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Initializing a texture
Public Function CacheTexture(ByVal FileName As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Set the texture path
    FileName = App.Path & GFX_PATH & FileName & GFX_EXT
    
    ' Prevent subscript out range
    If Not FileExist(FileName, True) Then Exit Function
    
    ' Set the max textures
    NumTextures = NumTextures + 1
    ReDim Preserve Texture(NumTextures)

    ' Set the texture path
    Texture(NumTextures).FilePath = FileName
    
    ' Load texture
    LoadTexture NumTextures
    
    ' Return function value
    CacheTexture = NumTextures

    ' Error handler
    Exit Function
errorhandler:
    HandleError "CacheTexture", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub CacheTextures(ByVal FileName As String, ByRef Tex() As Long, ByRef Count As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Set the texture path
    FileName = App.Path & GFX_PATH & FileName & Count & GFX_EXT
    
    ' Prevent subscript out range
    If Not FileExist(FileName, True) Then Exit Sub
    
    ' Set the max textures
    NumTextures = NumTextures + 1
    ReDim Preserve Texture(NumTextures)

    ' Set the texture path
    Texture(NumTextures).FilePath = FileName
    
    ' Load texture
    LoadTexture NumTextures
    
    ' Return values
    ReDim Preserve Tex(Count)
    Tex(Count) = NumTextures
    Count = Count + 1
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CacheTextures", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub SetTexture(ByVal TextureNum As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out of range
    If TextureNum > UBound(Texture) Or TextureNum < 0 Then Exit Sub

    ' Set texture
    If TextureNum <> CurrentTexture Then
        Call D3DDevice.SetTexture(0, Texture(TextureNum).Tex)
        CurrentTexture = TextureNum
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "SetTexture", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadTexture(ByVal TextureNum As Long)
    Dim Tex_Info As ImageInfo

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Create texture
    Set Texture(TextureNum).Tex = D3DX.CreateTextureFromFileEx(D3DDevice, Texture(TextureNum).FilePath, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, RGB(255, 255, 255), Tex_Info, ByVal 0)
                
    ' Set texture size
    Texture(TextureNum).Height = Tex_Info.Height
    Texture(TextureNum).Width = Tex_Info.Width

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadTexture", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DestroyDirectDraw()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Unload textures
    UnloadTextures
    
    ' Unload DirectX8 object
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    If Not D3D8 Is Nothing Then Set D3D8 = Nothing
    If Not DX8 Is Nothing Then Set DX8 = Nothing

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroyDirectDraw", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UnloadTextures()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Reload the textures
    If NumTextures > 0 Then
        For i = 1 To NumTextures
            If Not Texture(i).Tex Is Nothing Then Set Texture(i).Tex = Nothing
        Next
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UnloadTexutres", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' **************
' ** Blitting **
' **************
Public Sub RenderSprite(ByVal TextureNum As Long, ByVal x As Single, ByVal y As Single, ByVal SrcX As Single, ByVal SrcY As Single, ByVal Width As Single, ByVal Height As Single, Optional ByVal Color As Long = -1, Optional ByVal Rotation As Single = 0)
    Dim SrcRect As RECT
    Dim VertexSize As D3DVECTOR2
    Dim VertexPoint As D3DVECTOR2

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If TextureNum <= 0 Then Exit Sub

    ' Set the texture
    SetTexture TextureNum
    
    ' Create rectangle
    With SrcRect
        .Left = SrcX
        .top = SrcY
        .Right = .Left + Width
        .bottom = .top + Height
    End With
    
    ' Size
    With VertexSize
        .x = Width
        .y = Height
    End With

    ' Location
    With VertexPoint
        .x = x
        .y = y
    End With

    ' Render the texture
    D3DSprite.Draw Texture(TextureNum).Tex, SrcRect, SpriteScaleVector, VertexSize, Rotation, VertexPoint, Color

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderSprite", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderTexture(ByVal TextureNum As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, Optional ByVal Colour As Long = -1)
    Dim Box(0 To 3) As Vertex, TextureWidth As Long, TextureHeight As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If TextureNum <= 0 Then Exit Sub
    
    ' load the texture
    Call SetTexture(TextureNum)

    ' texture sizes
    TextureWidth = Texture(TextureNum).Width
    TextureHeight = Texture(TextureNum).Height

    ' Create the vertex
    Box(0) = CreateVertex(DestX, DestY, 1, Colour, (SrcX / TextureWidth), (SrcY / TextureHeight))
    Box(1) = CreateVertex(DestX + DestWidth, Box(0).y, Box(0).RHW, Box(0).Color, (SrcX + SrcWidth) / TextureWidth, Box(0).TV)
    Box(2) = CreateVertex(Box(0).x, DestY + DestHeight, Box(0).RHW, Box(0).Color, Box(0).TU, (SrcY + SrcHeight) / TextureHeight)
    Box(3) = CreateVertex(Box(1).x, Box(2).y, Box(0).RHW, Box(0).Color, Box(1).TU, Box(2).TV)

    ' Render the texture
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTexture", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub RenderTextureByRects(TextureNum As Long, sRECT As RECT, dRect As RECT, Optional Colour As Long = -1)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    RenderTexture TextureNum, dRect.Left, dRect.top, sRECT.Left, sRECT.top, dRect.Right - dRect.Left, dRect.bottom - dRect.top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.top, Colour

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "RenderTextureByRects", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function CreateVertex(ByVal x As Long, ByVal y As Long, ByVal RHW As Single, ByVal Color As Long, ByVal TU As Single, ByVal TV As Single) As Vertex
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return the new vertex
    With CreateVertex
        .x = x
        .y = y
        .RHW = RHW
        .Color = Color
        .TU = TU
        .TV = TV
    End With
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "CreateVertex", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function DX8Color(ByVal ColourNum As Long, Optional ByVal Alpha As Long = 255) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Return function value
    Select Case ColourNum
        Case Black
            DX8Color = D3DColorARGB(Alpha, 0, 0, 0)
        Case Blue
            DX8Color = D3DColorARGB(Alpha, 16, 104, 237)
        Case Green
            DX8Color = D3DColorARGB(Alpha, 119, 188, 84)
        Case Cyan
            DX8Color = D3DColorARGB(Alpha, 16, 224, 237)
        Case Red
            DX8Color = D3DColorARGB(Alpha, 201, 0, 0)
        Case Magenta
            DX8Color = D3DColorARGB(Alpha, 255, 0, 255)
        Case Brown
            DX8Color = D3DColorARGB(Alpha, 175, 149, 92)
        Case Grey
            DX8Color = D3DColorARGB(Alpha, 192, 192, 192)
        Case DarkGrey
            DX8Color = D3DColorARGB(Alpha, 128, 128, 128)
        Case BrightBlue
            DX8Color = D3DColorARGB(Alpha, 126, 182, 240)
        Case BrightGreen
            DX8Color = D3DColorARGB(Alpha, 126, 240, 137)
        Case BrightCyan
            DX8Color = D3DColorARGB(Alpha, 157, 242, 242)
        Case BrightRed
            DX8Color = D3DColorARGB(Alpha, 255, 0, 0)
        Case Pink
            DX8Color = D3DColorARGB(Alpha, 255, 118, 221)
        Case Yellow
            DX8Color = D3DColorARGB(Alpha, 255, 255, 0)
        Case White
            DX8Color = D3DColorARGB(Alpha, 255, 255, 255)
    End Select

    ' Error handler
    Exit Function
errorhandler:
    HandleError "DX8Color", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

' *************************
' ** Rendering in screen **
' *************************
Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
    Dim rec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Rectangle
    With rec
        .top = 24
        .Left = 0
        .Right = .Left + 32
        .bottom = .top + 32
    End With
    
    ' Render
    RenderSprite Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), rec.Left, rec.top, rec.Right, rec.bottom

    ' render dir blobs
    For i = 1 To 4
        With rec
            .Left = (i - 1) * 8
            .Right = 8
            
            ' find out whether render blocked or not
            If Not isDirBlocked(Map.Tile(x, y).DirBlock, CByte(i)) Then
                .top = 8
            Else
                .top = 16
            End If
            .bottom = 8
        End With
        
        ' Render
        RenderSprite Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), rec.Left, rec.top, rec.Right, rec.bottom
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDirection", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTileOutline()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optBlock.Value Then Exit Sub

    Call RenderSprite(Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), 0, 0, PIC_X, PIC_Y, D3DColorARGB(200, 255, 255, 255))
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltTileOutline", "modDirectDraw7", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapTile(ByVal x As Long, ByVal y As Long, Optional ByVal ScreenShot As Boolean = False)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(x, y)
        For i = 1 To MAX_MAP_LAYERS
            ' skip tile?
            If (.Layer(Ground, i).Tileset > 0 And .Layer(Ground, i).Tileset <= NumTilesets) And (.Layer(Ground, i).x >= 0 Or .Layer(Ground, i).y >= 0) Then
                ' render
                If Not ScreenShot Then
                    RenderSprite Tex_Tileset(.Layer(Ground, i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(Ground, i).x * PIC_X, .Layer(Ground, i).y * PIC_Y, PIC_X, PIC_Y
                Else
                    'Tex_Map.Surf.DrawFast X * PIC_X, Y * PIC_Y, Texture(Tex_Tileset(.Layer(Ground, i).Tileset)).Surf, Rec, CONST_DDDrawFASTFLAGS.DDDrawFAST_SRCCOLORKEY
                End If
            End If
        Next
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapTile", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapFringeTile(ByVal x As Long, ByVal y As Long, Optional ByVal ScreenShot As Boolean = False)
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    With Map.Tile(x, y)
        For i = 1 To MAX_MAP_LAYERS
            ' skip tile?
            If (.Layer(Fringe, i).Tileset > 0 And .Layer(Fringe, i).Tileset <= NumTilesets) And (.Layer(Fringe, i).x >= 0 Or .Layer(Fringe, i).y >= 0) Then
                ' render
                If ScreenShot = False Then
                    RenderSprite Tex_Tileset(.Layer(Fringe, i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(Fringe, i).x * PIC_X, .Layer(Fringe, i).y * PIC_Y, PIC_X, PIC_Y
                Else
                    'Tex_Map.Surf.DrawFast X * PIC_X, Y * PIC_Y, Texture(Tex_Tileset(.Layer(Fringe, i).Tileset)).Surf, Rec, CONST_DDDrawFASTFLAGS.DDDrawFAST_SRCCOLORKEY
                End If
            End If
        Next
    End With
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapFringeTile", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
    Dim Width As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Texture width
    Width = Texture(Tex_Target).Width / 2

    ' Render
    RenderSprite Tex_Target, ConvertMapX(x - ((Width - 32) / 2)), ConvertMapY(y), 0, 0, Width, Texture(Tex_Target).Height
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTarget", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal target As Long, ByVal x As Long, ByVal y As Long)
    Dim Width As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Size
    Width = Texture(Tex_Target).Width / 2

    ' Render
    RenderSprite Tex_Target, ConvertMapX(x - ((Width - 32) / 2)), ConvertMapY(y), Width, 0, Width, Texture(Tex_Target).Height
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHover", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBlood(ByVal Index As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' check if we should be seeing it
    If Blood(Index).Timer + 20000 < GetTickCount Then Exit Sub
        
    ' Render
    RenderSprite Tex_Blood, ConvertMapX(Blood(Index).x * PIC_X), ConvertMapY(Blood(Index).y * PIC_Y), (Blood(Index).Sprite - 1) * PIC_X, 0, PIC_X, PIC_Y

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBlood", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawBars()
    Dim tmpY As Long, tmpX As Long
    Dim sWidth As Long, sHeight As Long
    Dim sRECT As RECT
    Dim barWidth As Long
    Dim i As Long, npcNum As Long, partyIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' dynamic bar calculations
    sWidth = Texture(Tex_Bars).Width
    sHeight = Texture(Tex_Bars).Height / 4
    
    ' render health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).Num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).XOffset + 16 - (sWidth / 2)
                tmpY = MapNpc(i).y * PIC_Y + MapNpc(i).YOffset + 35
                
                ' calculate the width to fill
                barWidth = ((MapNpc(i).Vital(Vitals.HP) / sWidth) / (Npc(npcNum).HP / sWidth)) * sWidth
                
                ' draw bar background
                With sRECT
                    .top = sHeight * 1 ' HP bar background
                    .Left = 0
                    .Right = sWidth
                    .bottom = sHeight
                End With
                RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right, sRECT.bottom
                
                ' draw the bar proper
                With sRECT
                    .top = 0 ' HP bar
                    .Left = 0
                    .Right = barWidth
                    .bottom = sHeight
                End With
                RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right, sRECT.bottom
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).YOffset + 35 + sHeight + 1
            
            ' calculate the width to fill
            barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' draw bar background
            With sRECT
                .top = sHeight * 3 ' cooldown bar background
                .Left = 0
                .Right = sWidth
                .bottom = .top + sHeight
            End With
            
            ' Render
            RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.top
            
            ' draw the bar proper
            With sRECT
                .top = sHeight * 2 ' cooldown bar
                .Left = 0
                .Right = barWidth
                .bottom = .top + sHeight
            End With
            
            ' Render
            RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.top
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).XOffset + 16 - (sWidth / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).YOffset + 35
       
        ' calculate the width to fill
        barWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / sWidth)) * sWidth
       
        ' Rectangle
        With sRECT
            .top = sHeight * 1 ' HP bar background
            .Left = 0
            .Right = sWidth
            .bottom = sHeight
        End With
        
        ' Render
        RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right, sRECT.bottom
       
        ' Rectangle
        With sRECT
            .top = 0 ' HP bar
            .Left = 0
            .Right = barWidth
            .bottom = sHeight
        End With
        
        ' Render
        RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right, sRECT.bottom
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).XOffset + 16 - (sWidth / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).YOffset + 35
                    
                    ' calculate the width to fill
                    barWidth = ((GetPlayerVital(partyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' Rectangle
                    With sRECT
                        .top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .bottom = .top + sHeight
                    End With
                    
                    ' Render
                    RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.top
                    
                    ' Rectangle
                    With sRECT
                        .top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + barWidth
                        .bottom = .top + sHeight
                    End With
                    
                    ' Render
                    RenderSprite Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.top
                End If
            End If
        Next
    End If
                    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBars", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim Sprite As Long
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim looptime As Long
    Dim FrameCount As Long
    Dim x As Long, y As Long
    Dim lockindex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Clear animation?
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    ' Declarations
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    ' total width divided by frame count
    Width = Texture(Tex_Animation(Sprite)).Width / FrameCount
    Height = Texture(Tex_Animation(Sprite)).Height
    
    ' Rectangle
    With sRECT
        .top = 0
        .bottom = Height
        .Left = (AnimInstance(Index).FrameIndex(Layer) - 1) * Width
        .Right = sRECT.Left + Width
    End With
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).XOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (Height / 2) + Player(lockindex).YOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).Num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).XOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (Height / 2) + MapNpc(lockindex).YOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (Height / 2)
    End If
    
    ' Render
    RenderSprite Tex_Animation(Sprite), ConvertMapX(x), ConvertMapY(y), sRECT.Left, sRECT.top, sRECT.Right - sRECT.Left, sRECT.bottom - sRECT.top
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawAnimation", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItem(ByVal itemnum As Long)
    Dim PicNum As Long
    Dim rec As RECT
    Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' get the picture
    PicNum = Item(MapItem(itemnum).Num).Pic

    ' Prevent subscript ou range
    If PicNum < 1 Or PicNum > NumItems Then Exit Sub

    ' Set rectangle
    With rec
        .top = 0
        .bottom = 32
        .Left = 32
        .Right = 64
    End With
    
    ' Render
    RenderSprite Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), rec.Left, rec.top, rec.Right - rec.Left, rec.bottom - rec.top

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItem", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayer(ByVal Index As Long)
    Dim anim As Byte, i As Long, x As Long, y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Player sprite
    Sprite = GetPlayerSprite(Index)

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    ' Reset frame
    anim = Player(Index).Step

    ' Check for attacking animation
    If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If Player(Index).Attacking = 1 Then
            anim = 3
        End If
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP, DIR_UP_LEFT, DIR_UP_RIGHT
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN, DIR_DOWN_LEFT, DIR_DOWN_RIGHT
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    ' Rectangle
    With rec
        .top = spritetop * (Texture(Tex_Character(Sprite)).Height / 4)
        .bottom = (Texture(Tex_Character(Sprite)).Height / 4)
        .Left = anim * (Texture(Tex_Character(Sprite)).Width / 4)
        .Right = (Texture(Tex_Character(Sprite)).Width / 4)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset - ((Texture(Tex_Character(Sprite)).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Texture(Tex_Character(Sprite)).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - ((Texture(Tex_Character(Sprite)).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
    End If

    ' render the actual sprite
    Call DrawSprite(Sprite, x, y, rec)
    
    ' check for paperdolling
    For i = 1 To UBound(PaperdollOrder)
        If GetPlayerEquipment(Index, PaperdollOrder(i)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll > 0 Then
                Call DrawPaperdoll(x, y, Item(GetPlayerEquipment(Index, PaperdollOrder(i))).Paperdoll, anim, spritetop)
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayer", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim anim As Byte, i As Long, x As Long, y As Long, Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim attackspeed As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If MapNpc(MapNpcNum).Num = 0 Then Exit Sub ' no npc set
    
    ' Npc sprite
    Sprite = Npc(MapNpc(MapNpcNum).Num).Sprite

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' Npc attack speed
    attackspeed = 1000

    ' Reset frame
    anim = MapNpc(MapNpcNum).Step
    
    ' Check for attacking animation
    If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
        If MapNpc(MapNpcNum).Attacking = 1 Then
            anim = 3
        End If
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
        Case DIR_UP_LEFT
            spritetop = 3
        Case DIR_UP_RIGHT
            spritetop = 3
        Case DIR_DOWN_LEFT
            spritetop = 0
        Case DIR_DOWN_RIGHT
            spritetop = 0
    End Select

    ' Rectangle
    With rec
        .top = (Texture(Tex_Character(Sprite)).Height / 4) * spritetop
        .bottom = Texture(Tex_Character(Sprite)).Height / 4
        .Left = anim * (Texture(Tex_Character(Sprite)).Width / 4)
        .Right = (Texture(Tex_Character(Sprite)).Width / 4)
    End With

    ' Calculate the X
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset - ((Texture(Tex_Character(Sprite)).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Texture(Tex_Character(Sprite)).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - ((Texture(Tex_Character(Sprite)).Height / 4) - 32)
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
    End If

    ' Render
    Call DrawSprite(Sprite, x, y, rec)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpc", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPaperdoll(ByVal x2 As Long, ByVal y2 As Long, ByVal Sprite As Long, ByVal anim As Long, ByVal spritetop As Long)
    Dim rec As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    ' Rectangle
    With rec
        .top = spritetop * (Texture(Tex_Paperdoll(Sprite)).Height / 4)
        .bottom = (Texture(Tex_Paperdoll(Sprite)).Height / 4)
        .Left = anim * (Texture(Tex_Paperdoll(Sprite)).Width / 4)
        .Right = (Texture(Tex_Paperdoll(Sprite)).Width / 4)
    End With
    
    ' Render
    RenderSprite Tex_Paperdoll(Sprite), ConvertMapX(x2), ConvertMapY(y2), rec.Left, rec.top, rec.Right, rec.bottom

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPaperdoll", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal x2 As Long, y2 As Long, rec As RECT)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' Render
    RenderSprite Tex_Character(Sprite), ConvertMapX(x2), ConvertMapY(y2), rec.Left, rec.top, rec.Right, rec.bottom

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSprite", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapResource(ByVal Resource_Num As Long, Optional ByVal ScreenShot As Boolean = False)
    Dim Resource_master As Long
    Dim Resource_state As Long
    Dim Resource_sprite As Long
    Dim rec As RECT
    Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapResource(Resource_Num).x > Map.MaxX Then Exit Sub
    If MapResource(Resource_Num).y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(MapResource(Resource_Num).x, MapResource(Resource_Num).y).Data1
    
    ' Prevent subscript out range
    If Resource_master = 0 Then Exit Sub
    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    
    ' Get the Resource state
    Resource_state = MapResource(Resource_Num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' cut down everything if we're editing
    If InMapEditor Then
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If

    ' src rect
    With rec
        .top = 0
        .bottom = Texture(Tex_Resource(Resource_sprite)).Height
        .Left = 0
        .Right = Texture(Tex_Resource(Resource_sprite)).Width
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_Num).x * PIC_X) - (Texture(Tex_Resource(Resource_sprite)).Width / 2) + 16
    y = (MapResource(Resource_Num).y * PIC_Y) - Texture(Tex_Resource(Resource_sprite)).Height + 32
    
    ' render it
    If Not ScreenShot Then
        RenderSprite Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.top, rec.Right, rec.bottom
    Else
        'Call ScreenshotResource(Resource_sprite, x, y, rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapResource", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapDoor(ByVal Door_Num As Long, Optional ByVal ScreenShot As Boolean = False)
    Dim Door_master As Long
    Dim Door_sprite As Long
    Dim rec As RECT
    Dim x As Long, y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' make sure it's not out of map
    If MapDoor(Door_Num).x > Map.MaxX Then Exit Sub
    If MapDoor(Door_Num).y > Map.MaxY Then Exit Sub
    
    ' Get the Door type
    Door_master = Map.Tile(MapDoor(Door_Num).x, MapDoor(Door_Num).y).Data1
    
    ' Prevent subscript out range
    If Door_master = 0 Then Exit Sub

    ' Get the door state
    If MapDoor(Door_Num).State = 0 Then Door_sprite = Door(Door_master).ClosedImage Else Door_sprite = Door(Door_master).OpeningImage
    
    ' Prevent subscript out range
    If Door_sprite = 0 Then Exit Sub
    
    ' src rect
    With rec
        .top = 0
        .bottom = Texture(Tex_Door(Door_sprite)).Height
        .Left = 0
        .Right = Texture(Tex_Door(Door_sprite)).Width
    End With

    ' Set base x + y, then the offset due to size
    x = (MapDoor(Door_Num).x * PIC_X) - (Texture(Tex_Door(Door_sprite)).Width / 2) + 16
    y = (MapDoor(Door_Num).y * PIC_Y) - Texture(Tex_Door(Door_sprite)).Height + 32
    
    ' render it
    If Not ScreenShot Then
        Call RenderSprite(Tex_Door(Door_sprite), ConvertMapX(x), ConvertMapY(y), rec.Left, rec.top, rec.Right, rec.bottom)
    Else
        'Call ScreenshotDoor(Door_sprite, x, y, rec)
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapDoor", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawPanorama()
    Dim rec As RECT
    Dim PanoramaNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Declaration
    PanoramaNum = Map.Panorama
    
    ' Prevent subscript out range
    If PanoramaNum <= 0 Or PanoramaNum > NumPanoramas Then Exit Sub
    
    With rec
        .top = 0
        .bottom = ScreenY
        .Left = 0
        .Right = ScreenX
    End With

    ' Render
    RenderSprite Tex_Panorama(PanoramaNum), rec.Left, rec.top, rec.Left, rec.top, rec.Right, rec.bottom

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPanorama", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
    Dim Colour As Long, x As Long, y As Long, renderState As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Map.Fog <= 0 Or Map.Fog > NumFogs Then Exit Sub

    For x = 0 To ((Map.MaxX * 32) / 256) + 1
        For y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(Map.Fog), ConvertMapX((x * 256) + fogOffsetX), ConvertMapY((y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, D3DColorARGB(Map.FogOpacity, 255, 255, 255)
        Next
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawFog", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' ***************************
' ** Rendering in pictures **
' ***************************
' Game Editors
Public Sub EditorMap_DrawTileset()
    Dim Height As Long, Width As Long
    Dim Tileset As Long
    Dim SrcRec As RECT, DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InMapEditor Then Exit Sub
    
    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' Prevent subscript out range
    If Tileset < 1 Or Tileset > NumTilesets Then
        frmEditor_Map.picBack.Cls
        Exit Sub
    End If
 
    ' Texture sizes
    Height = Texture(Tex_Tileset(Tileset)).Height
    Width = Texture(Tex_Tileset(Tileset)).Width

    ' Rectangles
    With SrcRec
        .Left = frmEditor_Map.scrlPictureX.Value * PIC_X
        .top = frmEditor_Map.scrlPictureY.Value * PIC_Y
        .Right = SrcRec.Left + Width
        .bottom = SrcRec.top + Height
    End With

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    ' Render tileset
    RenderSprite Tex_Tileset(Tileset), 0, 0, SrcRec.Left, SrcRec.top, Width, Height

    ' draw boxes
    EditorMap_DrawBox (EditorTileX * PIC_X) - SrcRec.Left, (EditorTileY * PIC_Y) - SrcRec.top, (EditorTileWidth * 32), (EditorTileHeight * 32), DX8Color(Red)
    RenderSprite Tex_Blank, shpLocLeft, shpLocTop, 0, 0, PIC_X, PIC_Y, DX8Color(Blue)
    
    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Map.picBack.Width
        .Y1 = 0
        .y2 = frmEditor_Map.picBack.Height
    End With

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Map.picBack.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawTileset", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawBox(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, Optional ByVal Colour As Long = 0)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InMapEditor Then Exit Sub
    
    ' Draw rectangle
    RenderSprite Tex_Blank, x, y, 0, 0, w, 1, Colour
    RenderSprite Tex_Blank, x, y, 0, 0, 1, h, Colour
    RenderSprite Tex_Blank, x, y + h, 0, 0, w, 1, Colour
    RenderSprite Tex_Blank, x + w, y, 0, 0, 1, h, Colour
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawBox", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorMap_DrawMapItem()
    Dim itemnum As Long
    Dim DestRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InMapEditor Then Exit Sub
    
    ' Icon of item
    itemnum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    ' Prevent subscript out range
    If itemnum < 1 Or itemnum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' Render
    RenderSprite Tex_Item(itemnum), 0, 0, 32, 0, 32, 64

    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Map.Width
        .Y1 = 0
        .y2 = frmEditor_Map.Height
    End With
                    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Map.picMapItem.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorMap_DrawMapItem", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorSpell_DrawIcon()
    Dim IconNum As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_SPELL Then Exit Sub
    
    ' Icon of spell
    IconNum = frmEditor_Spell.scrlIcon.Value
 
    ' Prevent subscript out range
    If IconNum < 1 Or IconNum > NumSpells Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If

    ' Start rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' Render
    RenderSprite Tex_Spell(IconNum), 0, 0, 0, 0, PIC_X, PIC_Y
    
    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Spell.picSprite.Width
        .Y1 = 0
        .y2 = frmEditor_Spell.picSprite.Height
    End With
    
    ' Etart rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Spell.picSprite.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorSpell_DrawIcon", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawIcon()
    Dim IconNum As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_ITEM Then Exit Sub
    
    ' Icon of icon
    IconNum = frmEditor_Item.scrlPic.Value
    
    ' Prevent subscript out range
    If IconNum < 1 Or IconNum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ' Start rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' Render
    RenderSprite Tex_Item(IconNum), 0, 0, 0, 0, PIC_X, PIC_Y
    
    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Item.picItem.Width
        .Y1 = 0
        .y2 = frmEditor_Item.picItem.Height
    End With
    
    ' Etart rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Item.picItem.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawIcon", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorItem_DrawPaperdoll()
    Dim Sprite As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_ITEM Then Exit Sub
    
    ' The paperdoll
    Sprite = frmEditor_Item.scrlPaperdoll.Value

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' Render!
    RenderSprite Tex_Paperdoll(Sprite), 0, 0, 0, 0, Texture(Tex_Paperdoll(Sprite)).Width, Texture(Tex_Paperdoll(Sprite)).Height
                    
    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Item.picPaperdoll.Width
        .Y1 = 0
        .y2 = frmEditor_Item.picPaperdoll.Height
    End With
    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Item.picPaperdoll.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorItem_DrawPaperdoll", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorAnim_DrawAnim()
    Dim AnimationNum As Long
    Dim DestRect As D3DRECT
    Dim i As Long
    Dim Width As Long, Height As Long
    Dim looptime As Long
    Dim FrameCount As Long
    Dim ShouldRender As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_ANIMATION Then Exit Sub
    
    For i = 0 To 1
        ' Sprite of animation
        AnimationNum = frmEditor_Animation.scrlSprite(i).Value
        
        ' Prevent subscript out range
        If AnimationNum < 1 Or AnimationNum > NumAnimations Then
            frmEditor_Animation.picSprite(i).Cls
        Else
            ' Declarations
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ' Clear
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).Value > 0 Then
                    ' total width divided by frame count
                    Width = Texture(Tex_Animation(AnimationNum)).Width / frmEditor_Animation.scrlFrameCount(i).Value
                    Height = Texture(Tex_Animation(AnimationNum)).Height
                    
                    ' Init rendering
                    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    D3DDevice.BeginScene
                    
                    ' Render!
                    RenderSprite Tex_Animation(AnimationNum), 0, 0, (AnimEditorFrame(i) - 1) * Width, 0, (AnimEditorFrame(i) - 1) * Width + Width, Height
                    
                    ' Rectangle
                    With DestRect
                        .X1 = 0
                        .x2 = frmEditor_Animation.picSprite(i).Width
                        .Y1 = 0
                        .y2 = frmEditor_Animation.picSprite(i).Height
                    End With
                                
                    ' End rendering
                    D3DDevice.EndScene
                    D3DDevice.Present DestRect, DestRect, frmEditor_Animation.picSprite(i).hWnd, ByVal (0)
                End If
            End If
        End If
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorAnim_DrawAnim", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorTitle_DrawIcon()
    Dim IconNum As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_TITLE Then Exit Sub
    
    ' Icon of title
    IconNum = frmEditor_Title.scrlIcon.Value
    
    ' Prevent subscript out range
    If IconNum < 1 Or IconNum > NumTitles Then
        frmEditor_Title.picIcon.Cls
        Exit Sub
    End If

    ' Start rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' Render
    RenderSprite Tex_Title(IconNum), 0, 0, 0, 0, PIC_X, PIC_Y
    
    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Title.picIcon.Width
        .Y1 = 0
        .y2 = frmEditor_Title.picIcon.Height
    End With
    
    ' Etart rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Title.picIcon.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorTitle_DrawIcon", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorNpc_DrawSprite()
    Dim Sprite As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_NPC Then Exit Sub
    
    ' Sprite of npc
    Sprite = frmEditor_NPC.scrlSprite.Value

    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' render!
    RenderSprite Tex_Character(Sprite), 0, 0, 0, 0, Texture(Tex_Character(Sprite)).Width / 4, Texture(Tex_Character(Sprite)).Height / 4
                  
    ' Rentangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_NPC.picSprite.Width
        .Y1 = 0
        .y2 = frmEditor_NPC.picSprite.Height
    End With
    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_NPC.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorNpc_DrawSprite", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorDoor_DrawOpeningImage()
    Dim Sprite As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_DOOR Then Exit Sub
    
    ' Sprite of door
    Sprite = frmEditor_Door.scrlOpeningImage.Value
    
    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumDoors Then
        frmEditor_Door.picOpeningImage.Cls
        Exit Sub
    End If
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    ' Render
    Call RenderSprite(Tex_Door(Sprite), 0, 0, 0, 0, Texture(Tex_Door(Sprite)).Width, Texture(Tex_Door(Sprite)).Height)
    
    ' Rentangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Door.picOpeningImage.Width
        .Y1 = 0
        .y2 = frmEditor_Door.picOpeningImage.Height
    End With
    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Door.picOpeningImage.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorDoor_DrawOpeningImage", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub EditorDoor_DrawClosedImage()
    Dim Sprite As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not CurrentEditor = EDITOR_DOOR Then Exit Sub
    
    ' Sprite of door
    Sprite = frmEditor_Door.scrlClosedImage.Value
    
    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumDoors Then
        frmEditor_Door.picClosedImage.Cls
        Exit Sub
    End If

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    ' Render
    Call RenderSprite(Tex_Door(Sprite), 0, 0, 0, 0, Texture(Tex_Door(Sprite)).Width, Texture(Tex_Door(Sprite)).Height)
    
    ' Rentangle
    With DestRect
        .X1 = 0
        .x2 = frmEditor_Door.picClosedImage.Width
        .Y1 = 0
        .y2 = frmEditor_Door.picClosedImage.Height
    End With
    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmEditor_Door.picClosedImage.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "EditorDoor_DrawClosedImage", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' frmMain
Sub DrawInventory()
    Dim i As Long, x As Long, y As Long, itemnum As Long, itempic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT, DestRect As RECT
    Dim Colour As Long
    Dim tmpItem As Long, amountModifier As Long

    ' Don't render
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not InGame Or frmMain.picInventory.Visible = False Then Exit Sub

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    For i = 1 To MAX_INV
        ' The item
        itemnum = GetPlayerInvItemNum(MyIndex, i)
            
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            amountModifier = 0
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    tmpItem = TradeYourOffer(x).Num
                    If tmpItem = itemnum Then
                        ' check if currency
                        If Not Item(tmpItem).Type = ITEM_TYPE_CURRENCY Then
                            ' normal item, exit out
                            Exit For
                        Else
                            ' if amount = all currency, remove from inventory
                            If TradeYourOffer(x).Value = GetPlayerInvItemValue(MyIndex, i) Then
                                Exit Sub
                            Else
                                ' not all, change modifier to show change in currency count
                                amountModifier = TradeYourOffer(x).Value
                            End If
                        End If
                    End If
                Next
            End If
            
            ' Icon of item
            itempic = Item(itemnum).Pic
            
            If itempic > 0 And itempic <= NumItems Then
                If Texture(Tex_Item(itempic)).Width <= 64 Then ' more than 1 frame is handled by anim sub
                    ' Rectangles
                    With rec
                        .top = 0
                        .bottom = 32
                        .Left = 0
                        .Right = 32
                    End With

                    With rec_pos
                        .top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                        .bottom = .top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With
        
                    ' Render
                    RenderSprite Tex_Item(itempic), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec.Right, rec.bottom
                    
                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                        y = rec_pos.top + 22
                        x = rec_pos.Left
                        
                        Amount = GetPlayerInvItemValue(MyIndex, i) - amountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If Amount < 1000000 Then
                            Colour = DX8Color(White)
                        ElseIf Amount >= 1000000 And Amount < 10000000 Then
                            Colour = DX8Color(Yellow)
                        ElseIf Amount >= 10000000 Then
                            Colour = DX8Color(BrightGreen)
                        End If
                        DrawText Format$(ConvertCurrency(Str(Amount)), "#,###,###,###"), x, y, Colour
                    End If
                End If
            End If
        End If
NextLoop:
    Next
    
    ' Rectangle
    With DestRect
        .top = InvTop
        .bottom = InvTop + ((InvOffsetY + 32) * MAX_INV \ InvColumns)
        .Left = InvLeft
        .Right = InvLeft + ((InvOffsetX + 32) * InvColumns)
    End With
        
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picInventory.hWnd, ByVal (0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventory", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawInventoryItem()
    Dim DestRect As D3DRECT
    Dim itemnum As Long, itempic As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InGame Or frmMain.picTempInv.Visible = False Then Exit Sub
    
    ' The item
    itemnum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)

    ' Prevent subscript out range
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then
        frmMain.picTempInv.Cls
        Exit Sub
    End If
    
    ' Icon of item
    itempic = Item(itemnum).Pic
        
    ' Prevent subscript out range
    If itempic = 0 Or itempic > NumItems Then
        frmMain.picTempInv.Cls
        Exit Sub
    End If
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    D3DDevice.BeginScene
        
    ' Render
    RenderSprite Tex_Item(itempic), 2, 2, 0, 0, PIC_X, PIC_Y

    ' Rectangle
    With DestRect
        .Y1 = 2
        .y2 = .Y1 + PIC_Y
        .X1 = 2
        .x2 = .X1 + PIC_X
    End With

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picTempInv.hWnd, ByVal (0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawInventoryItem", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawItemDesc()
    Dim rec As RECT, rec_pos As RECT
    Dim itempic As Long
    Dim itemnum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InGame Or frmMain.picItemDesc.Visible = False Then Exit Sub
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    With rec_pos
        .top = 0
        .bottom = 64
        .Left = 0
        .Right = 64
    End With
    
    ' The item
    itemnum = LastItemDesc

    ' Prevent subscript out range
    If itemnum > 0 And itemnum <= MAX_ITEMS Then
        ' Icon of item
        itempic = Item(itemnum).Pic

        ' Prevent subscript out range
        If itempic = 0 Or itempic > NumItems Then Exit Sub
        
        With rec
            .top = 0
            .bottom = PIC_Y
            .Left = Texture(Tex_Item(itempic)).Width
            .Right = PIC_X
        End With

        ' Render
        RenderTexture Tex_Item(itempic), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec_pos.Right, rec_pos.bottom, rec.Right, rec.bottom
    End If

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present rec_pos, rec_pos, frmMain.picItemDescPic.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawItemDesc", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawPlayerSpells()
    Dim i As Long, x As Long, y As Long, SpellNum As Long, spellicon As Long
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT, DestRect As RECT
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not InGame Or frmMain.picSpells.Visible = False Then Exit Sub

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    For i = 1 To MAX_PLAYER_SPELLS
        ' The spell
        SpellNum = PlayerSpells(i)

        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            ' Icon of spell
            spellicon = Spell(SpellNum).Icon

            If spellicon > 0 And spellicon <= NumSpells Then
                ' Rectangles
                With rec_pos
                    .top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                    .bottom = 0
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                    .Right = 0
                End With
        
                If Not SpellCD(i) = 0 Then
                    Colour = D3DColorARGB(255, 210, 175, 175)
                Else
                    Colour = D3DColorARGB(255, 255, 255, 255)
                End If

                ' Render
                RenderSprite Tex_Spell(spellicon), rec_pos.Left, rec_pos.top, 0, 0, PIC_X, PIC_Y, Colour
            End If
        End If
    Next
    
    ' Rectangle
    With DestRect
        .top = SpellTop
        .bottom = SpellTop + ((SpellOffsetY + 32) * MAX_PLAYER_SPELLS \ SpellColumns)
        .Left = SpellLeft
        .Right = SpellLeft + ((SpellOffsetX + 32) * SpellColumns)
    End With
        
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picSpells.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerSpells", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawSpellDesc()
    Dim rec As RECT, rec_pos As RECT
    Dim spellpic As Long
    Dim SpellNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't rendr
    If Not InGame Or frmMain.picSpellDesc.Visible = False Then Exit Sub
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    With rec_pos
        .top = 0
        .bottom = 64
        .Left = 0
        .Right = 64
    End With
    
    ' The spell
    SpellNum = LastSpellDesc
        
    ' Prevent subscript out range
    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        ' Icon of spell
        spellpic = Spell(SpellNum).Icon

        ' Prevent subscript out range
        If spellpic = 0 Or spellpic > NumSpells Then Exit Sub
        
        With rec
            .top = 0
            .bottom = .top + PIC_Y
            .Left = Texture(Tex_Spell(spellpic)).Width
            .Right = .Left + PIC_X
        End With

        ' Render
        RenderTextureByRects Tex_Spell(spellpic), rec, rec_pos
    End If

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present rec_pos, rec_pos, frmMain.picSpellDescPic.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawSpellDesc", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedSpell()
    Dim rec As RECT, DestRect As D3DRECT
    Dim spellpic As Long
    Dim SpellNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InGame Or frmMain.picTempSpell.Visible = False Then Exit Sub
    
    ' The spell
    SpellNum = PlayerSpells(DragSpell)
        
    ' Prevent subscript out range
    If SpellNum <= 0 And SpellNum > MAX_SPELLS Then
        frmMain.picTempSpell.Cls
        Exit Sub
    End If
    
    ' Icon of spell
    spellpic = Spell(SpellNum).Icon

    ' Prevent subscript out range
    If spellpic = 0 Or spellpic > NumSpells Then
        frmMain.picTempSpell.Cls
        Exit Sub
    End If
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    ' Rectangle
    With DestRect
        .Y1 = 2
        .y2 = PIC_Y
        .X1 = 2
        .x2 = PIC_X
    End With
    
    ' Render
    RenderSprite Tex_Spell(spellpic), 2, 2, 0, 0, PIC_X, PIC_Y

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picTempSpell.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDraggedSpell", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawPlayerTitles()
    Dim i As Long, x As Long, y As Long, TitleNum As Long, titleicon As Long
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT, DestRect As RECT
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Prevent subscript out range
    If Not InGame Or frmMain.picTitles.Visible = False Then Exit Sub

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    For i = 1 To MAX_PLAYER_TITLES
        ' Declaration
        TitleNum = GetPlayerTitle(MyIndex, i)

        ' Prevent subscript out range
        If TitleNum > 0 And TitleNum <= MAX_TITLES Then
            ' Declaration
            titleicon = Title(TitleNum).Icon

            ' Prevent subscript out range
            If titleicon > 0 And titleicon <= NumTitles Then
                ' Rectangles
                With rec_pos
                    .top = TitleTop + ((TitleOffsetY + 32) * ((i - 1) \ TitleColumns))
                    .bottom = .top + PIC_Y
                    .Left = TitleLeft + ((TitleOffsetX + 32) * (((i - 1) Mod TitleColumns)))
                    .Right = .Left + PIC_X
                End With
        
                With rec
                    .top = 0
                    .bottom = 32
                    .Left = 0
                    .Right = 32
                End With

                ' Render
                RenderSprite Tex_Title(titleicon), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec.Right, rec.bottom
            End If
        End If
    Next
    
    ' Rectangle
    With DestRect
        .top = TitleTop
        .bottom = .top + ((TitleOffsetY + 32) * MAX_PLAYER_TITLES \ TitleColumns)
        .Left = TitleLeft
        .Right = .Left + ((TitleOffsetX + 32) * TitleColumns)
    End With
        
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picTitles.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerTitles", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawTitleDesc()
    Dim rec As RECT, rec_pos As RECT
    Dim titlepic As Long
    Dim TitleNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InGame Or frmMain.picTitleDesc.Visible = False Then Exit Sub
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    With rec_pos
        .top = 0
        .bottom = 64
        .Left = 0
        .Right = 64
    End With
    
    ' The title
    TitleNum = LastTitleDesc
        
    ' Prevent subscript out range
    If TitleNum > 0 And TitleNum <= MAX_TITLES Then
        ' Icon of title
        titlepic = Title(TitleNum).Icon

        ' Prevent subscript out range
        If titlepic = 0 Or titlepic > NumTitles Then Exit Sub
        
        With rec
            .top = 0
            .bottom = PIC_Y
            .Left = Texture(Tex_Title(titlepic)).Width
            .Right = .Left + PIC_X
        End With

        ' Render
        RenderTextureByRects Tex_Title(titlepic), rec, rec_pos
    End If

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present rec_pos, rec_pos, frmMain.picTitleDescPic.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTitleDesc", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawDraggedTitle()
    Dim rec As RECT, rec_pos As RECT
    Dim titlepic As Long
    Dim TitleNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InGame Or frmMain.picTempTitle.Visible = False Then Exit Sub
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    ' Rectangle
    With rec_pos
        .top = 2
        .bottom = PIC_Y
        .Left = 2
        .Right = PIC_X
    End With
    
    ' The title
    TitleNum = GetPlayerTitle(MyIndex, DragTitle)
        
    ' Prevent subscript out range
    If TitleNum > 0 And TitleNum <= MAX_TITLES Then
        ' Icon of title
        titlepic = Title(TitleNum).Icon

        ' Prevent subscript out range
        If titlepic = 0 Or titlepic > NumTitles Then Exit Sub
        
        With rec
            .top = 0
            .bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With

        ' Render
        RenderSprite Tex_Title(titlepic), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec_pos.Right, rec_pos.bottom
    End If

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present rec_pos, rec_pos, frmMain.picTempTitle.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawDraggedTitle", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawHotbar()
    Dim sRECT As RECT, dRect As RECT, DestRect As RECT
    Dim i As Long, n As Long
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InGame Or frmMain.picHotbar.Visible = False Then Exit Sub
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    For i = 1 To MAX_HOTBAR
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 And Item(Hotbar(i).Slot).Pic <= NumItems Then
                        ' render
                        RenderSprite Tex_Item(Item(Hotbar(i).Slot).Pic), HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR))), HotbarTop, 0, 0, PIC_X, PIC_Y
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 And Spell(Hotbar(i).Slot).Icon <= NumSpells Then
                        ' Rectangles
                        With sRECT
                            .top = 0
                            .Left = 0
                            .bottom = 32
                            .Right = 32
                        End With
                        
                        ' check for cooldown
                        For n = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(n) = Hotbar(i).Slot Then
                                ' has spell
                                If Not SpellCD(n) = 0 Then
                                    Colour = D3DColorARGB(255, 210, 175, 175)
                                Else
                                    Colour = D3DColorARGB(255, 255, 255, 255)
                                End If
                            End If
                        Next

                        ' Render
                        RenderSprite Tex_Spell(Spell(Hotbar(i).Slot).Icon), HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR))), HotbarTop, 0, 0, sRECT.Right, sRECT.bottom, Colour
                    End If
                End If
            Case 3 ' title
                If Len(Title(Hotbar(i).Slot).Name) > 0 Then
                    If Title(Hotbar(i).Slot).Icon > 0 And Title(Hotbar(i).Slot).Icon <= NumTitles Then
                        ' Render
                        RenderSprite Tex_Title(Title(Hotbar(i).Slot).Icon), HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR))), HotbarTop, 0, 0, PIC_X, PIC_Y
                    End If
                End If
        End Select
        
        ' Render text
        DrawText "F" & i, HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR))) + 4, HotbarTop + 18, DX8Color(White)
    Next
        
    With DestRect
        .top = HotbarTop
        .Left = HotbarLeft
        .bottom = .top + PIC_Y
        .Right = .Left + ((HotbarOffsetX + 32) * (MAX_HOTBAR))
    End With

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picHotbar.hWnd, ByVal (0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawHotbar", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawFace()
    Dim rec As RECT, rec_pos As RECT, faceNum As Long, SrcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InGame Or frmMain.picFace.Visible = False Then Exit Sub
    
    ' The face
    faceNum = GetPlayerSprite(MyIndex)
        
    ' Prevent subscript out range
    If NumFaces = 0 Or faceNum <= 0 Or faceNum > NumFaces Then Exit Sub

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    ' Rectangles
    With rec
        .top = 0
        .bottom = Texture(Tex_Face(faceNum)).Height
        .Left = 0
        .Right = Texture(Tex_Face(faceNum)).Width
    End With

    With rec_pos
        .top = 0
        .bottom = Texture(Tex_Face(faceNum)).Height
        .Left = 0
        .Right = Texture(Tex_Face(faceNum)).Width
    End With

    ' Render
    RenderSprite Tex_Face(faceNum), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec.Right, rec.bottom
    
    ' Rectangle
    With SrcRect
        .X1 = 0
        .x2 = frmMain.picFace.Width
        .Y1 = 0
        .y2 = frmMain.picFace.Height
    End With
    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present SrcRect, SrcRect, frmMain.picFace.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawFace", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawEquipment()
    Dim i As Long, itemnum As Long, itempic As Long
    Dim rec_pos As RECT, DestRect As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InGame Or frmMain.picCharacter.Visible = False Then Exit Sub

    ' Prevent subscript out range
    If NumItems = 0 Then Exit Sub

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    For i = 1 To Equipment.Equipment_Count - 1
        ' The item of equipment
        itemnum = GetPlayerEquipment(MyIndex, i)
            
        ' Prevent subscript out range
        If itemnum > 0 Then
            ' Icon of item
            itempic = Item(itemnum).Pic

            ' Rectangles
            With rec_pos
                .top = EqTop
                .bottom = .top + PIC_Y
                .Left = EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))
                .Right = .Left + PIC_X
            End With

            ' Render
            RenderSprite Tex_Item(itempic), rec_pos.Left, rec_pos.top, 0, 0, PIC_X, PIC_Y
        End If
    Next

    ' Rectangle
    With DestRect
        .top = EqTop
        .bottom = .top + PIC_Y
        .Left = EqLeft
        .Right = .Left + ((EqOffsetX + 32) * (Equipment_Count - 1))
    End With

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picCharacter.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEquipment", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawTrade()
    Dim i As Long, x As Long, y As Long, itemnum As Long, itempic As Long
    Dim Amount As Long
    Dim rec As RECT, rec_pos As RECT, DestRect As RECT
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InGame Or frmMain.picTrade.Visible = False Then Exit Sub

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    For i = 1 To MAX_INV
        ' Draw your own offer
        itemnum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).Num)

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ' Icon of item
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                ' Rectangles
                With rec_pos
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .bottom = .top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With

                With rec
                    .top = 0
                    .bottom = 32
                    .Left = 0
                    .Right = 32
                End With
    
                ' Render
                RenderSprite Tex_Item(itempic), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec.Right, rec.bottom

                ' If item is a stack - draw the amount you have
                If TradeYourOffer(i).Value > 1 Then
                    y = rec_pos.top + 22
                    x = rec_pos.Left
                    
                    Amount = TradeYourOffer(i).Value
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = DX8Color(White)
                    ElseIf Amount > 1000000 And Amount < 10000000 Then
                        Colour = DX8Color(Yellow)
                    ElseIf Amount > 10000000 Then
                        Colour = DX8Color(BrightGreen)
                    End If
                    DrawText ConvertCurrency(Str(Amount)), x, y, Colour
                End If
            End If
        End If
    Next
    
    ' Rectangle
    With DestRect
        .top = InvTop - 24
        .bottom = .top + ((InvOffsetY + 32) * MAX_INV \ InvColumns)
        .Left = InvLeft
        .Right = .Left + ((InvOffsetX + 32) * InvColumns)
    End With
        
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picYourTrade.hWnd, ByVal (0)

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
        
    For i = 1 To MAX_INV
        ' Draw their offer
        itemnum = TradeTheirOffer(i).Num

        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ' Icon of item
            itempic = Item(itemnum).Pic

            If itempic > 0 And itempic <= NumItems Then
                ' Rectangles
                With rec_pos
                    .top = InvTop - 24 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                    .bottom = .top + PIC_Y
                    .Left = InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))
                    .Right = .Left + PIC_X
                End With
        
                With rec
                    .top = 0
                    .bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                ' Render
                RenderSprite Tex_Item(itempic), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec.Right, rec.bottom

                ' If item is a stack - draw the amount you have
                If TradeTheirOffer(i).Value > 1 Then
                    y = rec_pos.top + 22
                    x = rec_pos.Left
                    
                    Amount = TradeTheirOffer(i).Value
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If Amount < 1000000 Then
                        Colour = DX8Color(White)
                    ElseIf Amount >= 1000000 And Amount < 10000000 Then
                        Colour = DX8Color(Yellow)
                    ElseIf Amount >= 10000000 Then
                        Colour = DX8Color(BrightGreen)
                    End If
                    DrawText ConvertCurrency(Str(Amount)), x, y, Colour
                End If
            End If
        End If
    Next
    
    ' Rectangle
    With DestRect
        .top = InvTop - 24
        .bottom = .top + ((InvOffsetY + 32) * MAX_INV \ InvColumns)
        .Left = InvLeft
        .Right = .Left + ((InvOffsetX + 32) * InvColumns)
    End With
        
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picTheirTrade.hWnd, ByVal (0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawTrade", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawShop()
    Dim i As Long, x As Long, y As Long, itemnum As Long, itempic As Long
    Dim Amount As String
    Dim rec As RECT, rec_pos As RECT, DestRect As RECT
    Dim Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InGame Or frmMain.picShop.Visible = False Then Exit Sub
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene

    For i = 1 To MAX_TRADES
        ' The item
        itemnum = Shop(InShop).TradeItem(i).Item 'GetPlayerInvItemNum(MyIndex, i)
        
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ' Icon of item
            itempic = Item(itemnum).Pic
            
            If itempic > 0 And itempic <= NumItems Then
                ' Rectangles
                With rec_pos
                    .top = ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                    .bottom = .top + PIC_Y
                    .Left = ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                    .Right = .Left + PIC_X
                End With
        
                With rec
                    .top = 0
                    .bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                ' render
                RenderSprite Tex_Item(itempic), rec_pos.Left, rec_pos.top, rec.Left, rec.top, rec.Right, rec.bottom

                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    y = rec_pos.top + 22
                    x = rec_pos.Left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = DX8Color(White)
                    ElseIf CLng(Amount) >= 1000000 And CLng(Amount) < 10000000 Then
                        Colour = DX8Color(Yellow)
                    ElseIf CLng(Amount) >= 10000000 Then
                        Colour = DX8Color(Green)
                    End If
                    DrawText ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
    Next
        
    ' Rectangle
    With DestRect
        .top = ShopTop
        .bottom = .top + ((ShopOffsetY + 32) * MAX_TRADES \ ShopColumns)
        .Left = ShopLeft
        .Right = .Left + ((ShopOffsetX + 32) * ShopColumns)
    End With
        
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picShopItems.hWnd, ByVal (0)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawShop", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawBank()
    Dim i As Long, x As Long, y As Long, itemnum As Long
    Dim Amount As String
    Dim sRECT As RECT, dRect As RECT, DestRect As RECT
    Dim Sprite As Long, Colour As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not InGame Or frmMain.picBank.Visible = False Then Exit Sub
                
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
        
    For i = 1 To MAX_BANK
        ' The item
        itemnum = GetBankItemNum(i)
            
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            ' Icon of item
            Sprite = Item(itemnum).Pic
            
            If Sprite > 0 Or Sprite <= NumItems Then
                ' Rectangles
                With dRect
                    .top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                    .bottom = .top + PIC_Y
                    .Left = BankLeft + ((BankOffsetX + 32) * (((i - 1) Mod BankColumns)))
                    .Right = .Left + PIC_X
                End With
        
                With sRECT
                    .top = 0
                    .bottom = PIC_Y
                    .Left = 0
                    .Right = PIC_X
                End With
                
                ' Render
                RenderSprite Tex_Item(Sprite), dRect.Left, dRect.top, sRECT.Left, sRECT.top, sRECT.Right, sRECT.bottom

                ' If item is a stack - draw the amount you have
                If GetBankItemValue(i) > 1 Then
                    y = dRect.top + 22
                    x = dRect.Left - 4
                
                    Amount = CStr(GetBankItemValue(i))
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        Colour = DX8Color(White)
                    ElseIf CLng(Amount) >= 1000000 And CLng(Amount) < 10000000 Then
                        Colour = DX8Color(Yellow)
                    ElseIf CLng(Amount) >= 10000000 Then
                        Colour = DX8Color(BrightGreen)
                    End If
                    DrawText ConvertCurrency(Amount), x, y, Colour
                End If
            End If
        End If
    Next
    
    ' Rectangle
    With DestRect
        .top = BankTop
        .bottom = .top + ((BankOffsetY + 32) * MAX_BANK \ BankColumns)
        .Left = BankLeft
        .Right = .Left + ((BankOffsetX + 32) * BankColumns)
    End With
            
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picBank.hWnd, ByVal (0)
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBank", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawBankItem()
    Dim sRECT As RECT, DestRect As RECT
    Dim itemnum As Long
    Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don' t render
    If Not InGame Or frmMain.picTempBank.Visible = False Then Exit Sub
    
    ' The item
    itemnum = GetBankItemNum(DragBankSlotNum)
    
    ' Prevent subscript out range
    If itemnum <= 0 Or itemnum > MAX_ITEMS Then
        frmMain.picTempBank.Cls
        Exit Sub
    End If
    
    ' Icon of item
    Sprite = Item(GetBankItemNum(DragBankSlotNum)).Pic

    ' Prevent subscript out range
    If Sprite <= 0 Or Sprite > NumItems Then
        frmMain.picTempBank.Cls
        Exit Sub
    End If
    
    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    D3DDevice.BeginScene

    ' Rectangles
    With DestRect
        .top = 2
        .bottom = .top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    With sRECT
        .top = 0
        .bottom = PIC_Y
        .Left = 0
        .Right = PIC_X
    End With

    ' Render
    RenderSprite Tex_Item(Sprite), DestRect.Left, DestRect.top, sRECT.Left, sRECT.top, sRECT.Right, sRECT.bottom

    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMain.picTempBank.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawBankItem", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' frmMenu
Public Sub NewCharacterDrawSprite()
    Dim Sprite As Long
    Dim DestRect As D3DRECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't render
    If Not InMenu Or frmMenu.picCharacter.Visible = False Then Exit Sub

    ' Prevent subscript out range
    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    ' Check sprite by gender
    If newCharSex = SEX_MALE Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite(newCharSprite)
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite(newCharSprite)
    End If
    
    ' Prevent subscript out range
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If

    ' Init rendering
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    D3DDevice.BeginScene
    
    ' Render
    RenderSprite Tex_Character(Sprite), 0, 0, 0, 0, Texture(Tex_Character(Sprite)).Width / 4, Texture(Tex_Character(Sprite)).Height / 4
    
    ' Rectangle
    With DestRect
        .X1 = 0
        .x2 = frmMenu.picSprite.Width
        .Y1 = 0
        .y2 = frmMenu.picSprite.Height
    End With
                    
    ' End rendering
    D3DDevice.EndScene
    D3DDevice.Present DestRect, DestRect, frmMenu.picSprite.hWnd, ByVal (0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "NewCharacterDrawSprite", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' Rendering all textures
Public Sub Render_Graphics()
    Dim x As Long
    Dim y As Long
    Dim i As Long
    Dim rec As RECT
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    
    If D3DDevice Is Nothing Then
        
        i = 1E+15
    End If

    ' don't render
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then InitDirectDraw: Exit Sub
    If frmMain.WindowState = vbMinimized Then Exit Sub
    If GettingMap Then Exit Sub

    ' update the camera
    UpdateCamera

    ' Start rendering
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene

    ' blit panorama
    DrawPanorama

    ' blit lower tiles
    If NumTilesets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapTile(x, y)
                End If
            Next
        Next
    End If
    
    ' render the decals
    For i = 1 To MAX_BYTE
        Call DrawBlood(i)
    Next
            
    ' Blit out the items
    If NumItems > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call DrawItem(i)
            End If
        Next
    End If
            
    ' draw animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If

    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    For y = 0 To Map.MaxY
        If NumCharacters > 0 Then
            ' Players
            For i = 1 To Player_HighIndex
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If Player(i).y = y Then
                        Call DrawPlayer(i)
                    End If
                End If
            Next
                
            ' Npcs
            For i = 1 To Npc_HighIndex
                If MapNpc(i).y = y Then
                    Call DrawNpc(i)
                End If
            Next
        End If
                
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For i = 1 To Resource_Index
                        If MapResource(i).y = y Then
                            Call DrawMapResource(i)
                        End If
                    Next
                End If
            End If
        End If
                
        ' Doors
        If NumDoors > 0 Then
            If Doors_Init Then
                If Door_Index > 0 Then
                    For i = 1 To Door_Index
                        If MapDoor(i).y = y Then
                            Call DrawMapDoor(i)
                        End If
                    Next
                End If
            End If
        End If
    Next

    ' render fog
    DrawFog
    
    ' animations
    If NumAnimations > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                DrawAnimation i, 1
            End If
        Next
    End If

    ' render out upper tiles
    If NumTilesets > 0 Then
        For x = TileView.Left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapFringeTile(x, y)
                End If
            Next
        Next
    End If
    
    ' Render tint
    RenderTexture Tex_White, 0, 0, 0, 0, ScreenX, ScreenY, 32, 32, D3DColorARGB(Map.Alpha, Map.Red, Map.Green, Map.Blue)
    
    ' blit out a square at mouse cursor
    If InMapEditor Then
        If frmEditor_Map.optBlock.Value = True Then
            For x = TileView.Left To TileView.Right
                For y = TileView.top To TileView.bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawDirection(x, y)
                    End If
                Next
            Next
        End If
        Call DrawTileOutline
    End If

    ' Draw the target icon
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(myTarget).x * 32) + Player(myTarget).XOffset, (Player(myTarget).y * 32) + Player(myTarget).YOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).x * 32) + MapNpc(myTarget).XOffset, (MapNpc(myTarget).y * 32) + MapNpc(myTarget).YOffset
        End If
    End If
            
    ' Render the bars
    DrawBars
    
    ' Draw the hover icon
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Player(i).Map = Player(MyIndex).Map Then
                If CurX = Player(i).x And CurY = Player(i).y Then
                    If myTargetType = TARGET_TYPE_PLAYER And myTarget = i Then
                        ' dont render lol
                    Else
                        DrawHover TARGET_TYPE_PLAYER, i, (Player(i).x * 32) + Player(i).XOffset, (Player(i).y * 32) + Player(i).YOffset
                    End If
                End If
            End If
        End If
    Next
            
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            If CurX = MapNpc(i).x And CurY = MapNpc(i).y Then
                If myTargetType = TARGET_TYPE_NPC And myTarget = i Then
                    ' dont render lol
                Else
                    DrawHover TARGET_TYPE_NPC, i, (MapNpc(i).x * 32) + MapNpc(i).XOffset, (MapNpc(i).y * 32) + MapNpc(i).YOffset
                End If
            End If
        End If
    Next

    ' draw FPS
    If BFPS Then
        Call DrawText("FPS: " & GameFPS, 10, 10, DX8Color(Yellow))
        Call DrawText("Ping: " & Ping, 10, 24, DX8Color(Yellow))
        Call DrawText("TCI: " & TickInterval, 10, 38, DX8Color(Yellow))
    End If
    
    ' draw cursor, player X and Y locations
    If BLoc Then
        Call DrawText(Trim$("cur x: " & CurX & " y: " & CurY), Camera.Left, Camera.top + 1, DX8Color(Yellow))
        Call DrawText(Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), Camera.Left, Camera.top + 15, DX8Color(Yellow))
        Call DrawText(Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), Camera.Left, Camera.top + 27, DX8Color(Yellow))
    End If

    ' draw player names
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
            Call DrawPlayerTitle(i)
        End If
    Next
    
    ' draw npc names
    For i = 1 To Npc_HighIndex
        If MapNpc(i).Num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    For i = 1 To Action_HighIndex
        Call DrawActionMsg(i)
    Next i

    ' Blit out map attributes
    If InMapEditor Then
        Call DrawMapAttributes
    End If

    ' Draw map name
    Call DrawText(Map.Name, DrawMapNameX, DrawMapNameY, DrawMapNameColor)
    
    ' End the rendering
    Call D3DDevice.EndScene
    Call D3DDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Render_Graphics", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub Renders()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' don't render
    If D3DDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then Exit Sub
    
    ' New char
    NewCharacterDrawSprite
    
    ' Render in pictures of frmmain
    DrawInventory
    DrawInventoryItem
    DrawItemDesc
    DrawPlayerSpells
    DrawDraggedSpell
    DrawSpellDesc
    DrawPlayerTitles
    DrawDraggedTitle
    DrawTitleDesc
    DrawHotbar
    DrawEquipment
    DrawFace
    DrawTrade
    DrawShop
    DrawBank
    DrawBankItem
    
    ' Render in editors
    EditorMap_DrawTileset
    EditorMap_DrawMapItem
    EditorSpell_DrawIcon
    EditorItem_DrawIcon
    EditorItem_DrawPaperdoll
    EditorTitle_DrawIcon
    EditorAnim_DrawAnim
    EditorNpc_DrawSprite
    EditorDoor_DrawOpeningImage
    EditorDoor_DrawClosedImage
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Renders", "modDirectDraw8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub UpdateCamera()
    Dim offsetX As Long, offsetY As Long, StartX As Long, StartY As Long, EndX As Long, EndY As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Center of screen
    offsetX = Player(MyIndex).XOffset + PIC_X
    offsetY = Player(MyIndex).YOffset + PIC_Y
    
    ' Start screen for rendering
    StartX = GetPlayerX(MyIndex) - ((MAX_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MAX_MAPY + 1) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If Player(MyIndex).XOffset > 0 Then
                offsetX = Player(MyIndex).XOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If Player(MyIndex).YOffset > 0 Then
                offsetY = Player(MyIndex).YOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (MAX_MAPX + 1) + 1
    EndY = StartY + (MAX_MAPY + 1) + 1

    If EndX > Map.MaxX Then
        offsetX = 32

        If EndX = Map.MaxX + 1 Then
            If Player(MyIndex).XOffset < 0 Then
                offsetX = Player(MyIndex).XOffset + PIC_X
            End If
        End If

        EndX = Map.MaxX
        StartX = EndX - MAX_MAPX - 1
    End If

    If EndY > Map.MaxY Then
        offsetY = 32

        If EndY = Map.MaxY + 1 Then
            If Player(MyIndex).YOffset < 0 Then
                offsetY = Player(MyIndex).YOffset + PIC_Y
            End If
        End If

        EndY = Map.MaxY
        StartY = EndY - MAX_MAPY - 1
    End If

    With TileView
        .top = StartY
        .bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .top = offsetY
        .bottom = .top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "UpdateCamera", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function ConvertMapX(ByVal x As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapX = x - (TileView.Left * PIC_X) - Camera.Left
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapX", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function ConvertMapY(ByVal y As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ConvertMapY = y - (TileView.top * PIC_Y) - Camera.top
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "ConvertMapY", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function InViewPort(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Not in view port
    If x < TileView.Left Then Exit Function
    If y < TileView.top Then Exit Function
    If x > TileView.Right Then Exit Function
    If y > TileView.bottom Then Exit Function
    
    ' Return function value
    InViewPort = True
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "InViewPort", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Function IsValidMapPoint(ByVal x As Long, ByVal y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Not is valid map point
    If x < 0 Then Exit Function
    If y < 0 Then Exit Function
    If x > Map.MaxX Then Exit Function
    If y > Map.MaxY Then Exit Function
    
    ' Return function value
    IsValidMapPoint = True
        
    ' Error handler
    Exit Function
errorhandler:
    HandleError "IsValidMapPoint", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function
