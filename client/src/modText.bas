Attribute VB_Name = "modText"
Option Explicit

' DirectX8 font
Public D3DFont As D3DXFont

' The font of game
Public GameFont As IFont

' Font variables
Public Const FONT_NAME As String = "Georgia"
Public Const FONT_SIZE As Byte = 8

Public Sub CreateFont()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Create font
    Set GameFont = New StdFont
    With GameFont
        .Size = FONT_SIZE
        .Name = FONT_NAME
    End With

    ' Create the font
    Set D3DFont = D3DX.CreateFont(D3DDevice, GameFont.hFont)
    Set GameFont = Nothing
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CreateFont", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

' text drawing on to back buffer
Public Sub DrawText(text As String, x, y, Color As Long)
    Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out range
    If text = vbNullString Then Exit Sub

    ' Set render position
    With rec
        .top = y
        .Left = x
    End With

    ' Draw
    D3DX.DrawText D3DFont, Color, text, rec, DT_LEFT

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
    Dim TextX As Long, TextY As Long
    Dim Name As String, Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Color = D3DColorARGB(255, 255, 96, 0)
    Else
        Color = DX8Color(BrightRed)
    End If

    ' Player name
    Name = GetPlayerTag(Index) & GetPlayerName(Index)
    
    ' Determine location for text
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset + (PIC_Y \ 2)
        TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - GetWidth((Trim$(Name)))
    Else
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (Texture(Tex_Character(GetPlayerSprite(Index))).Height / 4) + 16
        TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset - (PIC_X \ 2) + (Texture(Tex_Character(GetPlayerSprite(Index))).Width / 4) - GetWidth((Trim$(Name)))
    End If

    ' Draw name
    Call DrawText(Name, TextX, TextY, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawPlayerTitle(ByVal Index As Long)
    Dim TextX As Long, TextY As Long
    Dim Name As String, Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Prevent subscript out range
    If GetPlayerTUsing(Index) = 0 Then Exit Sub
    
    ' Player title
    Name = Trim$(Title(GetPlayerTUsing(Index)).Name)
    Color = DX8Color(Title(GetPlayerTUsing(Index)).Color)
    
    ' Determine location for text
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset + (PIC_Y \ 2) - 16
        TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - GetWidth((Trim$(Name)))
    Else
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (Texture(Tex_Character(GetPlayerSprite(Index))).Height / 4)
        TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset - (PIC_X \ 2) + (Texture(Tex_Character(GetPlayerSprite(Index))).Width / 4) - GetWidth((Trim$(Name)))
    End If
    
    ' Draw name
    Call DrawText(Name, TextX, TextY, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerTitle", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim Color As Long
    Dim Name As String
    Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    npcNum = MapNpc(Index).Num

    Select Case Npc(npcNum).Behaviour
        Case nAttackOnSight
            Color = DX8Color(BrightRed)
        Case nAttackWhenAttacked
            Color = DX8Color(Yellow)
        Case nGuard
            Color = DX8Color(Grey)
        Case nShopKeeper
            Color = DX8Color(Magenta)
        Case nQuest
            Color = DX8Color(Pink)
        Case Else
            Color = DX8Color(BrightGreen)
    End Select

    ' Npc name
    Name = Trim$(Npc(npcNum).Name)
    
    ' Determine location for text
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).YOffset - 16
        TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).XOffset + (PIC_X \ 2) - GetWidth((Trim$(Name)))
    Else
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).YOffset - (Texture(Tex_Character(Npc(npcNum).Sprite)).Height / 4) + 16
        TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).XOffset - (PIC_X \ 2) + (Texture(Tex_Character(Npc(npcNum).Sprite)).Width / 4) - GetWidth((Trim$(Name)))
    End If

    ' Draw name
    Call DrawText(Name, TextX, TextY, Color)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tX As Long
    Dim tY As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Don't render
    If Not frmEditor_Map.optAttribs.Value Then Exit Sub
    
    For x = TileView.Left To TileView.Right
        For y = TileView.top To TileView.bottom
            If IsValidMapPoint(x, y) Then
                With Map.Tile(x, y)
                    tX = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                    tY = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                    Select Case .Type
                        Case TILE_TYPE_BLOCKED
                            DrawText "B", tX, tY, DX8Color(BrightRed)
                        Case TILE_TYPE_WARP
                            DrawText "W", tX, tY, DX8Color(BrightBlue)
                        Case TILE_TYPE_ITEM
                            DrawText "I", tX, tY, DX8Color(White)
                        Case TILE_TYPE_NPCAVOID
                            DrawText "N", tX, tY, DX8Color(White)
                        Case TILE_TYPE_KEYOPEN
                            DrawText "O", tX, tY, DX8Color(White)
                        Case TILE_TYPE_RESOURCE
                            DrawText "O", tX, tY, DX8Color(Green)
                        Case TILE_TYPE_DOOR
                            DrawText "D", tX, tY, DX8Color(Brown)
                        Case TILE_TYPE_NPCSPAWN
                            DrawText "S", tX, tY, DX8Color(Yellow)
                        Case TILE_TYPE_SHOP
                            DrawText "S", tX, tY, DX8Color(BrightBlue)
                        Case TILE_TYPE_BANK
                            DrawText "B", tX, tY, DX8Color(Blue)
                        Case TILE_TYPE_HEAL
                            DrawText "H", tX, tY, DX8Color(BrightGreen)
                        Case TILE_TYPE_TRAP
                            DrawText "T", tX, tY, DX8Color(BrightRed)
                        Case TILE_TYPE_SLIDE
                            DrawText "S", tX, tY, DX8Color(BrightCyan)
                    End Select
                End With
            End If
        Next
    Next

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub DrawActionMsg(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMain.picScreen.Width \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        Call DrawText(ActionMsg(Index).message, x, y, DX8Color(ActionMsg(Index).Color))
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function GetWidth(ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    GetWidth = frmMain.TextWidth(text) \ 2
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal msg As String, ByVal Color As Integer)
    Dim S As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Message + New line?
    If Len(Trim$(frmMain.txtChat.text)) > 0 Then S = vbNewLine & msg Else S = msg
    
    ' Set the message
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(Color)
    frmMain.txtChat.SelText = S

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
