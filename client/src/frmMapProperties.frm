VERSION 5.00
Begin VB.Form frmEditor_MapProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Properties"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   376
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame7 
      Caption         =   "Fog"
      Height          =   855
      Left            =   1800
      TabIndex        =   47
      Top             =   1680
      Width           =   3735
      Begin VB.HScrollBar scrlFogSpeed 
         Height          =   255
         Left            =   1320
         Max             =   255
         TabIndex        =   50
         Top             =   480
         Width           =   1095
      End
      Begin VB.HScrollBar scrlFog 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   49
         Top             =   480
         Width           =   1095
      End
      Begin VB.HScrollBar scrlFogOpacity 
         Height          =   255
         Left            =   2520
         Max             =   255
         TabIndex        =   48
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblFogSpeed 
         AutoSize        =   -1  'True
         Caption         =   "Speed: 0"
         Height          =   195
         Left            =   1320
         TabIndex        =   53
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lblFog 
         AutoSize        =   -1  'True
         Caption         =   "Fog: None"
         Height          =   195
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   870
      End
      Begin VB.Label lblFogOpacity 
         AutoSize        =   -1  'True
         Caption         =   "Opacity: 0"
         Height          =   195
         Left            =   2520
         TabIndex        =   51
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Map Overlay"
      Height          =   855
      Left            =   120
      TabIndex        =   38
      Top             =   2640
      Width           =   5415
      Begin VB.HScrollBar scrlBlue 
         Height          =   255
         Left            =   2760
         Max             =   255
         TabIndex        =   42
         Top             =   480
         Width           =   1215
      End
      Begin VB.HScrollBar scrlGreen 
         Height          =   255
         Left            =   1440
         Max             =   255
         TabIndex        =   41
         Top             =   480
         Width           =   1215
      End
      Begin VB.HScrollBar scrlRed 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   40
         Top             =   480
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAlpha 
         Height          =   255
         Left            =   4080
         Max             =   255
         TabIndex        =   39
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblBlue 
         AutoSize        =   -1  'True
         Caption         =   "Blue: 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   46
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblGreen 
         AutoSize        =   -1  'True
         Caption         =   "Green: 0"
         Height          =   195
         Left            =   1440
         TabIndex        =   45
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblRed 
         AutoSize        =   -1  'True
         Caption         =   "Red: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   570
      End
      Begin VB.Label lblAlpha 
         AutoSize        =   -1  'True
         Caption         =   "Opacity: 0"
         Height          =   195
         Left            =   4080
         TabIndex        =   43
         Top             =   240
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame fraMusic 
      Caption         =   "Music - None"
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   1815
      Begin VB.HScrollBar scrlMusic 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame frmMaxSizes 
      Caption         =   "Max Sizes"
      Height          =   975
      Left            =   120
      TabIndex        =   19
      Top             =   3600
      Width           =   1815
      Begin VB.TextBox txtMaxX 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtMaxY 
         Height          =   285
         Left            =   960
         TabIndex        =   20
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Max X:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Max Y:"
         Height          =   195
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Map Links"
      Height          =   1815
      Left            =   3960
      TabIndex        =   14
      Top             =   3600
      Width           =   1575
      Begin VB.TextBox txtDownRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   34
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtUpRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   960
         TabIndex        =   33
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtDownLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   240
         TabIndex        =   32
         Text            =   "0"
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox txtUpLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Text            =   "0"
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtUp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         TabIndex        =   18
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtDown 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   600
         TabIndex        =   17
         Text            =   "0"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtRight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtLeft 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Text            =   "0"
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current map: 0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   1320
      End
   End
   Begin VB.Frame fraMapSettings 
      Caption         =   "Map Settings"
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5415
      Begin VB.HScrollBar scrlPanorama 
         Height          =   255
         Left            =   2760
         TabIndex        =   36
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   5175
      End
      Begin VB.ComboBox cmbMoral 
         Height          =   315
         ItemData        =   "frmMapProperties.frx":0000
         Left            =   120
         List            =   "frmMapProperties.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label lblPanorama 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Panorama: 0"
         Height          =   195
         Left            =   2760
         TabIndex        =   35
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Moral:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Boot Settings"
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
      Begin VB.TextBox txtBootMap 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtBootX 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   600
         TabIndex        =   6
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtBootY 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "0"
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Map:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   420
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Width           =   195
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   195
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   180
      End
   End
   Begin VB.Frame fraNPCs 
      Caption         =   "NPCs"
      Height          =   2415
      Left            =   2040
      TabIndex        =   2
      Top             =   3600
      Width           =   1815
      Begin VB.ListBox lstNpcs 
         Height          =   1620
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox cmbNpc 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1920
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   5520
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_MapProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbNpc_Click()
    Dim tmpString() As String
    Dim npcNum As Long
    Dim x As Long, tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(cmbNpc.List(cmbNpc.ListIndex))
    ' make sure it's not a clear
    If Not cmbNpc.List(cmbNpc.ListIndex) = "No NPC" Then
        npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
        Map.Npc(lstNpcs.ListIndex + 1) = npcNum
    Else
        Map.Npc(lstNpcs.ListIndex + 1) = 0
    End If
    
    ' re-load the list
    tmpIndex = lstNpcs.ListIndex
    lstNpcs.Clear
    For x = 1 To MAX_MAP_NPCS
        If Map.Npc(x) > 0 Then
            lstNpcs.AddItem x & ": " & Trim$(Npc(Map.Npc(x)).Name)
        Else
            lstNpcs.AddItem x & ": No NPC"
        End If
    Next
    lstNpcs.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbNpc_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOk_Click()
    Dim i As Long
    Dim sTemp As Long
    Dim x As Long, x2 As Long
    Dim y As Long, y2 As Long
    Dim tempArr() As TileRec

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If Not IsNumeric(txtMaxX.text) Then txtMaxX.text = Map.MaxX
    If Val(txtMaxX.text) < MAX_MAPX Then txtMaxX.text = MAX_MAPX
    If Val(txtMaxX.text) > MAX_BYTE Then txtMaxX.text = MAX_BYTE
    If Not IsNumeric(txtMaxY.text) Then txtMaxY.text = Map.MaxY
    If Val(txtMaxY.text) < MAX_MAPY Then txtMaxY.text = MAX_MAPY
    If Val(txtMaxY.text) > MAX_BYTE Then txtMaxY.text = MAX_BYTE

    With Map
        .Name = Trim$(txtName.text)
        .Music = scrlMusic.Value
        .Up = Val(txtUp.text)
        .Down = Val(txtDown.text)
        .Left = Val(txtLeft.text)
        .Right = Val(txtRight.text)
        .UpLeft = Val(txtUpLeft.text)
        .UpRight = Val(txtUpRight.text)
        .DownLeft = Val(txtDownLeft.text)
        .DownRight = Val(txtDownRight.text)
        .Moral = cmbMoral.ListIndex
        .Panorama = scrlPanorama.Value
        .BootMap = Val(txtBootMap.text)
        .BootX = Val(txtBootX.text)
        .BootY = Val(txtBootY.text)
        .Red = scrlRed.Value
        .Green = scrlGreen.Value
        .Blue = scrlBlue.Value
        .Alpha = scrlAlpha.Value
        .Fog = scrlFog.Value
        .FogSpeed = scrlFogSpeed.Value
        .FogOpacity = scrlFogOpacity.Value

        ' set the data before changing it
        tempArr = Map.Tile
        x2 = Map.MaxX
        y2 = Map.MaxY
        ' change the data
        .MaxX = Val(txtMaxX.text)
        .MaxY = Val(txtMaxY.text)
        ReDim Map.Tile(0 To .MaxX, 0 To .MaxY)

        If x2 > .MaxX Then x2 = .MaxX
        If y2 > .MaxY Then y2 = .MaxY

        For x = 0 To x2
            For y = 0 To y2
                .Tile(x, y) = tempArr(x, y)
            Next
        Next
    End With

    Call UpdateDrawMapName
    Unload frmEditor_MapProperties
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdOk_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdPlay_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Stop_Music
    Play_Music scrlMusic.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdPlay_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdStop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Stop_Music
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdStop_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Max values
    scrlFog.Max = NumFogs
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstNpcs_Click()
    Dim tmpString() As String
    Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' exit out if needed
    If Not cmbNpc.ListCount > 0 Then Exit Sub
    If Not lstNpcs.ListCount > 0 Then Exit Sub
    
    ' set the combo box properly
    tmpString = Split(lstNpcs.List(lstNpcs.ListIndex))
    npcNum = CLng(Left$(tmpString(0), Len(tmpString(0)) - 1))
    cmbNpc.ListIndex = Map.Npc(npcNum)
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstNpcs_Click", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAlpha_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAlpha.Caption = "Opacity: " & scrlAlpha.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAlpha_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBlue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblBlue.Caption = "Blue: " & scrlBlue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBlue_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFog_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlFog.Value = 0 Then
        lblFog.Caption = "None."
    Else
        lblFog.Caption = "Fog: " & scrlFog.Value
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFog_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFogOpacity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFogOpacity.Caption = "Opacity: " & scrlFogOpacity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFogOpacity_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlFogSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblFogSpeed.Caption = "Speed: " & scrlFogSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "ScrlFogSpeed_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlGreen_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblGreen.Caption = "Green: " & scrlGreen.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlGreen_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMusic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlMusic.Value > 0 Then fraMusic.Caption = "Music: " & scrlMusic.Value Else fraMusic.Caption = "Music: None"
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMusic_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPanorama_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblPanorama.Caption = "Panorama: " & scrlPanorama.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPanorama_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRed.Caption = "Red: " & scrlRed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRed_Change", "frmEditor_MapProperties", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
