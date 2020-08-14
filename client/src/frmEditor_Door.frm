VERSION 5.00
Begin VB.Form frmEditor_Door 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doors Editor"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8160
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   735
      Left            =   3120
      TabIndex        =   43
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmbOptions 
         Height          =   300
         ItemData        =   "frmEditor_Door.frx":0000
         Left            =   120
         List            =   "frmEditor_Door.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame fraData 
      Caption         =   "Data"
      Height          =   5775
      Left            =   3120
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlSound 
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   5280
         Width           =   4695
      End
      Begin VB.HScrollBar scrlAnimation 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   40
         Top             =   4680
         Width           =   4695
      End
      Begin VB.HScrollBar scrlRespawn 
         Height          =   255
         Left            =   120
         Max             =   6000
         TabIndex        =   38
         Top             =   4080
         Width           =   4695
      End
      Begin VB.HScrollBar scrlOpenWith 
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3480
         Width           =   4695
      End
      Begin VB.HScrollBar scrlClosedImage 
         Height          =   255
         Left            =   2520
         TabIndex        =   34
         Top             =   1080
         Width           =   2295
      End
      Begin VB.PictureBox picClosedImage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   2520
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   33
         Top             =   1440
         Width           =   2280
      End
      Begin VB.PictureBox picOpeningImage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Left            =   120
         ScaleHeight     =   112
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   32
         Top             =   1440
         Width           =   2280
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         MaxLength       =   30
         TabIndex        =   29
         Top             =   480
         Width           =   4695
      End
      Begin VB.HScrollBar scrlOpeningImage 
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblAnimation 
         AutoSize        =   -1  'True
         Caption         =   "Open animation: None"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   4440
         Width           =   1680
      End
      Begin VB.Label lblSound 
         Caption         =   "Sound: None"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label lblRespawn 
         AutoSize        =   -1  'True
         Caption         =   "Respawn Time (Seconds): 0"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   3840
         Width           =   2100
      End
      Begin VB.Label lblOpenWith 
         AutoSize        =   -1  'True
         Caption         =   "Open with: at the move"
         Height          =   180
         Left            =   120
         TabIndex        =   36
         Top             =   3240
         Width           =   1740
      End
      Begin VB.Label lblClosedImage 
         AutoSize        =   -1  'True
         Caption         =   "Closed image: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   35
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblOpeningImage 
         AutoSize        =   -1  'True
         Caption         =   "Opening image: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requeriments"
      Height          =   2055
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   20
         Top             =   480
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   19
         Top             =   1680
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   18
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   17
         Top             =   1680
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   16
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   15
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level: None"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   2520
         TabIndex        =   25
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   480
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   24
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
   End
   Begin VB.Frame fraWarp 
      Caption         =   "Warp"
      Height          =   1455
      Left            =   3120
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         Left            =   150
         Max             =   255
         TabIndex        =   9
         Top             =   480
         Width           =   2295
      End
      Begin VB.HScrollBar scrlX 
         Height          =   255
         Left            =   150
         Max             =   255
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlY 
         Height          =   255
         Left            =   2520
         Max             =   255
         TabIndex        =   7
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlDir 
         Height          =   255
         Left            =   2520
         Max             =   3
         TabIndex        =   6
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblMap 
         AutoSize        =   -1  'True
         Caption         =   "Map: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblX 
         AutoSize        =   -1  'True
         Caption         =   "X: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lblY 
         AutoSize        =   -1  'True
         Caption         =   "Y: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   11
         Top             =   840
         Width           =   315
      End
      Begin VB.Label lblDir 
         AutoSize        =   -1  'True
         Caption         =   "Dir: Up"
         Height          =   180
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame fraLista 
      Caption         =   "Lista"
      Height          =   7095
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstIndex 
         Height          =   6720
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Deletar"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
End
Attribute VB_Name = "frmEditor_Door"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbOptions_Click()
    Dim Index As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Declaration
    Index = cmbOptions.ListIndex
    
    ' open/close windows
    If Index = 0 Then
        fraData.Visible = True
    Else
        fraData.Visible = False
    End If
    
    If Index = 1 Then
        fraWarp.Visible = True
    Else
        fraWarp.Visible = False
    End If
    
    If Index = 2 Then
        fraRequirements.Visible = True
    Else
        fraRequirements.Visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbOptions_Click", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DoorEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearDoor EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Door(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    DoorEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DoorEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set value
    cmbOptions.ListIndex = 0
    
    ' set max values
    scrlOpeningImage.Max = NumDoors
    scrlClosedImage.Max = NumDoors
    scrlAnimation.Max = MAX_ANIMATIONS
    scrlSound.Max = Sound.Sound_Count - 1

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DoorEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimation_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimation.Value > 0 Then
        lblAnimation.Caption = "Open animation: " & scrlAnimation.Value
    Else
        lblAnimation.Caption = "Open animation: None"
    End If
    
    Door(EditorIndex).Animation = scrlAnimation.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimation_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlClosedImage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblClosedImage.Caption = "Closed image: " & scrlClosedImage.Value
    Door(EditorIndex).ClosedImage = scrlClosedImage.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlClosedImage_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDir_Change()
    Dim sDir As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlDir.Value
        Case DIR_UP
            sDir = "Up"
        Case DIR_DOWN
            sDir = "Down"
        Case DIR_RIGHT
            sDir = "Right"
        Case DIR_LEFT
            sDir = "Left"
    End Select
    lblDir.Caption = "Dir: " & sDir
    Door(EditorIndex).Dir = scrlDir.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlLevelReq.Value > 0 Then
        lblLevelReq.Caption = "Level: " & scrlLevelReq.Value
    Else
        lblLevelReq.Caption = "Level: None"
    End If
    
    Door(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblMap.Caption = "Map: " & scrlMap.Value
    Door(EditorIndex).Map = scrlMap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlOpeningImage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblOpeningImage.Caption = "Opening image: " & scrlOpeningImage.Value
    Door(EditorIndex).OpeningImage = scrlOpeningImage.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlOpeningImage_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlOpenWith_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlOpenWith.Value > 0 Then
        lblOpenWith.Caption = "Open with: " & Trim$(Item(scrlOpenWith.Value).Name)
    Else
        lblOpenWith.Caption = "Open with: at the move"
    End If
    
    Door(EditorIndex).OpenWith = scrlOpenWith.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlOpenWith_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRespawn_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblRespawn.Caption = "Respawn Time (Seconds): " & scrlRespawn.Value
    Door(EditorIndex).Respawn = scrlRespawn.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRespawn_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSound_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlSound.Value > 0 Then
        lblSound.Caption = "Sound: " & scrlSound.Value
    Else
        lblSound.Caption = "Sound: None"
    End If
    
    Door(EditorIndex).Sound = scrlSound.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSound_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "Str: "
        Case 2
            text = "End: "
        Case 3
            text = "Int: "
        Case 4
            text = "Agi: "
        Case 5
            text = "Will: "
    End Select
    
    lblStatReq(Index).Caption = text & scrlStatReq(Index).Value
    Door(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblX.Caption = "X: " & scrlX.Value
    Door(EditorIndex).x = scrlX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblY.Caption = "Y: " & scrlY.Value
    Door(EditorIndex).y = scrlY.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    tmpIndex = lstIndex.ListIndex
    Door(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Door(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Door", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
