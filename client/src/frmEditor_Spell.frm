VERSION 5.00
Begin VB.Form frmEditor_Spell 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spell Editor"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8385
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
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraData 
      Caption         =   "Data"
      Height          =   5055
      Left            =   3360
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtDesc 
         Height          =   855
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   4695
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Spell.frx":0000
         Left            =   120
         List            =   "frmEditor_Spell.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2880
         Width           =   4695
      End
      Begin VB.HScrollBar scrlMP 
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   4695
      End
      Begin VB.HScrollBar scrlCast 
         Height          =   255
         Left            =   120
         Max             =   60
         TabIndex        =   16
         Top             =   4080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlCool 
         Height          =   255
         Left            =   2520
         Max             =   60
         TabIndex        =   15
         Top             =   4080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlIcon 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   4095
      End
      Begin VB.PictureBox picSprite 
         AutoRedraw      =   -1  'True
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
         Height          =   480
         Left            =   4320
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   13
         Top             =   2040
         Width           =   480
      End
      Begin VB.HScrollBar scrlSound 
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4680
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblMP 
         Caption         =   "MP Cost: None"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label lblCast 
         Caption         =   "Casting Time: 0s"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label lblCool 
         AutoSize        =   -1  'True
         Caption         =   "Cooldown Time: 0s"
         Height          =   180
         Left            =   2520
         TabIndex        =   22
         Top             =   3840
         Width           =   1440
      End
      Begin VB.Label lblIcon 
         Caption         =   "Icon: None"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label lblSound 
         Caption         =   "Sound:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   4440
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Salvar"
      Height          =   375
      Left            =   3360
      TabIndex        =   61
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6720
      TabIndex        =   60
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Deletar"
      Height          =   375
      Left            =   5040
      TabIndex        =   59
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame fraEffects 
      Caption         =   "Effects"
      Height          =   4695
      Left            =   3360
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlVital 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   4695
      End
      Begin VB.HScrollBar scrlDuration 
         Height          =   255
         Left            =   120
         Max             =   60
         TabIndex        =   38
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlInterval 
         Height          =   255
         Left            =   2520
         Max             =   60
         TabIndex        =   37
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlRange 
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   4695
      End
      Begin VB.HScrollBar scrlAOE 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2520
         Width           =   4695
      End
      Begin VB.CheckBox chkAOE 
         Caption         =   "Area of Effect spell?"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   3015
      End
      Begin VB.HScrollBar scrlAnimCast 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   3120
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStun 
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3720
         Width           =   4695
      End
      Begin VB.HScrollBar scrlBaseStat 
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   4320
         Width           =   4695
      End
      Begin VB.Label lblVital 
         Caption         =   "Vital: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lblDuration 
         Caption         =   "Duration: 0s"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblInterval 
         Caption         =   "Interval: 0s"
         Height          =   255
         Left            =   2520
         TabIndex        =   46
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblRange 
         Caption         =   "Range: Self-cast"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label lblAOE 
         Caption         =   "AoE: Self-cast"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label lblAnimCast 
         AutoSize        =   -1  'True
         Caption         =   "Cast Anim: None"
         Height          =   180
         Left            =   120
         TabIndex        =   43
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Animation: None"
         Height          =   180
         Left            =   2520
         TabIndex        =   42
         Top             =   2880
         Width           =   1260
      End
      Begin VB.Label lblStun 
         Caption         =   "Stun Duration: None"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   3480
         Width           =   2895
      End
      Begin VB.Label lblBaseStat 
         Caption         =   "Base stat: None"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   4080
         Width           =   2895
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmbOptions 
         Height          =   300
         ItemData        =   "frmEditor_Spell.frx":0045
         Left            =   120
         List            =   "frmEditor_Spell.frx":0055
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Spell List"
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   6000
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   2175
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlLevel 
         Height          =   255
         Left            =   120
         Max             =   100
         TabIndex        =   7
         Top             =   480
         Width           =   4695
      End
      Begin VB.HScrollBar scrlAccess 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   6
         Top             =   1080
         Width           =   4695
      End
      Begin VB.ComboBox cmbClass 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1680
         Width           =   4695
      End
      Begin VB.Label lblLevel 
         Caption         =   "Level: None"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblAccess 
         Caption         =   "Access: None"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Class:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
   End
   Begin VB.Frame fraOthers 
      Caption         =   "Others"
      Height          =   1815
      Left            =   3360
      TabIndex        =   49
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame fraWarp 
         Caption         =   "Warp"
         Height          =   1455
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         Begin VB.HScrollBar scrlMap 
            Height          =   255
            Left            =   120
            Max             =   100
            TabIndex        =   54
            Top             =   480
            Width           =   2175
         End
         Begin VB.HScrollBar scrlX 
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   1080
            Width           =   2175
         End
         Begin VB.HScrollBar scrlY 
            Height          =   255
            Left            =   2400
            TabIndex        =   52
            Top             =   1080
            Width           =   2175
         End
         Begin VB.HScrollBar scrlDir 
            Height          =   255
            Left            =   2400
            TabIndex        =   51
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblMap 
            Caption         =   "Map: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblX 
            Caption         =   "X: 0"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblY 
            Caption         =   "Y: 0"
            Height          =   255
            Left            =   2400
            TabIndex        =   56
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label lblDir 
            Caption         =   "Dir: Down"
            Height          =   255
            Left            =   2400
            TabIndex        =   55
            Top             =   240
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "frmEditor_Spell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAOE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkAOE.Value = 0 Then
        Spell(EditorIndex).IsAoE = False
    Else
        Spell(EditorIndex).IsAoE = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkAOE_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).ClassReq = cmbClass.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClass_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

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
        fraRequirements.Visible = True
    Else
        fraRequirements.Visible = False
    End If
    
    If Index = 2 Then
        fraEffects.Visible = True
    Else
        fraEffects.Visible = False
    End If
    
    If Index = 3 Then
        fraOthers.Visible = True
    Else
        fraOthers.Visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbOptions_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If cmbType.ListIndex = SpellType.sWarp Then
        fraWarp.Visible = True
    Else
        fraWarp.Visible = False
    End If
    
    Spell(EditorIndex).Type = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ClearSpell EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Set max values
    scrlSound.Max = Sound.Sound_Count - 1
    scrlBaseStat.Max = Stats.Stat_Count - 1
    
    ' Set values
    cmbOptions.ListIndex = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccess_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAccess.Value > 0 Then
        lblAccess.Caption = "Access: " & scrlAccess.Value
    Else
        lblAccess.Caption = "Access: None"
    End If
    Spell(EditorIndex).AccessReq = scrlAccess.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccess_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnim.Value > 0 Then
        lblAnim.Caption = "Animation: " & Trim$(Animation(scrlAnim.Value).Name)
    Else
        lblAnim.Caption = "Animation: None"
    End If
    Spell(EditorIndex).SpellAnim = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnimCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAnimCast.Value > 0 Then
        lblAnimCast.Caption = "Cast Anim: " & Trim$(Animation(scrlAnimCast.Value).Name)
    Else
        lblAnimCast.Caption = "Cast Anim: None"
    End If
    Spell(EditorIndex).CastAnim = scrlAnimCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnimCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAOE_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlAOE.Value > 0 Then
        lblAOE.Caption = "AoE: " & scrlAOE.Value & " tiles."
    Else
        lblAOE.Caption = "AoE: Self-cast"
    End If
    Spell(EditorIndex).AoE = scrlAOE.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAOE_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBaseStat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Select Case scrlBaseStat.Value
        Case 0
            lblBaseStat.Caption = "Base stat: None"
        Case 1
            lblBaseStat.Caption = "Base stat: Strength"
        Case 2
            lblBaseStat.Caption = "Base stat: Intelligence"
        Case 3
            lblBaseStat.Caption = "Base stat: Agillity"
        Case 4
            lblBaseStat.Caption = "Base stat: Endurance"
        Case 5
            lblBaseStat.Caption = "Base stat: WillPower"
    End Select

    Spell(EditorIndex).BaseStat = scrlBaseStat.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCast_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCast.Caption = "Casting Time: " & scrlCast.Value & "s"
    Spell(EditorIndex).CastTime = scrlCast.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCast_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlCool_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblCool.Caption = "Cooldown Time: " & scrlCool.Value & "s"
    Spell(EditorIndex).CDTime = scrlCool.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlCool_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    Spell(EditorIndex).Dir = scrlDir.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDir_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDuration_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblDuration.Caption = "Duration: " & scrlDuration.Value & "s"
    Spell(EditorIndex).Duration = scrlDuration.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDuration_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlIcon_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlIcon.Value > 0 Then
        lblIcon.Caption = "Icon: " & scrlIcon.Value
    Else
        lblIcon.Caption = "Icon: None"
    End If
    Spell(EditorIndex).Icon = scrlIcon.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlInterval_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblInterval.Caption = "Interval: " & scrlInterval.Value & "s"
    Spell(EditorIndex).Interval = scrlInterval.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlInterval_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevel_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlLevel.Value > 0 Then
        lblLevel.Caption = "Level: " & scrlLevel.Value
    Else
        lblLevel.Caption = "Level: None"
    End If
    Spell(EditorIndex).LevelReq = scrlLevel.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevel_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMap_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblMap.Caption = "Map: " & scrlMap.Value
    Spell(EditorIndex).Map = scrlMap.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMap_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlMP_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlMP.Value > 0 Then
        lblMP.Caption = "MP Cost: " & scrlMP.Value
    Else
        lblMP.Caption = "MP Cost: None"
    End If
    Spell(EditorIndex).MPCost = scrlMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlMP_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRange_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRange.Value > 0 Then
        lblRange.Caption = "Range: " & scrlRange.Value & " tiles."
    Else
        lblRange.Caption = "Range: Self-cast"
    End If
    Spell(EditorIndex).Range = scrlRange.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRange_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    
    Spell(EditorIndex).Sound = scrlSound.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSound_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStun_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlStun.Value > 0 Then
        lblStun.Caption = "Stun Duration: " & scrlStun.Value & "s"
    Else
        lblStun.Caption = "Stun Duration: None"
    End If
    Spell(EditorIndex).StunDuration = scrlStun.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStun_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVital_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblVital.Caption = "Vital: " & scrlVital.Value
    Spell(EditorIndex).Vital = scrlVital.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVital_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlX_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblX.Caption = "X: " & scrlX.Value
    Spell(EditorIndex).x = scrlX.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlX_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlY_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblY.Caption = "Y: " & scrlY.Value
    Spell(EditorIndex).y = scrlY.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlY_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Spell(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Spell(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Spell(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Spell", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
