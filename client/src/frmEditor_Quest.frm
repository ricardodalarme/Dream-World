VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quest Editor"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8145
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
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   543
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraTask 
      Caption         =   "Task - 1"
      Height          =   3375
      Left            =   3120
      TabIndex        =   43
      Top             =   840
      Width           =   4935
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Quest.frx":0000
         Left            =   120
         List            =   "frmEditor_Quest.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   840
         Width           =   4695
      End
      Begin VB.HScrollBar scrlNum 
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   1440
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlValue 
         Height          =   255
         Left            =   2520
         TabIndex        =   52
         Top             =   1440
         Value           =   1
         Width           =   2295
      End
      Begin VB.TextBox txtMessage 
         Height          =   270
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         Width           =   4695
      End
      Begin VB.TextBox txtMessage 
         Height          =   270
         Index           =   2
         Left            =   120
         TabIndex        =   47
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtMessage 
         Height          =   270
         Index           =   3
         Left            =   2520
         TabIndex        =   46
         Top             =   2640
         Width           =   2295
      End
      Begin VB.HScrollBar scrlTask 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   45
         Top             =   240
         Value           =   1
         Width           =   4695
      End
      Begin VB.CheckBox chkInstant 
         Caption         =   "Finish instantly?"
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   57
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "Num: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "Value: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   54
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Message - At startup:"
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   51
         Top             =   1800
         Width           =   1665
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Message - Not finished:"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   50
         Top             =   2400
         Width           =   1800
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Message - To finish:"
         Height          =   180
         Index           =   5
         Left            =   2520
         TabIndex        =   49
         Top             =   2400
         Width           =   1545
      End
   End
   Begin VB.Frame fraData 
      Caption         =   "Data"
      Height          =   2775
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox chkRetry 
         Caption         =   "Repetitive?"
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox txtDescription 
         Height          =   1215
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   930
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   615
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmbOptions 
         Height          =   300
         ItemData        =   "frmEditor_Quest.frx":0072
         Left            =   120
         List            =   "frmEditor_Quest.frx":0082
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "List"
      Height          =   4575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
      Begin VB.ListBox lstIndex 
         Height          =   4200
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame fraRewards 
      Caption         =   "Rewards"
      Height          =   3255
      Left            =   3120
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlSpriteRew 
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   4695
      End
      Begin VB.ComboBox cmbClassRew 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   2280
         Width           =   4695
      End
      Begin VB.HScrollBar scrlVitalRew 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   34
         Top             =   1680
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlVitalValueRew 
         Height          =   255
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   33
         Top             =   1680
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatValueRew 
         Height          =   255
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   26
         Top             =   1080
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatRew 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   25
         Top             =   1080
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlExpRew 
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   480
         Width           =   2295
      End
      Begin VB.HScrollBar scrlLevelRew 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label lblSpriteRew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   38
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblVitalValueRew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vital value: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   36
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label lblVitalRew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vital slot: Str"
         Height          =   180
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Label lblStatRew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat slot: Str"
         Height          =   180
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblStatValueRew 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat value: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   27
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblExpRew 
         AutoSize        =   -1  'True
         Caption         =   "Exp: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblLevelRew 
         AutoSize        =   -1  'True
         Caption         =   "Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   3255
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlSpriteReq 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   2880
         Width           =   4695
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Left            =   120
         Max             =   5
         Min             =   1
         TabIndex        =   30
         Top             =   1080
         Value           =   1
         Width           =   2295
      End
      Begin VB.HScrollBar scrlStatValueReq 
         Height          =   255
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   29
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2280
         Width           =   4695
      End
      Begin VB.HScrollBar scrlQuestReq 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   4695
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lblSpriteReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         Width           =   465
      End
      Begin VB.Label lblStatValueReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat value: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   32
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stat slot: Str"
         Height          =   180
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label lblBlank 
         AutoSize        =   -1  'True
         Caption         =   "Class:"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblQuestReq 
         AutoSize        =   -1  'True
         Caption         =   "Quest: None"
         Height          =   180
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkInstant_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkInstant.Value = 0 Then
        Quest(EditorIndex).Task(QuestTask).Instant = False
    Else
        Quest(EditorIndex).Task(QuestTask).Instant = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkInstant_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub chkRetry_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkRetry.Value = 0 Then
        Quest(EditorIndex).Retry = False
    Else
        Quest(EditorIndex).Retry = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkRetry_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Quest(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlClassReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbClassRew_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Quest(EditorIndex).ClassRew = cmbClassRew.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlClassRew_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbOptions_Click()
    Dim Index As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Index = cmbOptions.ListIndex
    
    If Index = 0 Then ' Data
        fraData.Visible = True
    Else
        fraData.Visible = False
    End If
    
    If Index = 1 Then ' Requirements
        fraRequirements.Visible = True
    Else
        fraRequirements.Visible = False
    End If
    
    If Index = 2 Then ' Rewards
        fraRewards.Visible = True
    Else
        fraRewards.Visible = False
    End If
    
    If Index = 3 Then ' Task
        fraTask.Visible = True
    Else
        fraTask.Visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbOptions_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    QuestEditorTask
    Quest(EditorIndex).Task(QuestTask).Type = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ClearQuest EditorIndex
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    QuestEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set max values for requeriments
    scrlLevelReq.Max = MAX_LEVELS
    scrlQuestReq.Max = MAX_QUESTS
    scrlStatReq.Max = Stats.Stat_Count - 1
    scrlSpriteReq.Max = NumCharacters
    
    ' set max values for rewards
    scrlLevelRew.Max = MAX_LEVELS
    scrlStatRew.Max = Stats.Stat_Count - 1
    scrlVitalRew.Max = Vitals.Vital_Count - 1
    scrlSpriteRew.Max = NumCharacters
    
    ' set max values for others
    scrlTask.Max = MAX_QUEST_TASKS
    
    ' set values
    cmbOptions.ListIndex = 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    QuestEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlExpRew_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblExpRew.Caption = "Exp: " & scrlExpRew.Value
    Quest(EditorIndex).ExpRew = scrlExpRew.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlExpRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblLevelReq.Caption = "Level: " & scrlLevelReq.Value
    Quest(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelRew_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblLevelRew.Caption = "Level: " & scrlLevelRew.Value
    Quest(EditorIndex).LevelRew = scrlLevelRew.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlNum.Value > 0 Then
        If cmbType.ListIndex = QUEST_TYPE_COLLECTITEMS Then
            If Item(scrlNum.Value).Type = ITEM_TYPE_CURRENCY Then
                scrlValue.Enabled = True
            Else
                scrlValue.Enabled = False
            End If
        End If
    End If
    
    lblNum.Caption = "Num: " & scrlNum.Value
    Quest(EditorIndex).Task(QuestTask).Num = scrlNum.Value
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlQuestReq_Change()
    Dim sString As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If scrlQuestReq.Value = 0 Then sString = "None" Else sString = Trim$(Quest(scrlQuestReq.Value).Name)
    lblQuestReq.Caption = "Quest: " & sString
    Quest(EditorIndex).QuestReq = scrlQuestReq.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlQuestReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpriteReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblSpriteReq.Caption = "Sprite: " & scrlSpriteReq.Value
    Quest(EditorIndex).SpriteReq = scrlSpriteReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpriteReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpriteRew_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblSpriteRew.Caption = "Sprite: " & scrlSpriteRew.Value
    Quest(EditorIndex).SpriteRew = scrlSpriteRew.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpriteRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatReq_Change()
    Dim Index As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Declaration
    Index = scrlStatReq.Value

    ' Set the values
    lblStatReq.Caption = "Stat: " & GetStat(Index)
    scrlStatValueReq.Value = Quest(EditorIndex).StatReq(Index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatRew_Change()
    Dim Index As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Declaration
    Index = scrlStatRew.Value

    ' Set the values
    lblStatRew.Caption = "Stat: " & GetStat(Index)
    scrlStatValueRew.Value = Quest(EditorIndex).StatRew(Index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatValueReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblStatValueReq.Caption = "Stat value: " & scrlStatValueReq.Value
    Quest(EditorIndex).StatReq(scrlStatReq.Value) = scrlStatValueReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatValueReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatValueRew_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblStatValueRew.Caption = "Stat value: " & scrlStatValueRew.Value
    Quest(EditorIndex).StatRew(scrlStatRew.Value) = scrlStatValueRew.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatValueRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlTask_Change()
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' set the label value
    fraTask.Caption = "Task - " & scrlTask.Value
    
    QuestEditorInit

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlTask_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    Quest(EditorIndex).Task(QuestTask).Value = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVitalRew_Change()
    Dim Index As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Declaration
    Index = scrlVitalRew.Value

    ' Set the values
    lblVitalRew.Caption = "Vital: " & GetVital(Index)
    scrlVitalValueRew.Value = Quest(EditorIndex).VitalRew(Index)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVitalRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVitalValueRew_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblVitalValueRew.Caption = "Vital value: " & scrlVitalValueRew.Value
    Quest(EditorIndex).VitalRew(scrlVitalRew.Value) = scrlVitalValueRew.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVitalValueRew_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDescription_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Quest(EditorIndex).Description = txtDescription.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDescription_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtMessage_Change(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Quest(EditorIndex).Task(QuestTask).message(Index) = txtMessage(Index).text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtMessage_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    tmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Quest(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
