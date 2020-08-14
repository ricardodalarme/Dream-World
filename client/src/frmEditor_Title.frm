VERSION 5.00
Begin VB.Form frmEditor_Title 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Title Editor"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8280
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Salvar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   53
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      TabIndex        =   52
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Deletar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   51
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   8
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox cmbOptions 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmEditor_Title.frx":0000
         Left            =   120
         List            =   "frmEditor_Title.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7080
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraProperties 
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlSound 
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   4680
         Width           =   4575
      End
      Begin VB.HScrollBar scrlRemoveAnimation 
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   5880
         Width           =   4575
      End
      Begin VB.HScrollBar scrlUseAnimation 
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   5280
         Width           =   4575
      End
      Begin VB.CheckBox chkPassive 
         Caption         =   "Passive"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   3600
         Width           =   975
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmEditor_Title.frx":0034
         Left            =   120
         List            =   "frmEditor_Title.frx":003E
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3120
         Width           =   4575
      End
      Begin VB.TextBox txtDescription 
         Height          =   1125
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1080
         Width           =   4575
      End
      Begin VB.ComboBox cmbColor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmEditor_Title.frx":0053
         Left            =   120
         List            =   "frmEditor_Title.frx":0087
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   4080
         Width           =   4575
      End
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   4200
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   7
         Top             =   2280
         Width           =   480
      End
      Begin VB.HScrollBar scrlIcon 
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   3975
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblRemoveAnimation 
         AutoSize        =   -1  'True
         Caption         =   "Animation - Remove title: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   44
         Top             =   5640
         Width           =   2370
      End
      Begin VB.Label lblUseAnimation 
         AutoSize        =   -1  'True
         Caption         =   "Animation - Use title: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   42
         Top             =   5040
         Width           =   2070
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   17
         Top             =   2880
         Width           =   435
      End
      Begin VB.Label lblSound 
         AutoSize        =   -1  'True
         Caption         =   "Sound:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   525
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   840
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         Caption         =   "Color:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   465
      End
      Begin VB.Label lblIcon 
         AutoSize        =   -1  'True
         Caption         =   "Icon: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   24
         Top             =   2280
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   23
         Top             =   1680
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   22
         Top             =   1680
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   20
         Top             =   1080
         Width           =   2175
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   480
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   28
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   26
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   25
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.Frame fraRewards 
      Caption         =   "Rewards"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   3360
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.HScrollBar scrlVitalRew 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   120
         TabIndex        =   47
         Top             =   2280
         Width           =   2175
      End
      Begin VB.HScrollBar scrlVitalRew 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   2520
         TabIndex        =   45
         Top             =   1680
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatRew 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   34
         Top             =   480
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatRew 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   33
         Top             =   480
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatRew 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   32
         Top             =   1080
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatRew 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   2520
         Max             =   255
         TabIndex        =   31
         Top             =   1080
         Width           =   2175
      End
      Begin VB.HScrollBar scrlStatRew 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   30
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblVitalRew 
         AutoSize        =   -1  'True
         Caption         =   "MP: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   2040
         UseMnemonic     =   0   'False
         Width           =   450
      End
      Begin VB.Label lblVitalRew 
         AutoSize        =   -1  'True
         Caption         =   "HP: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   2520
         TabIndex        =   46
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatRew 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   1440
         UseMnemonic     =   0   'False
         Width           =   480
      End
      Begin VB.Label lblStatRew 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatRew 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatRew 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatRew 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   2520
         TabIndex        =   35
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmEditor_Title"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPassive_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If chkPassive.Value = 0 Then
        Title(EditorIndex).Passive = False
    Else
        Title(EditorIndex).Passive = True
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "chkPassive_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbColor_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Title(EditorIndex).Color = cmbColor.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbColor_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbOptions_Click()
    Dim Index As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Index = cmbOptions.ListIndex
    
    If Index = 0 Then ' Properties
        fraProperties.Visible = True
    Else
        fraProperties.Visible = False
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
        
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbOptions_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Title(EditorIndex).Type = cmbType.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TitleEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_TITLES Then Exit Sub
    
    ClearTitle EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Title(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    TitleEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TitleEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' set max values
    scrlIcon.Max = NumTitles
    scrlUseAnimation.Max = MAX_ANIMATIONS
    scrlRemoveAnimation.Max = MAX_ANIMATIONS
    scrlSound.Max = Sound.Sound_Count - 1

    ' set values
    cmbOptions.ListIndex = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    TitleEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    Title(EditorIndex).Icon = scrlIcon.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlIcon_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
        
    Title(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRemoveAnimation_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlRemoveAnimation.Value > 0 Then
        lblRemoveAnimation.Caption = "Animation - Remove title: " & scrlRemoveAnimation.Value
    Else
        lblRemoveAnimation.Caption = "Animation - Remove title: None"
    End If
    Title(EditorIndex).RemoveAnimation = scrlRemoveAnimation.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRemoveAnimation_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    
    Title(EditorIndex).Sound = scrlSound.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSound_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    Title(EditorIndex).StatReq(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatRew_Change(Index As Integer)
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
    
    lblStatRew(Index).Caption = text & scrlStatRew(Index).Value
    Title(EditorIndex).StatRew(Index) = scrlStatRew(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatRew_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlUseAnimation_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If scrlUseAnimation.Value > 0 Then
        lblUseAnimation.Caption = "Animation - Use title: " & scrlUseAnimation.Value
    Else
        lblUseAnimation.Caption = "Animation - Use title: None"
    End If
    Title(EditorIndex).UseAnimation = scrlUseAnimation.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlUseAnimation_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlVitalRew_Change(Index As Integer)
    Dim text As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "HP: "
        Case 2
            text = "MP: "
    End Select
    
    lblVitalRew(Index).Caption = text & scrlVitalRew(Index).Value
    Title(EditorIndex).VitalRew(Index) = scrlVitalRew(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlVitalRew_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDescription_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Title(EditorIndex).Description = txtDescription.text

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDescription_Change", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_TITLES Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Title(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Title(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Title", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
