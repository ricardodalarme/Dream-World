VERSION 5.00
Begin VB.Form frmEditor_Item 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Editor"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
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
   Icon            =   "frmEditor_Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraData 
      Caption         =   "Data"
      Height          =   4575
      Left            =   3360
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.HScrollBar scrlSound 
         Height          =   255
         Left            =   2520
         TabIndex        =   32
         Top             =   4200
         Width           =   2295
      End
      Begin VB.TextBox txtDesc 
         Height          =   975
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1080
         Width           =   4695
      End
      Begin VB.HScrollBar scrlRarity 
         Height          =   255
         Left            =   2520
         Max             =   5
         TabIndex        =   23
         Top             =   3600
         Width           =   2295
      End
      Begin VB.HScrollBar scrlPrice 
         Height          =   255
         LargeChange     =   100
         Left            =   120
         Max             =   30000
         TabIndex        =   22
         Top             =   3600
         Width           =   2295
      End
      Begin VB.HScrollBar scrlAnim 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   21
         Top             =   4200
         Width           =   2295
      End
      Begin VB.ComboBox cmbType 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":3332
         Left            =   120
         List            =   "frmEditor_Item.frx":3354
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3000
         Width           =   4695
      End
      Begin VB.TextBox txtName 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   4695
      End
      Begin VB.HScrollBar scrlPic 
         Height          =   255
         Left            =   120
         Max             =   255
         TabIndex        =   18
         Top             =   2400
         Width           =   4095
      End
      Begin VB.PictureBox picItem 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   17
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Type:"
         Height          =   180
         Left            =   120
         TabIndex        =   41
         Top             =   2760
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblSound 
         Caption         =   "Sound: None"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblRarity 
         AutoSize        =   -1  'True
         Caption         =   "Rarity: 0"
         Height          =   180
         Left            =   2520
         TabIndex        =   28
         Top             =   3360
         Width           =   660
      End
      Begin VB.Label lblPrice 
         AutoSize        =   -1  'True
         Caption         =   "Price: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   600
      End
      Begin VB.Label lblAnim 
         AutoSize        =   -1  'True
         Caption         =   "Anim: None"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   3960
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   180
         Left            =   120
         TabIndex        =   25
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         Caption         =   "Pic: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   450
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   735
      Left            =   3360
      TabIndex        =   33
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cmbOptions 
         Height          =   300
         ItemData        =   "frmEditor_Item.frx":33A1
         Left            =   120
         List            =   "frmEditor_Item.frx":33AE
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   240
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6720
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item List"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   5460
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraOthers 
      Caption         =   "Others"
      Height          =   4335
      Left            =   3360
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Frame fraEquipment 
         Caption         =   "Equipment Data"
         Height          =   3975
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         Begin VB.HScrollBar scrlStatBonus 
            Height          =   255
            Index           =   1
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   72
            Top             =   1800
            Width           =   1335
         End
         Begin VB.ComboBox cmbTool 
            Height          =   300
            ItemData        =   "frmEditor_Item.frx":33CE
            Left            =   120
            List            =   "frmEditor_Item.frx":33DE
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   480
            Width           =   4455
         End
         Begin VB.HScrollBar scrlDamage 
            Height          =   255
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   70
            Top             =   1200
            Width           =   1335
         End
         Begin VB.HScrollBar scrlStatBonus 
            Height          =   255
            Index           =   2
            LargeChange     =   10
            Left            =   1680
            Max             =   255
            TabIndex        =   69
            Top             =   1800
            Width           =   1335
         End
         Begin VB.HScrollBar scrlStatBonus 
            Height          =   255
            Index           =   3
            LargeChange     =   10
            Left            =   3240
            Max             =   255
            TabIndex        =   68
            Top             =   1800
            Width           =   1335
         End
         Begin VB.HScrollBar scrlStatBonus 
            Height          =   255
            Index           =   4
            LargeChange     =   10
            Left            =   120
            Max             =   255
            TabIndex        =   67
            Top             =   2400
            Width           =   1335
         End
         Begin VB.HScrollBar scrlStatBonus 
            Height          =   255
            Index           =   5
            LargeChange     =   10
            Left            =   1680
            Max             =   255
            TabIndex        =   66
            Top             =   2400
            Width           =   1335
         End
         Begin VB.HScrollBar scrlSpeed 
            Height          =   255
            LargeChange     =   100
            Left            =   3240
            Max             =   3000
            Min             =   100
            SmallChange     =   100
            TabIndex        =   65
            Top             =   1200
            Value           =   100
            Width           =   1335
         End
         Begin VB.HScrollBar scrlPaperdoll 
            Height          =   255
            Left            =   3240
            TabIndex        =   64
            Top             =   2400
            Width           =   1335
         End
         Begin VB.PictureBox picPaperdoll 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   120
            ScaleHeight     =   72
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   296
            TabIndex        =   63
            Top             =   2760
            Width           =   4440
         End
         Begin VB.HScrollBar scrlProtection 
            Height          =   255
            LargeChange     =   10
            Left            =   1680
            Max             =   255
            TabIndex        =   62
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Str: 0"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   82
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   585
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Object Tool:"
            Height          =   180
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblDamage 
            AutoSize        =   -1  'True
            Caption         =   "Damage: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   80
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   825
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ End: 0"
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   79
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   645
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Int: 0"
            Height          =   180
            Index           =   3
            Left            =   3240
            TabIndex        =   78
            Top             =   1560
            UseMnemonic     =   0   'False
            Width           =   585
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Agi: 0"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   77
            Top             =   2160
            UseMnemonic     =   0   'False
            Width           =   615
         End
         Begin VB.Label lblStatBonus 
            AutoSize        =   -1  'True
            Caption         =   "+ Will: 0"
            Height          =   180
            Index           =   5
            Left            =   1680
            TabIndex        =   76
            Top             =   2160
            UseMnemonic     =   0   'False
            Width           =   630
         End
         Begin VB.Label lblSpeed 
            AutoSize        =   -1  'True
            Caption         =   "Speed: 0.1 sec"
            Height          =   180
            Left            =   3240
            TabIndex        =   75
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   1140
         End
         Begin VB.Label lblPaperdoll 
            AutoSize        =   -1  'True
            Caption         =   "Paperdoll: 0"
            Height          =   180
            Left            =   3240
            TabIndex        =   74
            Top             =   2160
            Width           =   915
         End
         Begin VB.Label lblProtection 
            AutoSize        =   -1  'True
            Caption         =   "Protection: 0"
            Height          =   180
            Left            =   1680
            TabIndex        =   73
            Top             =   960
            UseMnemonic     =   0   'False
            Width           =   990
         End
      End
      Begin VB.Frame fraBag 
         Caption         =   "Bag - 1"
         Height          =   1455
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         Begin VB.HScrollBar scrlBag 
            Height          =   255
            Left            =   120
            Max             =   5
            Min             =   1
            TabIndex        =   57
            Top             =   240
            Value           =   1
            Width           =   4455
         End
         Begin VB.HScrollBar scrlNum 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   56
            Top             =   1080
            Width           =   2175
         End
         Begin VB.HScrollBar scrlValue 
            Height          =   255
            Left            =   2400
            Max             =   255
            TabIndex        =   55
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label lblNum 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   60
            Top             =   840
            Width           =   555
         End
         Begin VB.Label lblItemName 
            AutoSize        =   -1  'True
            Caption         =   "Item: None"
            Height          =   180
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            Caption         =   "Value: 0"
            Height          =   180
            Left            =   2400
            TabIndex        =   58
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   645
         End
      End
      Begin VB.Frame fraSpell 
         Caption         =   "Spell Data"
         Height          =   1215
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         Begin VB.HScrollBar scrlSpell 
            Height          =   255
            Left            =   1080
            Max             =   255
            TabIndex        =   51
            Top             =   720
            Value           =   1
            Width           =   3495
         End
         Begin VB.Label lblSpell 
            AutoSize        =   -1  'True
            Caption         =   "Num: 0"
            Height          =   180
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   555
         End
         Begin VB.Label lblSpellName 
            AutoSize        =   -1  'True
            Caption         =   "Name: None"
            Height          =   180
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame fraVitals 
         Caption         =   "Consume Data"
         Height          =   2175
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   4695
         Begin VB.HScrollBar scrlAddHp 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   46
            Top             =   480
            Width           =   4455
         End
         Begin VB.HScrollBar scrlAddMP 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   45
            Top             =   1080
            Width           =   4455
         End
         Begin VB.HScrollBar scrlAddExp 
            Height          =   255
            Left            =   120
            Max             =   255
            TabIndex        =   44
            Top             =   1680
            Width           =   4455
         End
         Begin VB.Label lblAddHP 
            AutoSize        =   -1  'True
            Caption         =   "Add HP: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   240
            UseMnemonic     =   0   'False
            Width           =   780
         End
         Begin VB.Label lblAddMP 
            AutoSize        =   -1  'True
            Caption         =   "Add MP: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   48
            Top             =   840
            UseMnemonic     =   0   'False
            Width           =   795
         End
         Begin VB.Label lblAddExp 
            AutoSize        =   -1  'True
            Caption         =   "Add Exp: 0"
            Height          =   180
            Left            =   120
            TabIndex        =   47
            Top             =   1440
            UseMnemonic     =   0   'False
            Width           =   840
         End
      End
   End
   Begin VB.Frame fraRequirements 
      Caption         =   "Requirements"
      Height          =   3255
      Left            =   3360
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox cmbClassReq 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1680
         Width           =   4695
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   120
         Max             =   5
         TabIndex        =   36
         Top             =   2280
         Width           =   4695
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   99
         TabIndex        =   35
         Top             =   2880
         Width           =   4695
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   10
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   9
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   3240
         Max             =   255
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   1680
         Max             =   255
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class:"
         Height          =   180
         Left            =   120
         TabIndex        =   40
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         Caption         =   "Access: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         Caption         =   "Level: 0"
         Height          =   180
         Left            =   120
         TabIndex        =   38
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Str: 0"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "End: 0"
         Height          =   180
         Index           =   2
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Int: 0"
         Height          =   180
         Index           =   3
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         UseMnemonic     =   0   'False
         Width           =   435
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Agi: 0"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         Caption         =   "Will: 0"
         Height          =   180
         Index           =   5
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmEditor_Item"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private LastIndex As Long

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).ClassReq = cmbClassReq.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbClassReq_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
        fraOthers.Visible = True
    Else
        fraOthers.Visible = False
    End If
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbOptions_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbTool_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    Item(EditorIndex).Data3 = cmbTool.ListIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbTool_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmbType_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ItemEditorType cmbType.ListIndex
    Item(EditorIndex).Type = cmbType.ListIndex

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmbType_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorCancel
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdCancel_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdDelete_Click()
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    ClearItem EditorIndex
    
    tmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdDelete_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ItemEditorOk
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "cmdSave_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Set max values
    scrlPic.Max = NumItems
    scrlAnim.Max = MAX_ANIMATIONS
    scrlPaperdoll.Max = NumPaperdolls
    scrlSound.Max = Sound.Sound_Count - 1
    scrlBag.Max = MAX_BAG
    scrlNum.Max = MAX_ITEMS
    
    ' Set values
    cmbOptions.ListIndex = 0

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "Form_Load", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub lstIndex_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ItemEditorInit
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "lstIndex_Click", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAccessReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblAccessReq.Caption = "Access: " & scrlAccessReq.Value
    Item(EditorIndex).AccessReq = scrlAccessReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddExp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddExp.Caption = "Add Exp: " & scrlAddExp.Value
    Item(EditorIndex).AddEXP = scrlAddExp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddExp_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddHp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddHP.Caption = "Add HP: " & scrlAddHp.Value
    Item(EditorIndex).AddHP = scrlAddHp.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddHP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAddMp_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblAddMP.Caption = "Add MP: " & scrlAddMP.Value
    Item(EditorIndex).AddMP = scrlAddMP.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAddMP_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlAnim_Change()
    Dim sString As String
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    If scrlAnim.Value = 0 Then
        sString = "None"
    Else
        sString = Trim$(Animation(scrlAnim.Value).Name)
    End If
    lblAnim.Caption = "Anim: " & sString
    Item(EditorIndex).Animation = scrlAnim.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlAnim_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlBag_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    scrlNum.Value = Item(EditorIndex).BagItem(scrlBag.Value)
    scrlValue.Value = Item(EditorIndex).BagValue(scrlBag.Value)

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlBag_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlDamage_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblDamage.Caption = "Damage: " & scrlDamage.Value
    Item(EditorIndex).Damage = scrlDamage.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlDamage_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlLevelReq_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblLevelReq.Caption = "Level: " & scrlLevelReq
    Item(EditorIndex).LevelReq = scrlLevelReq.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlNum_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblNum.Caption = "Num: " & scrlNum.Value

    If scrlNum.Value > 0 Then
        lblItemName.Caption = "Item: " & Trim$(Item(scrlNum.Value).Name)
    End If
    
    Item(EditorIndex).BagItem(scrlBag.Value) = scrlNum.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlNum_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPaperdoll_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPaperdoll.Caption = "Paperdoll: " & scrlPaperdoll.Value
    Item(EditorIndex).Paperdoll = scrlPaperdoll.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPaperdoll_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPic_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPic.Caption = "Pic: " & scrlPic.Value
    Item(EditorIndex).Pic = scrlPic.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPic_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlPrice_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblPrice.Caption = "Price: " & scrlPrice.Value
    Item(EditorIndex).Price = scrlPrice.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlPrice_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlProtection_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblProtection.Caption = "Protection: " & scrlProtection.Value
    Item(EditorIndex).Protection = scrlProtection.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlProtection_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlRarity_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblRarity.Caption = "Rarity: " & scrlRarity.Value
    Item(EditorIndex).Rarity = scrlRarity.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlRarity_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    
    Item(EditorIndex).Sound = scrlSound.Value

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSound_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpeed_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    lblSpeed.Caption = "Speed: " & scrlSpeed.Value / 1000 & " sec"
    Item(EditorIndex).Speed = scrlSpeed.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpeed_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlSpell_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    
    If scrlSpell.Value > 0 Then
        lblSpellName.Caption = "Name: " & Trim$(Spell(scrlSpell.Value).Name)
    Else
        lblSpellName.Caption = "Name: None"
    End If
    
    lblSpell.Caption = "Spell: " & scrlSpell.Value
    
    Item(EditorIndex).Data1 = scrlSpell.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlSpell_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlStatBonus_Change(Index As Integer)
    Dim text As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Select Case Index
        Case 1
            text = "+ Str: "
        Case 2
            text = "+ End: "
        Case 3
            text = "+ Int: "
        Case 4
            text = "+ Agi: "
        Case 5
            text = "+ Will: "
    End Select
            
    lblStatBonus(Index).Caption = text & scrlStatBonus(Index).Value
    Item(EditorIndex).Add_Stat(Index) = scrlStatBonus(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatBonus_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
    Item(EditorIndex).Stat_Req(Index) = scrlStatReq(Index).Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlStatReq_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub scrlValue_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lblValue.Caption = "Value: " & scrlValue.Value
    Item(EditorIndex).BagValue(scrlBag.Value) = scrlValue.Value
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "scrlValue_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub

    Item(EditorIndex).Desc = txtDesc.text
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtDesc_Change", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
    Dim tmpIndex As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If EditorIndex = 0 Or EditorIndex > MAX_ITEMS Then Exit Sub
    tmpIndex = lstIndex.ListIndex
    Item(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Item(EditorIndex).Name, EditorIndex - 1
    lstIndex.ListIndex = tmpIndex
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "txtName_Validate", "frmEditor_Item", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
