VERSION 5.00
Begin VB.Form frmSearch 
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox inDoEvents 
      Caption         =   "Do Events"
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   1800
      Width           =   1455
   End
   Begin VB.ComboBox inCompare 
      Height          =   315
      ItemData        =   "frmSearch.frx":0000
      Left            =   7440
      List            =   "frmSearch.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox inContaining 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Text            =   "Containing"
      Top             =   1320
      Width           =   6255
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmSearch.frx":004D
      Left            =   7440
      List            =   "frmSearch.frx":0063
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   720
      Width           =   2055
   End
   Begin VB.ComboBox inFOF 
      Height          =   315
      ItemData        =   "frmSearch.frx":0097
      Left            =   7440
      List            =   "frmSearch.frx":00A4
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   720
      Width           =   6300
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   1095
      Left            =   8040
      TabIndex        =   5
      Top             =   2160
      Width           =   1500
   End
   Begin VB.ListBox lstResults 
      Height          =   2010
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   7860
   End
   Begin VB.TextBox txtBase 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   6285
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Compare method:"
      Height          =   195
      Left            =   7440
      TabIndex        =   15
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Containing:"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Subfolders Depth:"
      Height          =   195
      Left            =   7440
      TabIndex        =   10
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label outStatus1 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   3840
      Width           =   9615
   End
   Begin VB.Label Label3 
      Caption         =   "Search Word:"
      Height          =   270
      Left            =   75
      TabIndex        =   6
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "("""" to search all drives)"
      Height          =   195
      Left            =   1080
      TabIndex        =   3
      Top             =   410
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "Base Directory:"
      Height          =   255
      Left            =   30
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents Search As uSc_FileSearch
Attribute Search.VB_VarHelpID = -1

Private Sub cmdCancel_Click()
Search.SrcCancel
End Sub

Private Sub cmdSearch_Click()
    lstResults.Clear
    Search.Src txtBase.Text, txtSearch.Text, inContaining.Text, Val(Left(inFOF.Text, 1))
End Sub



Private Sub Combo1_Change()
If IsNumeric(Combo1.Text) Then Search.Depth = Val(Combo1.Text)
End Sub

Private Sub Combo1_Click()
    If IsNumeric(Combo1.Text) Then
        Search.Depth = Val(Combo1.Text)
    ElseIf IsNumeric(Left(Combo1.Text, InStr(1, Combo1.Text, " ", vbTextCompare))) Then
        Search.Depth = Val(Left(Combo1.Text, InStr(1, Combo1.Text, " ", vbTextCompare)))
    Else
        Combo1.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
Set Search = New uSc_FileSearch
txtBase.Text = "C:\Program Files"
txtSearch.Text = "*.txt"
inContaining.Text = "millions"
inFOF.ListIndex = 0
inCompare.ListIndex = 1
Combo1.ListIndex = 0
inDoEvents.Value = 0
End Sub

Private Sub inCompare_Click()
Search.fCompare = Val(Left(inCompare.Text, 1))
End Sub

Private Sub inDoEvents_Click()
Search.EnableDoEvents = inDoEvents.Value
End Sub

Private Sub Search_CurrentFolder(Path As String)
outStatus1.Caption = Path
End Sub

Private Sub Search_FileFound(Path As String)
    lstResults.AddItem ("FILE:" & Path)
End Sub

Private Sub Search_FolderFound(Path As String)
    lstResults.AddItem ("FOLD:" & Path)
End Sub

Private Sub Search_SearchComplete(NoFiles As Double, NoFolders As Double, Files As String, Folders As String, Canceled As Boolean)
outStatus1.Caption = "Search Complete | " & NoFiles & "Files Found and " & NoFolders & "Folders found | Canceled: " & Canceled
End Sub

