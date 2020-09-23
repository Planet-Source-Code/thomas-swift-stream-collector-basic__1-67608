VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form Form1
BorderStyle     =   1 'Fixed Single
Caption         =   "Stream Collector Basic"
ClientHeight    =   3810
ClientLeft      =   45
ClientTop       =   435
ClientWidth     =   7065
LinkTopic       =   "Form1"
LockControls    =   -1 'True
MaxButton       =   0 'False
ScaleHeight     =   3810
ScaleWidth      =   7065
StartUpPosition =   2 'CenterScreen
Begin VB.Timer Timer1
Enabled         =   0 'False
Interval        =   400
Left            =   1215
Top             =   3360
End
Begin TabDlg.SSTab SSTab1
Height          =   3720
Left            =   2655
TabIndex        =   4
Top             =   15
Width           =   4335
_ExtentX        =   7646
_ExtentY        =   6562
_Version        =   393216
Tabs            =   2
TabHeight       =   520
TabCaption(0)   =   "Play"
TabPicture(0)   =   "Form1.frx":0000
Tab(0).ControlEnabled=   -1 'True
Tab(0).Control(0)=   "WindowsMediaPlayer1"
Tab(0).Control(0).Enabled=   0 'False
Tab(0).ControlCount=   1
TabCaption(1)   =   "Edit"
TabPicture(1)   =   "Form1.frx":001C
Tab(1).ControlEnabled=   0 'False
Tab(1).Control(0)=   "Command1"
Tab(1).Control(1)=   "Station"
Tab(1).Control(2)=   "StationURL"
Tab(1).Control(3)=   "Label1"
Tab(1).ControlCount=   4
Begin VB.CommandButton Command1
Caption         =   "Add"
Height          =   300
Left            =   -73282
TabIndex        =   8
Top             =   1230
Width           =   855
End
Begin VB.TextBox Station
Height          =   300
Left            =   -74940
TabIndex        =   7
Top             =   570
Width           =   4215
End
Begin VB.TextBox StationURL
Height          =   300
Left            =   -74940
TabIndex        =   6
Top             =   900
Width           =   4215
End
Begin VB.Label Label1
Alignment       =   2 'Center
BackStyle       =   0 'Transparent
Caption         =   "Label1"
Height          =   255
Left            =   -74955
TabIndex        =   9
Top             =   360
Width           =   4200
End
Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1
Height          =   3345
Left            =   60
TabIndex        =   5
Top             =   330
Width           =   4230
URL             =   ""
rate            =   1
balance         =   0
currentPosition =   0
defaultFrame    =   ""
playCount       =   1
autoStart       =   -1 'True
currentMarker   =   0
invokeURLs      =   -1 'True
baseURL         =   ""
volume          =   50
mute            =   0 'False
uiMode          =   "mini"
stretchToFit    =   -1 'True
windowlessVideo =   -1 'True
enabled         =   -1 'True
enableContextMenu=   -1 'True
fullScreen      =   0 'False
SAMIStyle       =   ""
SAMILang        =   ""
SAMIFilename    =   ""
captioningID    =   ""
enableErrorDialogs=   -1 'True
_cx             =   7461
_cy             =   5900
End
End
Begin VB.ComboBox Combo1
Height          =   315
ItemData        =   "Form1.frx":0038
Left            =   75
List            =   "Form1.frx":0042
Sorted          =   -1 'True
TabIndex        =   3
Text            =   "Combo1"
Top             =   30
Width           =   2550
End
Begin VB.CommandButton Command3
Caption         =   "New"
Height          =   345
Left            =   1800
TabIndex        =   2
Top             =   3375
Width           =   795
End
Begin VB.CommandButton Command2
Caption         =   "Delete"
Height          =   300
Left            =   105
TabIndex        =   1
Top             =   3405
Width           =   930
End
Begin VB.ListBox List1
Height          =   2985
Left            =   75
TabIndex        =   0
Top             =   360
Width           =   2550
End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const BS_FLAT = &H8000&
Private Const GWL_STYLE = (-16)
Private Const WS_CHILD = &H40000000
Private CurSub As String
Private MemSub As String
Private MemStat As String
Private PlaySub As String
Private PlayStat As String
Private Sub Combo1_Change()
    Label1.Caption = Combo1
End Sub
Private Sub Combo1_Click()
    On Error Resume Next
    List1.SetFocus
    Label1.Caption = Combo1
    CurSub = Combo1
    PopulateEntrys
    List1.Selected(0) = True
End Sub
Private Sub Command3_Click()
    Dim X As Integer
    For X = 0 To List1.ListCount - 1
        If List1.Selected(X) = True Then List1.Selected(X) = False
    Next X
    StationURL = ""
    Station = ""
    Command1.Enabled = True
    Command1.Caption = "Add"
    SSTab1.Tab = 1
End Sub
Private Sub Form_Load()
    btnFlat Command1
    btnFlat Command2
    btnFlat Command3
    SSTab1.Tab = 0
    Command1.Enabled = False
    PopulateCombo
    PopulateEntrys
    List1.Selected(0) = True
End Sub
Private Sub Command1_Click()
    Dim X As Integer
    If Command1.Caption = "Add" Then
        For X = 0 To List1.ListCount - 1
            If Station = List1.List(X) And CurSub = Combo1 Then Exit Sub
        Next X
        SetInitEntry Station, "URL", StationURL, App.Path & "\MyINI.ini"
        SetInitEntry Station, "State", Combo1, App.Path & "\MyINI.ini"
        MemSub = Station
        MemStat = Combo1
        PopulateCombo
        PopulateEntrys
        QuickRecall
        Command1.Caption = "Update"
    ElseIf Command1.Caption = "Update" Then
        SetInitEntry Station, "State", Combo1, App.Path & "\MyINI.ini"
        SetInitEntry Station, "URL", StationURL, App.Path & "\MyINI.ini"
        MemSub = Station
        MemStat = Combo1
        PopulateCombo
        PopulateEntrys
        QuickRecall
    End If
End Sub
Private Sub QuickRecall()
    Dim X As Integer
    For X = 0 To Combo1.ListCount - 1
        If Combo1.List(X) = MemStat Then
            Combo1.ListIndex = X
        End If
    Next
    For X = 0 To List1.ListCount - 1
        If List1.List(X) = MemSub Then
            List1.Selected(X) = True
        End If
    Next
End Sub
Private Sub Command2_Click()
    Dim X As Integer
    For X = 0 To List1.ListCount - 1
        If List1.Selected(X) = True Then
            SetInitEntry List1.List(X), vbNullString, vbNullString, App.Path & "\MyINI.ini"
        End If
    Next X
    PopulateCombo
    PopulateEntrys
End Sub
Private Sub PopulateCombo()
    Dim sParts() As String
    Dim i As Integer
    Combo1.Clear
    sParts() = Split(GetInitEntry(vbNullString, vbNullString, vbNullString, App.Path & "\MyINI.ini"), Chr(0))
    For i = 0 To UBound(sParts) - 1
        AddToCombo GetInitEntry(sParts(i), "State", vbNullString, App.Path & "\MyINI.ini")
    Next i
    Combo1.ListIndex = 0
End Sub
Private Sub AddToCombo(sItem As String)
    Dim i As Integer
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = sItem Then Exit Sub
    Next
    Combo1.AddItem sItem
End Sub
Private Sub PopulateEntrys()
    Dim sParts() As String
    Dim i As Integer
    List1.Clear
    sParts() = Split(GetInitEntry(vbNullString, vbNullString, vbNullString, App.Path & "\MyINI.ini"), Chr(0))
    For i = 0 To UBound(sParts) - 1
        If GetInitEntry(sParts(i), "State", vbNullString, App.Path & "\MyINI.ini") = Combo1 Then List1.AddItem sParts(i)
    Next i
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WindowsMediaPlayer1.Controls.stop
    WindowsMediaPlayer1.Close
End Sub
Private Sub List1_Click()
    Dim X As Integer
    For X = 0 To List1.ListCount - 1
        If List1.Selected(X) = True Then
            Station = List1
            StationURL = GetInitEntry(List1.List(X), "URL", "", App.Path & "\MyINI.ini")
            Exit For
        End If
    Next X
    Command1.Enabled = True
    Command1.Caption = "Update"
End Sub
Private Sub List1_dblClick()
    PlaySub = Station
    PlayStat = Combo1
    WindowsMediaPlayer1.URL = GetInitEntry(List1, "URL", "", App.Path & "\MyINI.ini")
    SSTab1.Tab = 0
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error Resume Next
    If SSTab1.Tab = 0 Then QuickPlayRecall
    List1.SetFocus
End Sub
Private Sub SSTab1_DblClick()
    List1.SetFocus
End Sub
Private Sub QuickPlayRecall()
    Dim X As Integer
    For X = 0 To Combo1.ListCount - 1
        If Combo1.List(X) = PlayStat Then
            Combo1.ListIndex = X
        End If
    Next
    For X = 0 To List1.ListCount - 1
        If List1.List(X) = PlaySub Then
            List1.Selected(X) = True
        End If
    Next
End Sub
Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    List1.SetFocus
    If SSTab1.Tab = 0 Then QuickPlayRecall
End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    WindowsMediaPlayer1.URL = GetInitEntry(List1, "URL", "", App.Path & "\MyINI.ini")
End Sub
Public Function btnFlat(Button As CommandButton)
    SetWindowLong Button.hwnd, GWL_STYLE, WS_CHILD Or BS_FLAT
    Button.Visible = True 'Make the button visible (its automaticly hidden when the SetWindowLong call is executed because we reset the button's Attributes)
End Function
