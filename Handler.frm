VERSION 4.00
Begin VB.Form Handler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Window Handler"
   ClientHeight    =   4725
   ClientLeft      =   1125
   ClientTop       =   1665
   ClientWidth     =   6615
   Height          =   5190
   Icon            =   "Handler.frx":0000
   Left            =   1065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6615
   Top             =   1260
   Width           =   6735
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3600
      TabIndex        =   24
      Text            =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1920
      TabIndex        =   23
      Text            =   "0"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Set window text"
      Height          =   300
      Left            =   4440
      TabIndex        =   21
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1200
      TabIndex        =   20
      Top             =   4320
      Width           =   3135
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Set height\width\top\left"
      Height          =   300
      Left            =   4440
      TabIndex        =   18
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Text            =   "0"
      Top             =   3630
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Text            =   "0"
      Top             =   3630
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Refresh"
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Not Ontop"
      Height          =   300
      Left            =   5520
      TabIndex        =   12
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Ontop"
      Height          =   300
      Left            =   4440
      TabIndex        =   11
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Disable"
      Height          =   300
      Left            =   3360
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Enable"
      Height          =   300
      Left            =   2280
      TabIndex        =   9
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Focus"
      Height          =   300
      Left            =   1200
      TabIndex        =   8
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Flash"
      Height          =   300
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Close"
      Height          =   300
      Left            =   5520
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Restore"
      Height          =   300
      Left            =   4440
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Maximize"
      Height          =   300
      Left            =   3360
      TabIndex        =   4
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Minimize"
      Height          =   300
      Left            =   2280
      TabIndex        =   3
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show"
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Left:"
      Height          =   195
      Left            =   2880
      TabIndex        =   25
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Top:"
      Height          =   195
      Left            =   1200
      TabIndex        =   22
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Window text:"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   4350
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      Top             =   3675
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   3675
      Width           =   510
   End
End
Attribute VB_Name = "Handler"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow& Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long)
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function getparent Lib "user32" Alias "GetParent" (ByVal hWnd As Long) As Long

Private Sub Command1_Click()
ShowWindow FindWindow(vbNullString, List1.Text), 0
End Sub


Private Sub Command10_Click()
EnableWindow FindWindow(vbNullString, List1.Text), 0
End Sub

Private Sub Command11_Click()
SetWindowPos FindWindow(vbNullString, List1.Text), -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Command12_Click()
SetWindowPos FindWindow(vbNullString, List1.Text), -2, 0, 0, 0, 0, &H1 Or &H2
End Sub



Private Sub Command13_Click()
List1.Clear
Dim TN As String
wndow = GetWindow(Me.hWnd, 0)
While wndow <> 0
TN = Space(GetWindowTextLength(wndow) + 1)
If GetWindowText(wndow, TN, GetWindowTextLength(wndow) + 1) > 0 And Left(TN, Len(TN) - 1) <> Me.Caption Then List1.AddItem Left(TN, Len(TN) - 1)
wndow = GetWindow(wndow, 2)
Wend
End Sub

Private Sub Command14_Click()
SetWindowPos FindWindow(vbNullString, List1.Text), 0, (Text5), (Text4), (Text2), (Text1), 0
End Sub

Private Sub Command15_Click()
SetWindowText FindWindow(vbNullString, List1.Text), Text3
End Sub


Private Sub Command2_Click()
ShowWindow FindWindow(vbNullString, List1.Text), 5
End Sub

Private Sub Command3_Click()
ShowWindow FindWindow(vbNullString, List1.Text), 6
End Sub

Private Sub Command4_Click()
ShowWindow FindWindow(vbNullString, List1.Text), 3
End Sub

Private Sub Command5_Click()
ShowWindow FindWindow(vbNullString, List1.Text), 9
End Sub

Private Sub Command6_Click()
SendMessageA FindWindow(vbNullString, List1.Text), &H10, 0, 0
End Sub

Private Sub Command7_Click()
FlashWindow FindWindow(vbNullString, List1.Text), 1
End Sub

Private Sub Command8_Click()
BringWindowToTop FindWindow(vbNullString, List1.Text)
End Sub


Private Sub Command9_Click()
EnableWindow FindWindow(vbNullString, List1.Text), 1
End Sub

Private Sub Form_Load()
Command13_Click
End Sub

Private Sub List1_DblClick()
Text3 = List1.Text
End Sub


