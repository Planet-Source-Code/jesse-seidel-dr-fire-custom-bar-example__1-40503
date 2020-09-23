VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   9615
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   9615
      TabIndex        =   4
      Top             =   720
      Width           =   9615
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   3480
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Open"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   9600
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Preview:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "  Custom Bar Example"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
   Begin VB.Line Line3 
      X1              =   9840
      X2              =   9840
      Y1              =   8760
      Y2              =   240
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9840
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   8760
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Number As Integer
Dim Filename As String * 100
Dim Length As Long
Dim s
Dim i
Dim g

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Command1_Click()
If cd1.Filename = "" Then Exit Sub
Image1.Picture = Picture1.Picture
End Sub

Private Sub Command2_Click()
On Local Error GoTo Errout
cd1.Filter = "MATRIX Bar Skin (*.MBS)|*.mbs"
cd1.ShowOpen
Picture1.Picture = LoadPicture(cd1.Filename)

Errout:

End Sub

Private Sub Form_Load()
On Error GoTo err
Image1.Picture = LoadPicture("C:\mbs.bmp")

err:
Exit Sub

End Sub

Private Sub Form_Unload(Cancel As Integer)
SavePicture Image1.Picture, "C:\mbs.bmp"
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Label3_Click()
Unload Me
End Sub
