VERSION 5.00
Begin VB.Form addsearch 
   Caption         =   "Form2"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   4920
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Editbtn 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Picture         =   "addsearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Picture         =   "addsearch.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7680
      TabIndex        =   2
      Top             =   480
      Width           =   1110
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      Picture         =   "addsearch.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD STUDENT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      Picture         =   "addsearch.frx":3186
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Total Records"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "addsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As ADODB.Connection
Public rss As ADODB.Recordset
Public str As String
Sub connect()
Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\college\ProfileDB.mdb;Persist Security Info=False"
Set rss = New ADODB.Recordset
    rss.ActiveConnection = conn
    rss.CursorLocation = adUseClient
    rss.CursorType = adOpenDynamic
    rss.LockType = adLockOptimistic
    rss.Source = "SELECT * FROM ProfileTBL"
    rss.Open

End Sub
Sub reload()
rss.Close
rss.Open "Select * from ProfileTBL", conn, adOpenDynamic, adLockPessimistic
End Sub


Private Sub Command1_Click()
addsave.Show
End Sub

Private Sub Command2_Click()
subsearch.Show
End Sub

Private Sub deletebtn_Click()
subsearch.Show
End Sub

Private Sub Editbtn_Click()
subsearch.Show
End Sub

Private Sub Form_Load()
connect
Text1.Text = rss.RecordCount
reload
End Sub

