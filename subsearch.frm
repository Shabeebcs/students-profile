VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form subsearch 
   Caption         =   "Form1"
   ClientHeight    =   10020
   ClientLeft      =   1890
   ClientTop       =   945
   ClientWidth     =   18675
   LinkTopic       =   "Form1"
   ScaleHeight     =   10020
   ScaleWidth      =   18675
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\college\ProfileDB.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\college\ProfileDB.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ProfileTBL"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
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
      Height          =   405
      Left            =   14640
      TabIndex        =   7
      Top             =   1320
      Width           =   1110
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   11520
      Picture         =   "subsearch.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1035
   End
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7200
      TabIndex        =   0
      ToolTipText     =   "Type Name to be Searched Here..."
      Top             =   1320
      Width           =   3855
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "subsearch.frx":1082
      Height          =   3735
      Left            =   960
      TabIndex        =   2
      Top             =   3480
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   6588
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   26
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " Records"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13200
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Search "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "subsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String
Dim confirm As Integer
Public conn As ADODB.Connection
Public rss As ADODB.Recordset
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

Private Sub Command1_Click()
rss.Filter = adFilterNone
rss.Requery
Text1.Text = ""
txtSearch.Text = ""
Combo1.Text = ""
End Sub

Private Sub DataGrid1_Click()
view.Text1.Text = DataGrid1.Columns(0).Text
view.Text2.Text = DataGrid1.Columns(1).Text
view.Text5.Text = DataGrid1.Columns(2).Text
view.DTPicker1.Value = DataGrid1.Columns(3).Text
view.Text15.Text = DataGrid1.Columns(4).Text
view.Combo1.Text = DataGrid1.Columns(5).Text
view.Combo2.Text = DataGrid1.Columns(6).Text
view.Combo3.Text = DataGrid1.Columns(7).Text
view.Text3.Text = DataGrid1.Columns(8).Text
view.Text11.Text = DataGrid1.Columns(9).Text
view.Text4.Text = DataGrid1.Columns(10).Text
view.Image1.Picture = LoadPicture(DataGrid1.Columns(11).Text)
view.Text6.Text = DataGrid1.Columns(12).Text
view.Text7.Text = DataGrid1.Columns(13).Text
view.Text8.Text = DataGrid1.Columns(14).Text
view.Text9.Text = DataGrid1.Columns(15).Text
view.Text14.Text = DataGrid1.Columns(16).Text
view.Text10.Text = DataGrid1.Columns(17).Text
view.Text12.Text = DataGrid1.Columns(18).Text
view.Text13.Text = DataGrid1.Columns(19).Text
view.Show
End Sub

Private Sub Form_Load()
connect
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\college\ProfileDB.mdb;Persist Security Info=False"
'rs.Open "Select * from ProfileTBL", con, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rss
Combo1.AddItem "Name"
Combo1.AddItem "RollNo"
Combo1.AddItem "Dept"
Combo1.AddItem "Semester"
Combo1.AddItem "Year"
Combo1.AddItem "Backlogs"
Combo1.AddItem "Btech"
Combo1.AddItem "AddonCourse"
End Sub

Private Sub txtSearch_Change()
If txtSearch.Text = "" Then
            Me.Show
        ElseIf (Combo1.Text = "Name") Then
            rss.Filter = "Name Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
            Text1.Text = rss.RecordCount
        ElseIf (Combo1.Text = "RollNo") Then
           rss.Filter = "RollNo Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
        ElseIf (Combo1.Text = "Dept") Then
           rss.Filter = "Dept Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
        ElseIf (Combo1.Text = "Semester") Then
           rss.Filter = "Semester Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
        ElseIf (Combo1.Text = "Year") Then
           rss.Filter = "Year Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
        ElseIf (Combo1.Text = "Backlogs") Then
           rss.Filter = "Backlogs Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
        ElseIf (Combo1.Text = "Btech") Then
           rss.Filter = "Btech Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
        Else
           rss.Filter = "AddonCourse Like '" & Me.txtSearch.Text & "*'"
           Set DataGrid1.DataSource = rss
           Text1.Text = rss.RecordCount
            
        
End If
End Sub

Private Sub txtSearch_Click()
DataGrid1.Refresh
End Sub
