VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form view 
   Caption         =   "Form4"
   ClientHeight    =   12285
   ClientLeft      =   2085
   ClientTop       =   270
   ClientWidth     =   20340
   LinkTopic       =   "Form4"
   ScaleHeight     =   12285
   ScaleWidth      =   20340
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "view.frx":0000
      Height          =   855
      Left            =   12480
      TabIndex        =   46
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1508
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   12360
      Top             =   720
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   873
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
   Begin VB.TextBox Text15 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   45
      Top             =   4200
      Width           =   3615
   End
   Begin VB.TextBox Text14 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   43
      Top             =   5760
      Width           =   3135
   End
   Begin VB.CommandButton updatebtn 
      Caption         =   "UPDATE"
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
      Left            =   9960
      Picture         =   "view.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   10560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "RollNo"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   22
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   840
      Width           =   3615
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "Dept"
      DataSource      =   "Adodc1"
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
      Left            =   2040
      TabIndex        =   19
      Text            =   "Select Department"
      Top             =   5040
      Width           =   3615
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "Course"
      DataSource      =   "Adodc1"
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
      Left            =   2040
      TabIndex        =   18
      Text            =   "Select Course"
      Top             =   6000
      Width           =   3615
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "Semester"
      DataSource      =   "Adodc1"
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
      Left            =   2040
      TabIndex        =   17
      Text            =   "Select Semester"
      Top             =   6840
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2040
      TabIndex        =   16
      Top             =   7560
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   15
      Top             =   9960
      Width           =   3615
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
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
      Left            =   6240
      Picture         =   "view.frx":1097
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   840
      Width           =   1815
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
      Left            =   6240
      Picture         =   "view.frx":2119
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      DataField       =   "Year"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text6 
      DataField       =   "Tenth"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9720
      TabIndex        =   11
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      DataField       =   "PlusTwo"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9720
      TabIndex        =   10
      Top             =   6960
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      DataField       =   "Btech"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      Top             =   8040
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      DataField       =   "Backlogs"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   8
      Top             =   9120
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      DataField       =   "PlacedCompany"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   7
      Top             =   6960
      Width           =   3135
   End
   Begin VB.CommandButton Firstbtn 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   1320
      Picture         =   "view.frx":319B
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   10800
      Width           =   1095
   End
   Begin VB.CommandButton Previousbtn 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   2760
      Picture         =   "view.frx":421D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10800
      Width           =   1215
   End
   Begin VB.CommandButton Nextbtn 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      Picture         =   "view.frx":529F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   10800
      Width           =   1215
   End
   Begin VB.CommandButton Lastbtn 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      Picture         =   "view.frx":6321
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   10800
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   9000
      Width           =   3615
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   1
      Top             =   8040
      Width           =   3135
   End
   Begin VB.TextBox Text13 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15000
      TabIndex        =   0
      Top             =   9000
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   18600
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Bindings        =   "view.frx":73A3
      DataField       =   "DOB"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   2040
      TabIndex        =   20
      Top             =   3360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   132055041
      CurrentDate     =   42751
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4215
      Left            =   15240
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label17 
      Caption         =   "Addon Course"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   44
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Roll No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   41
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   40
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   39
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   38
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   37
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   36
      Top             =   6000
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   35
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   34
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      DataField       =   "Phone"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "STUDENT PROFILE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   31
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "10th  Percentage"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   30
      Top             =   5880
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "12th Percentage"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   29
      Top             =   7080
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "Btech Percentage/CGPA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   28
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label15 
      Caption         =   "Backlogs"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   27
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Placed company"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   26
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   9120
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   24
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12480
      TabIndex        =   23
      Top             =   9120
      Width           =   1215
   End
End
Attribute VB_Name = "view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Public str As String
Public confirm As Integer
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
Sub clear()
Text2.Text = ""
Text5.Text = ""
Text1.Text = ""
Text15.Text = ""
DTPicker1.Value = "10/05/2005"
Combo1.Text = "Select Department"
Combo2.Text = "Select Course"
Combo3.Text = "Select Semester"
Text3.Text = ""
Text11.Text = ""
Text4.Text = ""
Image1.Picture = LoadPicture("")
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text14.Text = ""
Text10.Text = ""
Text12.Text = ""
Text13.Text = ""
End Sub


Private Sub deletebtn_Click()
confirm = MsgBox("Do you want to delete the Student Profile", vbYesNo + vbCritical, "Deletion Confirmation")
If confirm = vbYes Then
rss.Delete adAffectCurrent
MsgBox "Record has been Deleted successfully", vbInformation, "Message"
rss.Update
refreshdata
Else
MsgBox "Profile Not Deleted ..!!", vbInformation, "Message"
Call clear
End If
End Sub
Sub refreshdata()
rss.Close
rss.Open "Select * from ProfileTBL", conn, adOpenStatic, adLockPessimistic
If rss.EOF Then
MsgBox "No Record Found"
End If
End Sub

Private Sub Form_Load()
subsearch.Hide
connect
Set DataGrid1.DataSource = rss
Combo1.AddItem "Computer Science and Engineering"
Combo1.AddItem "Electrical and Electronics Engineering"
Combo1.AddItem "Civil Engineering"
Combo1.AddItem "Mechanical Engineering"
Combo1.AddItem "Bio-Technology Engineering"
Combo1.AddItem "Electronics and Communiation Engineering"
Combo2.AddItem "BTECH"
Combo3.AddItem "S1"
Combo3.AddItem "S2"
Combo3.AddItem "S3"
Combo3.AddItem "S4"
Combo3.AddItem "S5"
Combo3.AddItem "S6"
Combo3.AddItem "S7"
Combo3.AddItem "S8"
Combo3.AddItem "PASSOUT"
display
subsearch.Show
End Sub

Sub display()
Text2.Text = rss!Name
Text5.Text = rss!Year
Text1.Text = rss!Rollno
Text15.Text = rss!Gender
DTPicker1.Value = rss!DOB
Combo1.Text = rss!Dept
Combo2.Text = rss!Course
Combo3.Text = rss!Semester
Text3.Text = rss!Address
Text11.Text = rss!Email
Text4.Text = rss!phone
Image1.Picture = LoadPicture(rss!Photo)
Text6.Text = rss!Tenth
Text7.Text = rss!PlusTwo
Text8.Text = rss!Btech
Text9.Text = rss!Backlogs
Text14.Text = rss!AddonCourse
Text10.Text = rss!PlacedCompany
Text12.Text = rss!Designation
Text13.Text = rss!Salary
End Sub


Private Sub savebtn_Click()
rss.addnew
rss.Fields("Name").Value = Text2.Text
rss.Fields("Year").Value = Text5.Text
rss.Fields("RollNo").Value = Text1.Text
rss.Fields("DOB").Value = DTPicker1.Value
rss.Fields("Gender").Value = Text15.Text
rss.Fields("Dept").Value = Combo1.Text
rss.Fields("Course").Value = Combo2.Text
rss.Fields("Semester").Value = Combo3.Text
rss.Fields("Address").Value = Text3.Text
rss.Fields("Email").Value = Text11.Text
rss.Fields("Phone").Value = Text4.Text
rss.Fields("Photo").Value = str
rss.Fields("Tenth").Value = Text6.Text
rss.Fields("PlusTwo").Value = Text7.Text
rss.Fields("Btech").Value = Text8.Text
rss.Fields("Backlogs").Value = Text9.Text
rss.Fields("AddonCourse").Value = Text14.Text
rss.Fields("PlacedCompany").Value = Text10.Text
rss.Fields("Designation").Value = Text12.Text
rss.Fields("Salary").Value = Text13.Text
rss.Update
refreshdata
MsgBox "DATA IS SAVED SUCCESSFULLY ..!!!", vbInformation
End Sub

Private Sub updatebtn_Click()
Adodc1.Recordset.Fields("Name") = Text2.Text
Adodc1.Recordset.Fields("Year") = Text5.Text
Adodc1.Recordset.Fields("RollNo") = Text1.Text
Adodc1.Recordset.Fields("DOB") = DTPicker1.Value
Adodc1.Recordset.Fields("Gender") = Text15.Text
Adodc1.Recordset.Fields("Dept") = Combo1.Text
Adodc1.Recordset.Fields("Course") = Combo2.Text
Adodc1.Recordset.Fields("Semester") = Combo3.Text
Adodc1.Recordset.Fields("Address") = Text3.Text
Adodc1.Recordset.Fields("Email") = Text11.Text
Adodc1.Recordset.Fields("Phone") = Text4.Text
Adodc1.Recordset.Fields("Photo") = str
Adodc1.Recordset.Fields("Tenth") = Text6.Text
Adodc1.Recordset.Fields("PlusTwo") = Text7.Text
Adodc1.Recordset.Fields("Btech") = Text8.Text
Adodc1.Recordset.Fields("Backlogs") = Text9.Text
Adodc1.Recordset.Fields("AddonCourse") = Text14.Text
Adodc1.Recordset.Fields("PlacedCompany") = Text10.Text
Adodc1.Recordset.Fields("Designation") = Text12.Text
Adodc1.Recordset.Fields("Salary") = Text13.Text
Adodc1.Recordset.Update
refreshdata
MsgBox "DATA IS UPDATED SUCCESSFULLY ..!!!", vbInformation
End Sub
