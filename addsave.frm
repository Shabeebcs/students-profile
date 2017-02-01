VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form addsave 
   Caption         =   "Form3"
   ClientHeight    =   11730
   ClientLeft      =   1305
   ClientTop       =   855
   ClientWidth     =   19860
   LinkTopic       =   "Form3"
   ScaleHeight     =   11730
   ScaleWidth      =   19860
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
      TabIndex        =   41
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
      TabIndex        =   40
      Top             =   5880
      Width           =   3135
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
      Height          =   465
      Left            =   2040
      TabIndex        =   19
      Top             =   2520
      Width           =   3735
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
      Height          =   465
      Left            =   2040
      TabIndex        =   18
      Top             =   1080
      Width           =   3735
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
      TabIndex        =   16
      Text            =   "Select Department"
      Top             =   5130
      Width           =   3735
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
      TabIndex        =   15
      Text            =   "Select Course"
      Top             =   6000
      Width           =   3735
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
      TabIndex        =   14
      Text            =   "Select Semester"
      Top             =   6840
      Width           =   3735
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
      Height          =   1245
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   7560
      Width           =   3735
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
      Height          =   465
      Left            =   2040
      TabIndex        =   12
      Top             =   9960
      Width           =   3735
   End
   Begin VB.CommandButton addnew 
      Caption         =   "ADD NEW"
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
      Left            =   6120
      Picture         =   "addsave.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   960
      Width           =   2055
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
      Left            =   6120
      Picture         =   "addsave.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.CommandButton uploadbtn 
      Caption         =   "UPLOAD PICTURE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17160
      Picture         =   "addsave.frx":2104
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   2295
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
      Height          =   465
      Left            =   2040
      TabIndex        =   8
      Top             =   1800
      Width           =   3735
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
      TabIndex        =   7
      Top             =   5760
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   8160
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
      Height          =   525
      Left            =   9720
      TabIndex        =   4
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
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   6960
      Width           =   3135
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
      Height          =   465
      Left            =   2040
      TabIndex        =   2
      Top             =   9150
      Width           =   3735
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
      Top             =   9120
      Width           =   3135
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   18600
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Bindings        =   "addsave.frx":3186
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
      TabIndex        =   17
      Top             =   3360
      Width           =   3735
      _ExtentX        =   6588
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
      Format          =   131727361
      CurrentDate     =   42751
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4215
      Left            =   16200
      Stretch         =   -1  'True
      Top             =   360
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
      TabIndex        =   39
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
      Left            =   240
      TabIndex        =   38
      Top             =   2520
      Width           =   855
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
      Left            =   240
      TabIndex        =   37
      Top             =   1080
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
      Left            =   240
      TabIndex        =   36
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
      Left            =   240
      TabIndex        =   35
      Top             =   4320
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
      Left            =   240
      TabIndex        =   34
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
      Left            =   240
      TabIndex        =   33
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
      Left            =   240
      TabIndex        =   32
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
      Left            =   240
      TabIndex        =   31
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
      Left            =   240
      TabIndex        =   30
      Top             =   9960
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
      Left            =   240
      TabIndex        =   29
      Top             =   1800
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
      Left            =   8040
      TabIndex        =   28
      Top             =   120
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
      TabIndex        =   27
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
      Top             =   9240
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
      TabIndex        =   23
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
      Left            =   240
      TabIndex        =   22
      Top             =   9240
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
      TabIndex        =   21
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
      Height          =   255
      Left            =   12480
      TabIndex        =   20
      Top             =   9240
      Width           =   615
   End
End
Attribute VB_Name = "addsave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim str As String

Private Sub addnew_Click()
rs.addnew
addnew.Enabled = False
savebtn.Enabled = True
End Sub

Private Sub Form_Load()
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\college\ProfileDB.mdb;Persist Security Info=False"
rs.Open "Select * from ProfileTBL", con, adOpenDynamic, adLockPessimistic
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
savebtn.Enabled = False
End Sub

Sub clear()
Text2.Text = ""
Text5.Text = ""
Text1.Text = ""
DTPicker1.Value = "10/05/2005"
Text15.Text = ""
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



Private Sub savebtn_Click()
rs.Fields("Name").Value = Text2.Text
rs.Fields("Year").Value = Text5.Text
rs.Fields("RollNo").Value = Text1.Text
rs.Fields("DOB").Value = DTPicker1.Value
rs.Fields("Gender").Value = Text15.Text
rs.Fields("Dept").Value = Combo1.Text
rs.Fields("Course").Value = Combo2.Text
rs.Fields("Semester").Value = Combo3.Text
rs.Fields("Address").Value = Text3.Text
rs.Fields("Email").Value = Text11.Text
rs.Fields("Phone").Value = Text4.Text
rs.Fields("Photo").Value = str
rs.Fields("Tenth").Value = Text6.Text
rs.Fields("PlusTwo").Value = Text7.Text
rs.Fields("Btech").Value = Text8.Text
rs.Fields("Backlogs").Value = Text9.Text
rs.Fields("AddonCourse").Value = Text14.Text
rs.Fields("PlacedCompany").Value = Text10.Text
rs.Fields("Designation").Value = Text12.Text
rs.Fields("Salary").Value = Text13.Text
MsgBox "DATA IS SAVED SUCCESSFULLY ..!!!", vbInformation
rs.Update
clear
savebtn.Enabled = False
addnew.Enabled = True
End Sub


Private Sub uploadbtn_Click()
CommonDialog1.ShowOpen
str = CommonDialog1.FileName
Image1.Picture = LoadPicture(str)
End Sub

    
