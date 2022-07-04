VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\Kuliah\Semester 4\Pemrograman\Lat 19\Database.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TCUSTOMER"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SELESAI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6600
      TabIndex        =   4
      Top             =   8040
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":0000
      Height          =   3255
      Left            =   3960
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   3
      Top             =   3720
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PROSES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6960
      TabIndex        =   2
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   5520
      TabIndex        =   0
      Top             =   1800
      Width           =   6375
   End
   Begin VB.Label Label3 
      Caption         =   "Perintah SQL:"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "ini caption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   7200
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Perintah Dasar"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AWAL As String
Private Sub Command1_Click()
    On Error GoTo SALAH
    
    SQLTEKS = "SELECT * FROM TCUSTOMER ORDER BY " + Text1.Text
    Data1.RecordSource = SQLTEKS
    Data1.Refresh
    DBGrid1.Refresh
    
    Label2.Caption = SQLTEKS
    KOTACUST = Text1.Text
    Exit Sub
    
SALAH:
    MsgBox "Anda salah menuliskan nama FIELD", vbExclamation
    Exit Sub
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Activate()
    AWAL = " "
    Text1.Text = AWAL
    Label2.Caption = ""
    
    Form1.WindowState = 2
End Sub

