VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15255
   BeginProperty Font 
      Name            =   "Cooper Black"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command77 
      BackColor       =   &H008080FF&
      Caption         =   "Menu"
      Height          =   735
      Left            =   11040
      TabIndex        =   31
      Top             =   6840
      Width           =   3015
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   10680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   840
      Top             =   10320
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cooper Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   30
      Text            =   "Text1"
      Top             =   7320
      Width           =   5535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Explorar"
      Height          =   495
      Left            =   3960
      TabIndex        =   28
      Top             =   9480
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   7200
      TabIndex        =   27
      Top             =   8760
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   3960
      TabIndex        =   26
      Top             =   8760
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   840
      TabIndex        =   25
      Top             =   8760
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   7200
      TabIndex        =   24
      Top             =   8040
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   3960
      TabIndex        =   23
      Top             =   8040
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   495
      Left            =   840
      TabIndex        =   22
      Top             =   8040
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   10200
      ScaleHeight     =   4035
      ScaleWidth      =   3795
      TabIndex        =   21
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Text10 
      DataField       =   "contra"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   6720
      Width           =   5535
   End
   Begin VB.TextBox Text9 
      DataField       =   "Nnombre"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   1920
      Width           =   5535
   End
   Begin VB.TextBox Text8 
      DataField       =   "Apelliso"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2520
      Width           =   5535
   End
   Begin VB.TextBox Text7 
      DataField       =   "fechanac"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3120
      Width           =   5535
   End
   Begin VB.TextBox Text99 
      DataField       =   "grado"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   3720
      Width           =   5535
   End
   Begin VB.TextBox Text5 
      DataField       =   "seccion"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4320
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      DataField       =   "direccion"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4920
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      DataField       =   "telefono"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5520
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      DataField       =   "email"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   6120
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Password."
      Height          =   375
      Left            =   840
      TabIndex        =   29
      Top             =   7320
      Width           =   2775
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Password."
      Height          =   375
      Left            =   840
      TabIndex        =   19
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "Email"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Telefono"
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Direccion."
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   4920
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Seccion."
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   4320
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Grado."
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Fecha_Nac."
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Apellido"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Nombre."
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Codigo."
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Alumnos"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '
Adodc1.Recordset.MovePrevious '
If Adodc1.Recordset.BOF Then '
Adodc1.Recordset.MoveNext '
End If '
x = App.Path '
If Text6 = Error Then '
Text6.Text = "q.jpg" '
Else '
Picture1.Picture = LoadPicture(x & "/" & Text6.Text) '
End If '
End Sub '
Private Sub Command2_Click() '
Adodc1.Recordset.Update '
Adodc1.Recordset.MoveFirst '
x = App.Path '
If Text6 = Error Then '
Text6.Text = "q.jpg" '
Else '
Picture1.Picture = LoadPicture(x & "/" & Text6.Text) '
End If '
End Sub '
Private Sub Command3_Click() '
Adodc1.Recordset.MoveNext '
If Adodc1.Recordset.EOF Then '
Adodc1.Recordset.MovePrevious '
End If '
x = App.Path '
If Text6 = "" Then '
Text6.Text = "q.jpg" '
Else '
Picture1.Picture = LoadPicture(x & "/" & Text6.Text) '
End If '
End Sub '
Private Sub Command4_Click() '
Adodc1.Recordset.AddNew '
End Sub '
Private Sub Command5_Click() '
Adodc1.Recordset.MoveLast '
x = App.Path '
If Text6 = Error Then '
Text6.Text = "q.jpg" '
Else '
Picture1.Picture = LoadPicture(x & "/" & Text6.Text) '
End If '
End Sub '
Private Sub Command6_Click() '
Adodc1.Recordset.Delete '
End Sub '
Private Sub Command7_Click() '
CommonDialog1.ShowOpen '
Picture1.Picture = LoadPicture(CommonDialog1.FileName) '
Text6.Text = CommonDialog1.FileTitle
End Sub
Private Sub Command8_Click()
If MsgBox("¿En serio, quieres salir?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
End '
Else '
Cancel = Value
End If
End Sub '

Private Sub Command77_Click()
Form2.Show
End Sub

Private Sub Form_Load() '
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\estudiante\Desktop\20LASTLABA\Visual Basic.mdb;Persist Security Info=False" '
Adodc1.CursorType = adOpenDynamic '
Adodc1.RecordSource = "alumno" '
Adodc1.Refresh '
Adodc1.Recordset.MoveFirst '
x = App.Path '
If Text6 = Error Then '
Text6.Text = "q.jpg" '
Else '
Picture1.Picture = LoadPicture(x & "/" & Text6.Text) '
End If '
End Sub '

