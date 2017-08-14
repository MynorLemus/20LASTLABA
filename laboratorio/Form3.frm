VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   Caption         =   "Form3"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form3"
   ScaleHeight     =   10425
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command77 
      BackColor       =   &H008080FF&
      Caption         =   "Menu"
      Height          =   735
      Left            =   1320
      TabIndex        =   27
      Top             =   8640
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "cod"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox Text3 
      DataField       =   "contraseña"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   5760
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      DataField       =   "Email"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   5160
      Width           =   5535
   End
   Begin VB.TextBox Text5 
      DataField       =   "Telefono"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4560
      Width           =   5535
   End
   Begin VB.TextBox Text99 
      DataField       =   "Direccion"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox Text7 
      DataField       =   "DNI"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   3360
      Width           =   5535
   End
   Begin VB.TextBox Text8 
      DataField       =   "apellido"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2760
      Width           =   5535
   End
   Begin VB.TextBox Text9 
      DataField       =   "Nombre"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2160
      Width           =   5535
   End
   Begin VB.PictureBox Picture1 
      Height          =   4095
      Left            =   10560
      ScaleHeight     =   4035
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   1560
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4560
      TabIndex        =   6
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   7200
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   7800
      TabIndex        =   2
      Top             =   7920
      Width           =   3015
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Explorar"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   8640
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      DataField       =   "foto"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6360
      Width           =   5535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5520
      Top             =   10920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   1200
      Top             =   10560
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "profesores"
      Height          =   375
      Left            =   3480
      TabIndex        =   26
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Codigo."
      Height          =   375
      Left            =   1200
      TabIndex        =   25
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Nombre."
      Height          =   375
      Left            =   1200
      TabIndex        =   24
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Apellido"
      Height          =   375
      Left            =   1200
      TabIndex        =   23
      Top             =   2760
      Width           =   2775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "DNI"
      Height          =   375
      Left            =   1200
      TabIndex        =   22
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Email"
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Foto"
      Height          =   375
      Left            =   1200
      TabIndex        =   20
      Top             =   6360
      Width           =   2775
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Direccion."
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   3960
      Width           =   2775
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Telefono"
      Height          =   375
      Left            =   1200
      TabIndex        =   18
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Password."
      Height          =   375
      Left            =   1200
      TabIndex        =   17
      Top             =   5760
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
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
Adodc1.RecordSource = "profesor" '
Adodc1.Refresh '
Adodc1.Recordset.MoveFirst '
x = App.Path '
If Text6 = Error Then '
Text6.Text = "q.jpg" '
Else '
Picture1.Picture = LoadPicture(x & "/" & Text6.Text) '
End If '
End Sub '

