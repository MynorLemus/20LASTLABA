VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H0000FF00&
   Caption         =   "Form5"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   7605
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command77 
      BackColor       =   &H008080FF&
      Caption         =   "Menu"
      Height          =   735
      Left            =   5640
      TabIndex        =   11
      Top             =   6120
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "nombre"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox Text9 
      DataField       =   "iprofesor"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Data Adodc1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\20LASTLABA\Visual Basic.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   1095
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "curso"
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   5040
      TabIndex        =   4
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   4440
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
      Top             =   5160
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   8280
      TabIndex        =   0
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cursos"
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Id profesor"
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Materia"
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1800
      Width           =   2775
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() '
Adodc1.Recordset.MovePrevious '
If Adodc1.Recordset.BOF Then '
Adodc1.Recordset.MoveNext '
End If '

End Sub '
Private Sub Command2_Click() '
Adodc1.Recordset.Update '
Adodc1.Recordset.MoveFirst '

End Sub '
Private Sub Command3_Click() '
Adodc1.Recordset.MoveNext '
If Adodc1.Recordset.EOF Then '
Adodc1.Recordset.MovePrevious '
End If '

End Sub '
Private Sub Command4_Click() '
Adodc1.Recordset.AddNew '
End Sub '
Private Sub Command5_Click() '
Adodc1.Recordset.MoveLast '

End Sub '
Private Sub Command6_Click() '
Adodc1.Recordset.Delete '
End Sub '

Private Sub Command77_Click()
Form2.Show
End Sub
