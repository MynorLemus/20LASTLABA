VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form4"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "@Adobe Gothic Std B"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   8445
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command77 
      BackColor       =   &H008080FF&
      Caption         =   "Menu"
      Height          =   735
      Left            =   5040
      TabIndex        =   15
      Top             =   5880
      Width           =   3015
   End
   Begin VB.CommandButton Command6 
      Caption         =   "eliminar"
      Height          =   495
      Left            =   7560
      TabIndex        =   14
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ultimo"
      Height          =   495
      Left            =   4320
      TabIndex        =   13
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Nuevo"
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      Top             =   4800
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Siguiente"
      Height          =   495
      Left            =   7560
      TabIndex        =   11
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "anterior"
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Data Adodc1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\estudiante\Desktop\20LASTLABA\Visual Basic.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   1095
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "notas"
      Top             =   5760
      Width           =   3135
   End
   Begin VB.TextBox Text9 
      DataField       =   "unidad"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Width           =   5535
   End
   Begin VB.TextBox Text8 
      DataField       =   "promedio"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox Text7 
      DataField       =   "idcurso"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      DataField       =   "idalum"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "id curso"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Promedio"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ID alumno"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unidad"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nota"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form4"
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
