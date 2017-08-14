VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000000FF&
   Caption         =   "Form2"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Adobe Gothic Std B"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H008080FF&
      Caption         =   "Notas"
      Height          =   735
      Left            =   5160
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Curso"
      Height          =   735
      Left            =   1200
      TabIndex        =   4
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Profesores"
      Height          =   735
      Left            =   5160
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   4080
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Alumnos"
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
End Sub

Private Sub Command2_Click()

If MsgBox("¿En serio, quieres salir?", vbQuestion + vbYesNo, "SALIR") = vbYes Then
End '
Else '
Cancel = Value
End If
End Sub '


Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub

Private Sub Command5_Click()
Form4.Show
End Sub
