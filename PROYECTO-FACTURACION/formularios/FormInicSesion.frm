VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form2"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11025
   LinkTopic       =   "Form2"
   Picture         =   "FormInicSesion.frx":0000
   ScaleHeight     =   6615
   ScaleWidth      =   11025
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2535
      Left            =   720
      Picture         =   "FormInicSesion.frx":2524D
      ScaleHeight     =   2475
      ScaleWidth      =   4035
      TabIndex        =   8
      Top             =   2880
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   1935
      Left            =   8520
      Picture         =   "FormInicSesion.frx":26751
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Entrar"
      Height          =   1935
      Left            =   6000
      Picture         =   "FormInicSesion.frx":26A5B
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   720
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "Cargo"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Line Line2 
      X1              =   5640
      X2              =   5640
      Y1              =   5880
      Y2              =   6600
   End
   Begin VB.Line Line1 
      X1              =   5280
      X2              =   5280
      Y1              =   5880
      Y2              =   6600
   End
   Begin VB.Label Label2 
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   5655
      Left            =   120
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1 = "proveedor" And Text2 = "iiivvv" Then
 s = MsgBox("Bienvenido proveedor", 0, ("Bienvenido"))
 Form1.Hide
 Form2.Show
 Form2.admin.Enabled = False
 Form2.Frame2.Enabled = False
ElseIf Text1 = "administrador" And Text2 = "vvviii" Then
 z = MsgBox("Bienvenido administrador", 0, ("¡Bienvenido administrador!"))
 Form1.Hide
 Form2.Show
Else
 e = MsgBox("Error, verifique su usuario o contraseña", vbCritical, ("Error"))
 Text1.SetFocus
End If
Label3.Caption = "Usuario"
Label4.Caption = "Cargo"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Activate()
Text1.SetFocus
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  Command1.SetFocus
 End If
 If Text2 = "" Then
  If KeyAscii = 8 Then
   Text1.SetFocus
  End If
End If
If Text1 = "proveedor" Or Text1 = "PROVEEDOR" Or Text1 = "administrador" Or Text1 = "ADMINISTRADOR" Then
 Label3.Caption = Text1.Text
 Label4.Caption = "Bienvenido"
Else
 Label3.Caption = "Usuario no identificado"
 Label4.Caption = "Error"
End If
End Sub
