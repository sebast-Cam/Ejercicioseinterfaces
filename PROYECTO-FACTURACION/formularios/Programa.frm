VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10320
   LinkTopic       =   "Form4"
   Picture         =   "Programa.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   10320
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Colocar empleados"
      Height          =   255
      Left            =   960
      TabIndex        =   26
      Top             =   8400
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2400
      TabIndex        =   25
      Text            =   "Text2"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Ingresar Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   23
      Top             =   8280
      Width           =   1695
   End
   Begin VB.ListBox List3 
      Height          =   2595
      Left            =   7080
      TabIndex        =   22
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ingresar Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cargos"
      Height          =   1575
      Left            =   600
      TabIndex        =   16
      Top             =   3240
      Width           =   8895
      Begin VB.OptionButton Option3 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   19
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cajero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   18
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Administrador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   855
      Left            =   6360
      TabIndex        =   15
      Top             =   9000
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar"
      Height          =   855
      Left            =   3600
      TabIndex        =   14
      Top             =   9000
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Añadir"
      Height          =   855
      Left            =   840
      TabIndex        =   13
      Top             =   9000
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   5160
      TabIndex        =   10
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   720
      TabIndex        =   7
      Top             =   5520
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   5520
      Top             =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Hora Retiro"
      Height          =   255
      Left            =   7200
      TabIndex        =   21
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label10 
      Caption         =   "Hora Ingreso"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Empleados"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Hora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   8640
      TabIndex        =   6
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   6960
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Empleados que terminaron su turno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8280
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Empleados Activos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6600
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Empleado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Registro de Empleados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1 = 0 Then
   Command1.Enabled = False
   Command2.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
Else
   Command1.Enabled = True
   Command2.Enabled = True
   Command4.Enabled = True
   Command5.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Text1.Text <> "" Then
  If Option1.Value = True Then
      List1.AddItem Text1.Text + " " + Option1.Caption
  ElseIf Option2.Value = True Then
      List1.AddItem Text1.Text + " " + Option2.Caption
  ElseIf Option3.Value = True Then
      List1.AddItem Text1.Text + " " + Option3.Caption
  Else
      MsgBox "Seleccione alguno de los cargos", vbExclamation, "Aviso"
  End If
Else
  MsgBox "Debe colocar el nombre del empleado", vbExclamation, "Aviso"
End If
If Text1 <> "" Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
End If
If List1.ListCount = 0 Then
   Command4.Enabled = False
   Command5.Enabled = False
Else
   Command4.Enabled = True
   Command5.Enabled = True
End If
End Sub

Private Sub Command2_Click()
o = MsgBox("¿Desea eliminar todo el registro de empleados?", vbYesNo, "Confirmación")
If o = vbYes Then
   If List1.ListCount > 0 Then
      List1.Clear
      List2.Clear
      List3.Clear
      co1 = 0
      co2 = 0
      Label5.Caption = co1
      Label6.Caption = co2
      Text1.Text = ""
      Text1.SetFocus
    End If
ElseIf o = vbNo Then
   MsgBox "No se elimino el registro de empleados", vbInformation, "Confirmación"
End If
End Sub



Private Sub Command3_Click()
Form4.Hide
Form2.Show
End Sub

Private Sub Command4_Click()
List2.AddItem Text2.Text
Dim co1 As Integer
For co1 = 0 To List2.ListCount - 1
Next
Label5.Caption = co1
End Sub

Private Sub Command5_Click()
List3.AddItem Text2.Text
Dim co2 As Integer
For co2 = 0 To List3.ListCount - 1
Next
Label6.Caption = co2
End Sub

Private Sub Form_Load()
If Check1 = 0 Then
   Command1.Enabled = False
   Command2.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
Else
   Command1.Enabled = True
   Command2.Enabled = True
   Command4.Enabled = True
   Command5.Enabled = True
End If
If Text1 <> "" Then
    Command1.Enabled = True
    Command2.Enabled = True
Else
    Command1.Enabled = False
    Command2.Enabled = False
End If
If List1.ListCount = 0 Then
   Command4.Enabled = False
   Command5.Enabled = False
Else
   Command4.Enabled = True
   Command5.Enabled = True
End If
End Sub



Private Sub Text1_Change()
If Text1 <> "" Then
    Command1.Enabled = True
    Command2.Enabled = True
Else
    Command1.Enabled = False
    Command2.Enabled = False
End If
If List1.ListCount = 0 Then
   Command4.Enabled = False
   Command5.Enabled = False
Else
   Command4.Enabled = True
   Command5.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Text2.Text = Time
Text3.Text = Date
End Sub
