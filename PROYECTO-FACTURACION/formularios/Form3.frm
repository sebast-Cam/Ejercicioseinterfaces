VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8955
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   3360
      TabIndex        =   11
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar"
      Height          =   615
      Left            =   480
      TabIndex        =   10
      Top             =   5520
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Volver"
      Height          =   615
      Left            =   6240
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Borrar"
      Height          =   615
      Left            =   6240
      TabIndex        =   2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Registrar"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   8
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Nombre del Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "REGISTROS DE PRODUCTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim listaprod As DAO.Database
Dim Producto As DAO.Recordset
Dim rutasgbd As String
rutasgbd = App.Path & "\listaprod.mdb"
 Set listaprod = OpenDatabase(App.Path & "\listaprod.mdb")
 Set Producto = listaprod.OpenRecordset("Producto")
     Producto.Index = "Codigo"
     Producto.Seek "=", Text3.Text
     If Producto.NoMatch Then
        Producto.AddNew
        Producto.Fields("Codigo") = Text3.Text
        Producto.Fields("Nombre") = Text1.Text
        Producto.Fields("Precio") = Text2.Text
        Producto.Update
        Text3.SetFocus
        Text3.Text = ""
        Text2.Text = ""
        Text1.Text = ""
        Producto.Close
        listaprod.Close
        MsgBox ("Se registro el producto"), vbInformation, ("Registro completado")
     Else
        MsgBox ("Ya existe un producto así"), 0, ("Aviso")
     End If
End Sub

Private Sub Command2_Click()
Dim Producto As DAO.Recordset
 Set listaprod = OpenDatabase(App.Path & "\listaprod.mdb")
 Set Producto = listaprod.OpenRecordset("Producto")
     Producto.Index = "Codigo"
     Producto.Seek "=", Text3.Text
     If Producto.NoMatch Then
     MsgBox "Este producto no está registrado", vbExclamation, "Aviso"
     Else
     p = MsgBox("¿Desea eliminar este producto?", vbYesNo, "Eliminar")
      If p = vbYes Then
      Producto.Delete
      Producto.Close
      Text3.SetFocus
       ElseIf p = vbNo Then
       MsgBox "No se elimino el producto", 0, "Eliminar"
       End If
       End If
 listaprod.Close
 Text3.Text = ""
 Text2.Text = ""
 Text1.Text = ""
End Sub

Private Sub Command3_Click()
Form3.Hide
Form2.Show
End Sub

Private Sub Command4_Click()
Dim listaprod As DAO.Database
Dim Producto As DAO.Recordset
 Set listaprod = OpenDatabase(App.Path & "\listaprod.mdb")
 Set Producto = listaprod.OpenRecordset("Producto")
     Producto.Index = "Codigo"
     Producto.Seek "=", Text3.Text
    If Producto.NoMatch Then
     MsgBox "No existe este producto", vbExclamation, "Error en la modificación"
    Else
     z = MsgBox("¿Desea modificar el producto?", vbYesNo, "Confirmación")
     If z = vbYes Then
      Producto.Edit
      Producto.Fields("Nombre") = Text1.Text
      Producto.Fields("Precio") = Text2.Text
      Producto.Update
      MsgBox "El producto ha sido modificado el producto exitosamente", vbInformation, "Modificación"
     Else
       MsgBox "El producto no ha sido modificado", vbInformation, "Modificación"
    End If
    End If
End Sub

Private Sub Command5_Click()
Dim Producto As DAO.Recordset
 Set listaprod = OpenDatabase(App.Path & "\listaprod.mdb")
 Set Producto = listaprod.OpenRecordset("Producto")
     Producto.Index = "Codigo"
     Producto.Seek "=", Text3.Text
     If Producto.NoMatch Then
     MsgBox ("No está registrado este producto"), vbExclamation, ("aviso")
     Else
     Text3.Text = Producto.Fields("Codigo")
     Text2.Text = Producto.Fields("Precio")
     Text1.Text = Producto.Fields("Nombre")
     End If
     Text3.SetFocus
     Producto.Close
End Sub

Private Sub Form_Load()
If Text3.Text <> "" And Text2.Text <> "" And Text1.Text <> "" Then
  Command4.Enabled = True
  Command1.Enabled = True
Else
  Command4.Enabled = False
  Command1.Enabled = False
End If
End Sub

Private Sub Text1_Change()
If Text3.Text <> "" And Text2.Text <> "" And Text1.Text <> "" Then
  Command4.Enabled = True
  Command1.Enabled = True
Else
  Command4.Enabled = False
  Command1.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Text3.Text <> "" And Text2.Text <> "" And Text1.Text <> "" Then
  Command4.Enabled = True
  Command1.Enabled = True
Else
  Command4.Enabled = False
  Command1.Enabled = False
End If
End Sub

Private Sub Text3_Change()
If Text3.Text <> "" And Text2.Text <> "" And Text1.Text <> "" Then
  Command4.Enabled = True
  Command1.Enabled = True
Else
  Command4.Enabled = False
  Command1.Enabled = False
End If
End Sub
