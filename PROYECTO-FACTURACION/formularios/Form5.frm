VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10635
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   8505
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text24 
      Height          =   285
      Left            =   5760
      TabIndex        =   37
      Top             =   3960
      Width           =   3495
   End
   Begin VB.TextBox Text23 
      Height          =   285
      Left            =   1080
      TabIndex        =   36
      Top             =   3960
      Width           =   4695
   End
   Begin VB.TextBox Text22 
      Height          =   285
      Left            =   5760
      TabIndex        =   35
      Top             =   3720
      Width           =   3495
   End
   Begin VB.TextBox Text21 
      Height          =   285
      Left            =   1080
      TabIndex        =   34
      Top             =   3720
      Width           =   4695
   End
   Begin VB.TextBox Text20 
      Height          =   285
      Left            =   5760
      TabIndex        =   33
      Top             =   3480
      Width           =   3495
   End
   Begin VB.TextBox Text19 
      Height          =   285
      Left            =   1080
      TabIndex        =   32
      Top             =   3480
      Width           =   4695
   End
   Begin VB.TextBox Text18 
      Height          =   285
      Left            =   5760
      TabIndex        =   31
      Top             =   3240
      Width           =   3495
   End
   Begin VB.TextBox Text17 
      Height          =   285
      Left            =   1080
      TabIndex        =   30
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   5760
      TabIndex        =   29
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   1080
      TabIndex        =   28
      Top             =   3000
      Width           =   4695
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   5760
      TabIndex        =   27
      Top             =   2760
      Width           =   3495
   End
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   5760
      TabIndex        =   25
      Top             =   2520
      Width           =   3495
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   1080
      TabIndex        =   24
      Top             =   2520
      Width           =   4695
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   5760
      TabIndex        =   23
      Top             =   2280
      Width           =   3495
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   1080
      TabIndex        =   22
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   5760
      TabIndex        =   21
      Top             =   2040
      Width           =   3495
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1080
      TabIndex        =   20
      Top             =   2040
      Width           =   4695
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5760
      TabIndex        =   19
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1080
      TabIndex        =   18
      Top             =   1800
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8520
      TabIndex        =   17
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sumar"
      Height          =   975
      Left            =   840
      TabIndex        =   16
      Top             =   7200
      Width           =   4335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Volver"
      Height          =   855
      Left            =   8040
      TabIndex        =   12
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar"
      Height          =   855
      Left            =   3240
      TabIndex        =   11
      Top             =   6000
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7200
      TabIndex        =   10
      Text            =   "Cantidad"
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5040
      TabIndex        =   5
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Buscar"
      Height          =   855
      Left            =   5640
      TabIndex        =   2
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Registrar en caja"
      Height          =   855
      Left            =   840
      TabIndex        =   1
      Top             =   6000
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Precios:"
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Productos:"
      Height          =   495
      Left            =   1080
      TabIndex        =   14
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Precio total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8640
      TabIndex        =   13
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000E&
      Caption         =   "Cantidad"
      Height          =   375
      Left            =   7200
      TabIndex        =   9
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Precio"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      Caption         =   "Producto"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Código"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "FACTURA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Combo1.Text = "Cantidad"
End Sub

Private Sub Command1_Click()
If Combo1.Text = "Cantidad" Then
MsgBox "Coloque la cantidad del producto que desea vender", vbExclamation, "Aviso"
Else
  If Text5.Text = "" And Text6.Text = "" Then
     Text5.Text = Text2.Text
     Text6.Text = Val(Text3.Text) * Val(Combo1.Text)
  Else
     If Text7.Text = "" And Text8.Text = "" Then
        Text7.Text = Text2.Text
        Text8.Text = Val(Text3.Text) * Val(Combo1.Text)
     Else
        If Text9.Text = "" And Text10.Text = "" Then
           Text9.Text = Text2.Text
           Text10.Text = Val(Text3.Text) * Val(Combo1.Text)
        Else
           If Text11.Text = "" And Text12.Text = "" Then
              Text11.Text = Text2.Text
              Text12.Text = Val(Text3.Text) * Val(Combo1.Text)
           Else
              If Text13.Text = "" And Text14.Text = "" Then
                 Text13.Text = Text2.Text
                 Text14.Text = Val(Text3.Text) * Val(Combo1.Text)
              Else
                 If Text15.Text = "" And Text16.Text = "" Then
                    Text15.Text = Text2.Text
                    Text16.Text = Val(Text3.Text) * Val(Combo1.Text)
                 Else
                    If Text17.Text = "" And Text18.Text = "" Then
                       Text17.Text = Text2.Text
                       Text18.Text = Val(Text3.Text) * Val(Combo1.Text)
                    Else
                       If Text19.Text = "" And Text20.Text = "" Then
                          Text19.Text = Text2.Text
                          Text20.Text = Val(Text3.Text) * Val(Combo1.Text)
                       Else
                          If Text21.Text = "" And Text22.Text = "" Then
                            Text21.Text = Text2.Text
                            Text22.Text = Val(Text3.Text) * Val(Combo1.Text)
                          Else
                             If Text23.Text = "" And Text24.Text = "" Then
                                Text23.Text = Text2.Text
                                Text24.Text = Val(Text3.Text) * Val(Combo1.Text)
                             Else
                             MsgBox "La factura esta llena", vbExclamation, "Factura"
                             End If
                           End If
                        End If
                    End If
                End If
             End If
        End If
     End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
End If
End If
End If
End Sub

Private Sub Command2_Click()
Dim Producto As DAO.Recordset
 Set listaprod = OpenDatabase(App.Path & "\listaprod.mdb")
 Set Producto = listaprod.OpenRecordset("Producto")
     Producto.Index = "Codigo"
     Producto.Seek "=", Text1.Text
     If Producto.NoMatch Then
     MsgBox ("No está registrado este producto"), vbExclamation, ("aviso")
     Else
     Text1.Text = Producto.Fields("Codigo")
     Text3.Text = Producto.Fields("Precio")
     Text2.Text = Producto.Fields("Nombre")
     End If
     Text1.SetFocus
     Producto.Close
End Sub

Private Sub Command3_Click()
Text4.Text = erased
Text5.Text = erased
Text6.Text = erased
Text7.Text = erased
Text8.Text = erased
Text9.Text = erased
Text10.Text = erased
Text11.Text = erased
Text12.Text = erased
Text13.Text = erased
Text14.Text = erased
Text15.Text = erased
Text16.Text = erased
Text17.Text = erased
Text18.Text = erased
Text19.Text = erased
Text20.Text = erased
Text21.Text = erased
Text22.Text = erased
Text23.Text = erased
Text24.Text = erased
End Sub

Private Sub Command4_Click()
Form5.Hide
Form2.Show
End Sub

Private Sub Command5_Click()
Text4.Text = Val(Text2.Text) + Val(Text4.Text) + Val(Text6.Text) + Val(Text8.Text) + Val(Text10.Text) + Val(Text12.Text) + Val(Text14.Text) + Val(Text16.Text) + Val(Text18.Text) + Val(Text20.Text)
End Sub

Private Sub Form_Load()
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.AddItem "6"
Combo1.AddItem "7"
Combo1.AddItem "8"
Combo1.AddItem "9"
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
  Command1.Enabled = True
Else
  Command1.Enabled = False
End If
End Sub


Private Sub Text1_Change()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
  Command1.Enabled = True
Else
  Command1.Enabled = False
End If
End Sub

Private Sub Text2_Change()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
  Command1.Enabled = True
Else
  Command1.Enabled = False
End If
End Sub

Private Sub Text3_Change()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
  Command1.Enabled = True
Else
  Command1.Enabled = False
End If
End Sub

