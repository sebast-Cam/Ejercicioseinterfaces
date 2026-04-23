VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000004&
   Caption         =   "Form2"
   ClientHeight    =   6600
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9780
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   6600
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de facturación"
      Height          =   2895
      Left            =   4800
      TabIndex        =   5
      Top             =   240
      Width           =   3975
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura"
         Height          =   1695
         Left            =   1200
         MaskColor       =   &H00C0FFFF&
         Picture         =   "Form2.frx":2524D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir"
      Height          =   2415
      Left            =   5520
      Picture         =   "Form2.frx":27817
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Administración"
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Registro de empleados"
         Height          =   1815
         Left            =   960
         Picture         =   "Form2.frx":27B21
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Proveedores"
      Height          =   2895
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Registro de Productos"
         Height          =   1695
         Left            =   960
         MaskColor       =   &H00C0C0FF&
         Picture         =   "Form2.frx":27E2B
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Menu sp 
      Caption         =   "Sistema de proveedores"
      Begin VB.Menu rp 
         Caption         =   "Registro de productos"
      End
   End
   Begin VB.Menu admin 
      Caption         =   "Administración"
      Begin VB.Menu re 
         Caption         =   "Registro de empleados"
      End
   End
   Begin VB.Menu sisfa 
      Caption         =   "Sistema de facturación"
      Begin VB.Menu fact 
         Caption         =   "Factura"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form2.Hide
Form3.Show
End Sub

Private Sub Command3_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
Form2.Hide
Form5.Show
End Sub

Private Sub fact_Click()
Form2.Hide
Form5.Show
End Sub

Private Sub re_Click()
Form2.Hide
Form4.Show
End Sub

Private Sub rp_Click()
Form2.Hide
Form3.Show
End Sub
