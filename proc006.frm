VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form op2 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ORDEN DE PAGO EN DOLARES"
   ClientHeight    =   4770
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4770
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   3135
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   5895
      Begin VB.TextBox t_retib 
         Height          =   285
         Left            =   2280
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox t_cotiz 
         Height          =   285
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox t_total 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox t_retgan 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox t_op 
         Height          =   315
         Left            =   2280
         TabIndex        =   0
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF0000&
         Caption         =   "Ret. I.B. U$s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000C0&
         Caption         =   "Cotizacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "TOTAL U$s en Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Ret. Gan. U$s:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Total U$s por O.P."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   4560
      TabIndex        =   5
      Top             =   3360
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "proc006.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "proc006.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   4515
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "op2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub btnacepta_Click()
Me.Hide
End Sub

Private Sub btnsale_Click()
 Me.Hide
End Sub


Sub limpia()
 t_op = " "
 t_retgan = " "
 t_retib = " "
 T_TOTAL = " "
 t_cotiz = " "
 
 End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF12 Then
  gen_tools.Show
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 2)
End If

End Sub



Private Sub t_op_LostFocus()
Call totales
End Sub
Sub totales()
t_op = Format$(Val(t_op), "######0.00")
t_retgan = Format$(Val(t_retgan), "######0.00")
t_retib = Format$(Val(t_retib), "######0.00")
T_TOTAL = Format$(Val(t_op) + Val(t_retgan) + Val(t_retib), "######0.00")
op.t_totald = t_op

End Sub

Private Sub t_retgan_LostFocus()
Call totales
End Sub

Private Sub t_retib_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
End If
End Sub

Private Sub t_retib_LostFocus()
Call totales
End Sub
