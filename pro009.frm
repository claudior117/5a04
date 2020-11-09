VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form pro_estructura2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "INGRESA PIEZA A ESTRUCTURA"
   ClientHeight    =   2175
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox c_pieza 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         TabIndex        =   0
         Text            =   "Combo1"
         Top             =   720
         Width           =   8775
      End
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10440
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6840
         MaxLength       =   5
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_cantidad 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9240
         MaxLength       =   8
         TabIndex        =   1
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   9000
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10320
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000C000&
         Caption         =   "Pieza"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   9135
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1920
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8820
            MinWidth        =   8820
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "22/02/2011"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:10 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "pro_estructura2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984




Private Sub c_pieza_LostFocus()
If c_pieza.ListIndex < 0 Then
  MsgBox ("Error al seleccionar Pieza")
  c_pieza.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()

c_pieza.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  Case Is = 27
        Me.Hide
End Select
End Sub


Private Sub Form_Load()
Call carga_piezas(c_pieza)
c_pieza.ListIndex = 0
End Sub

Private Sub t_cantidad_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F5] Comvierte Unidades x Envase"
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Call solonum(KeyAscii, 1)

End If
End Sub

Sub cargarenglon(t As String)
  d = c_pieza
  cu = Format$(Val(t_cantidad), "######0.00")
  u = Left$(t_unidad, 8)
  ip = c_pieza.ItemData(c_pieza.ListIndex)
  If ip <> pro_estructura.c_prov.ItemData(pro_estructura.c_prov.ListIndex) Then
   If t = "A" Then
    r = pro_estructura.msf1.Rows
    pro_estructura.msf1.AddItem r & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr(9) & ip & Chr(9) & "P"
   Else
    r = t_renglon
    pro_estructura.msf1.AddItem r & Chr(9) & d & Chr(9) & cu & Chr(9) & u & Chr(9) & ip & Chr(9) & "P", r
    pro_estructura.msf1.RemoveItem r + 1
   End If
   para.producto_sel = 0
 Else
   MsgBox ("Imposible agregar una pieza dentro de si misma")
   para.producto_sel = 0
 End If
End Sub
 
  
Sub limpia()
t_cantidad = ""
t_detalle = ""
t_unidad = ""
t_renglon = ""
End Sub

Sub pasa()
 If t_renglon = "" Then
   Call cargarenglon("A")
   c_pieza.SetFocus
   
  Else
   Call cargarenglon("M")
   Me.Hide
  End If
  Call limpia
  
End Sub

Private Sub t_unidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
     Call pasa
End If

End Sub
