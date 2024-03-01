VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form abm_COMP_COMPRA1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2175
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   17760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2175
   ScaleWidth      =   17760
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_envase 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12600
      MaxLength       =   5
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   17535
      Begin VB.TextBox t_unidad 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11160
         MaxLength       =   5
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox t_pusindto 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   7440
         MaxLength       =   8
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_dto 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10080
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_fechaultcompra 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6360
         MaxLength       =   8
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_precioultcompra 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   5400
         MaxLength       =   8
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_ref 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   4320
         MaxLength       =   8
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   3240
         MaxLength       =   8
         TabIndex        =   18
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox t_detalle 
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
         Left            =   1320
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   5295
      End
      Begin VB.TextBox t_basico 
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox t_importe 
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
         Left            =   15000
         MaxLength       =   21
         TabIndex        =   8
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox c_tasa 
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
         Left            =   13320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_pu 
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
         Left            =   8160
         MaxLength       =   21
         TabIndex        =   3
         Top             =   720
         Width           =   1815
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
         Left            =   6720
         MaxLength       =   21
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2040
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Envase"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   12360
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11040
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "% Dto."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10080
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   14880
         TabIndex        =   16
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Tasa Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   13320
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Pu"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8040
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Producto"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   240
         Width           =   5415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   17760
      _ExtentX        =   31327
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4410
            MinWidth        =   4410
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   11465
            MinWidth        =   11465
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "01/03/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:10 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "Unidad"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9360
      TabIndex        =   25
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "abm_COMP_COMPRA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub c_tasa_LostFocus()
If c_tasa.ListIndex < 0 Then
  c_tasa.ListIndex = 0
End If
End Sub

Private Sub Form_Activate()
t_basico.SetFocus
End Sub

Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2, g5 where "
If tipo = "I" Then
  q = q & "  [id_producto] = " & Val(t_basico)
Else
  q = q & "  [cod_barra] = '" & RTrim$(t_basico) & "'"
End If
q = q & " and a2.[id_unidad] = g5.[id_unidad]"
rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  t_ip = rs("id_producto")
  t_pu = rs("PRECIO_ULT_COMPRA")
  c_tasa.ListIndex = rs("cod_tasaiva")
  t_detalle.Enabled = False
  t_precioultcompra = rs("PRECIO_ULT_COMPRA")
  t_fechaultcompra = rs("fecha_ULT_COMPRA")
  t_unidad = rs("unidad")
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select
End Sub

Sub carga()
If IsNumeric(t_basico) Then
 If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_detalle.Enabled = True
    t_detalle.SetFocus
 Else
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
       Call busca("B") 'busca por cod. barra
    End If
 End If
Else
  Call busca("B") 'busca por cod. barra
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 8)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)

For i = 0 To 9
  c_tasa.AddItem para.tasaiva(i)
Next i
c_tasa.ListIndex = 0
End Sub


Private Sub t_basico_GotFocus()
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
Me.StatusBar1.Panels.item(2) = "[F8] Lista Precios - [1] Prod. s/ codificar - [Esc] Cancela"
para.producto_sel = 0
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If

End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga

End If

End Sub

Private Sub t_basico_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub

Sub cargarenglon(t As String)
  
  ip = t_ip
  If Len(t_detalle) = 0 Then
    d = " "
  Else
    d = t_detalle
  End If
  If Len(t_unidad) = 0 Then
    t_unidad = " "
  End If
  
  cu = Format$(Val(t_cantidad), "######0.00")
  ti = Format$(c_tasa, "####0.00")
  im = Format$(t_importe, "#####0.00")
  pu = Format$(t_pu, "#####0.00")
  tr = Format$(t_ref, "00000000")
  puc = Format$(t_precioultcompra, "#####0.00")
  ful = Format$(t_fechaultcompra, "dd/mm/yyyy")
  dto = Format$(t_dto, "###0.00")
  pusd = Format$(t_pusindto, "#####0.00")
  
  If t = "A" Then
    r = ABM_COMP_COMPRA.msf1.Rows
    ABM_COMP_COMPRA.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & pu & Chr(9) & ti & Chr$(9) & dto & Chr(9) & im & Chr(9) & tr & Chr(9) & puc & Chr(9) & ful & Chr(9) & pusd & Chr(9) & t_unidad & Chr(9) & Chr(9) & t_envase
  Else
    r = t_renglon
    ABM_COMP_COMPRA.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & cu & Chr(9) & pu & Chr(9) & ti & Chr(9) & dto & Chr(9) & im & Chr(9) & tr & Chr(9) & puc & Chr(9) & ful & Chr(9) & pusd & Chr(9) & t_unidad & Chr(9) & Chr(9) & t_envase, r
    ABM_COMP_COMPRA.msf1.RemoveItem r + 1
  End If
   
  s = 0
  v = 0
  For i = 1 To ABM_COMP_COMPRA.msf1.Rows - 1
      r = ABM_COMP_COMPRA.msf1.TextMatrix(i, 7)
      s = s + r
      v = v + (r * ABM_COMP_COMPRA.msf1.TextMatrix(i, 5) / 100)
  Next i
  ABM_COMP_COMPRA.t_subtotal = s
  ABM_COMP_COMPRA.t_iva = v
  ABM_COMP_COMPRA.sacatotales
  
End Sub
 
  
Sub limpia()
t_cantidad = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_basico = ""
t_unidad = ""
t_envase = ""
End Sub

Private Sub T_detalle_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Continua - [Esc] Cancela"

End Sub

Private Sub t_envase_LostFocus()
If Not IsNumeric(t_envase) Then
  t_envase = "1"
End If
End Sub

Private Sub t_importe_GotFocus()
t_pusindto = t_pu
If Val(t_dto) > 0 Then
    pu = (Val(t_pu) - (Val(t_pu) * Val(t_dto) / 100))
    t_pu = Format$(pu, "#####0.00")
End If
t_importe = Format$(Val(t_cantidad) * Val(t_pu), "#####0.00")
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If t_renglon = "" Then
   Call cargarenglon("A")
   Call limpia
   t_basico.SetFocus
  
  Else
   Call cargarenglon("M")
   Me.Hide
  End If
  
Else
  Call solonum(KeyAscii, 1)
End If
End Sub

