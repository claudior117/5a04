VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form fsc_tique1 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO DE ARTICULOS"
   ClientHeight    =   2355
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   15915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   15915
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_tasaib 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   405
      Left            =   10800
      MaxLength       =   8
      TabIndex        =   20
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   15735
      Begin VB.TextBox t_tipo 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   4200
         MaxLength       =   5
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
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
         Left            =   9720
         MaxLength       =   5
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Left            =   120
         Picture         =   "fsc001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox t_ip 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   3120
         MaxLength       =   5
         TabIndex        =   16
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
         Left            =   2760
         MaxLength       =   49
         TabIndex        =   1
         Top             =   720
         Width           =   5415
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
         Left            =   720
         MaxLength       =   20
         TabIndex        =   0
         Top             =   720
         Width           =   1935
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
         Left            =   13920
         MaxLength       =   11
         TabIndex        =   6
         Top             =   720
         Width           =   1695
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
         Left            =   12480
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1215
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
         Left            =   10800
         MaxLength       =   10
         TabIndex        =   4
         Top             =   720
         Width           =   1455
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
         Left            =   8280
         MaxLength       =   8
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_renglon 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Unidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   9600
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Basico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   13920
         TabIndex        =   14
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Tasa Iva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   12360
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Pu"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   10800
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   8160
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   240
         Width           =   5415
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1980
      Width           =   15915
      _ExtentX        =   28072
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   194028
            MinWidth        =   194028
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "05/12/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:48 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "fsc_tique1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim Fiscaltq As Driver



Private Sub c_tasa_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta "
End Sub

Private Sub c_tasa_LostFocus()
If c_tasa.ListIndex < 0 Then
  c_tasa.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
ABM_PROD.Show
End Sub

Private Sub Form_Activate()

t_basico.SetFocus

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
    Call TabEnter2(Me, 6)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()


For i = 0 To 9
  c_tasa.AddItem para.tasaiva(i)
Next i
c_tasa.ListIndex = 0

Set cl_fiscal = New fiscal
cl_fiscal.carga (glo.sucursalf)

End Sub


Private Sub t_basico_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER]Acepta - [ESC]Sale - [F3]Cantidad - [F6]Dto1 - [F7]Dto2 - [F8]Lista Precios"

t_detalle.Enabled = False
If para.producto_sel > 0 Then
  t_basico = para.producto_sel
End If
End Sub

Private Sub t_basico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
  vta_listaprecios.Show
End If

If KeyCode = vbKeyF3 Then
  u = InputBox$("Cantidad a facturar", "Tique Fiscal", 1)
  If Val(u) <= 0 Then
    u = 0
  End If
  t_cantidad = Format$(u, "####0.00")
End If


If KeyCode = vbKeyF6 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d1 = rs("descuento1")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d1, "##0.00") & "%"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(fsc_tique.T_TOTAL) * d1 / 100, "######0.00")
  t_importe = t_pu
  
End If

If KeyCode = vbKeyF7 And t_renglon = "" Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  d2 = rs("descuento2")
  Set rs = Nothing
  t_basico = 1
  t_detalle = "Descuento " & Format$(d2, "##0.00") & "%"
  c_tasa.ListIndex = 1
  t_ip = 1
  t_cantidad = "1.00"
  t_pu = Format$(-Val(fsc_tique.T_TOTAL) * d2 / 100, "######0.00")
  t_importe = t_pu

End If

End Sub

Private Sub t_basico_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call carga
End If
End Sub

Sub carga()
If t_basico = "" Then
  t_basico = "1"
End If

If IsNumeric(t_basico) Then
 If Val(t_basico) <= 1 Then
    t_basico = 1
    t_ip = 1
    t_detalle.Enabled = True
    c_tasa.ListIndex = buscaindice2(c_tasa, para.tasageneral)
    t_detalle.SetFocus
    t_tasaib = para.tasaib
 Else
    If Len(t_basico) <= 5 Then
       Call busca("I") 'busca por id. producto
    Else
     If Mid$(t_basico, 1, 2) <> "21" Then
       Call busca("B") 'busca por cod. barra
     Else
       Call busca("P") 'cod. barra interno que lleva el precio
     End If
    End If
    
 End If
End If
End Sub
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2, g5, g12 where a2.[id_unidad] = g5.[id_unidad] and a2.[id_tasaib] = g12.[id_tasaib]  "

Select Case tipo
Case Is = "I"
  q = q & " and [id_producto] = " & Val(t_basico)
Case Is = "B"
  q = q & " and [cod_barra] = '" & RTrim$(t_basico) & "'"
Case Is = "P"
  cp = Val(Mid$(t_basico, 3, 2))
  q = q & " and [id_producto] = " & cp
End Select


rs.MaxRecords = 1
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
  t_detalle = rs("descripcion")
  c_tasa.ListIndex = rs("cod_tasaiva")
  t_unidad = rs("unidad")
  t_tasaib = rs("tasaib")

  If tipo <> "P" Then
    If para.tipoprecioventa = 1 Then
        t_pu = rs("precio_final")
    Else
        t_pu = rs("pu")
    End If
    t_ip = rs("id_producto")
  Else
    t_pu = Format$(Val(Mid$(t_basico, 5, 6) & "." & Mid$(t_basico, 11, 2)), "#######0.00")
    t_ip = Format$(cp, "00000")
   
  End If
    
  
  
  
  If rs("tipo_carga_tique") = "A" Then
    If Val(t_cantidad) <= 0 Then
      t_cantidad = 1
    End If
    t_importe = Format$(Val(t_cantidad) * Val(t_pu), "#####0.00")
    
     If Val(fsc_tique.T_TOTAL) + Val(t_importe) >= Val(fsc_tique.t_limite) Then 'cl_fiscal.limitetique Then
       MsgBox ("El importe del comprobante supera el limite establecido para la impresora. Cierre el tique actual y abra uno nuevo para seguir cargando")
     Else
       Call fsc_tique.cargarenglon2("A")
     End If
     Call limpia
     t_basico.SetFocus
  End If
  
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub

Private Sub t_basico_LostFocus()
Call barraesag(Me)
End Sub

Private Sub t_cantidad_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER]Acepta"
End Sub

Private Sub t_cantidad_KeyPress(KeyAscii As Integer)
   Call solonum(KeyAscii, 1)
End Sub


  
Sub limpia()
t_cantidad = ""
t_basico = ""
t_detalle = ""
t_pu = ""
t_importe = ""
t_ip = ""
t_unidad = ""
End Sub

Private Sub T_detalle_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta - [ESC] Sale"
End Sub

Private Sub t_importe_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta "
t_importe = Format$(Val(t_cantidad) * Val(t_pu), "#####0.00")
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If t_renglon = "" Then
   If Val(fsc_tique.T_TOTAL) + Val(t_importe) >= Val(fsc_tique.t_limite) Then 'cl_fiscal.limitetique Then
     MsgBox ("El importe del comprobante supera el limite establecido para la impresora. Cierre el tique actual y abra uno nuevo para seguir cargando")
   Else
     Call fsc_tique.cargarenglon2("A")
   End If
  End If
  Call limpia
  t_basico.SetFocus
Else
  Call solonum(KeyAscii, 1)
End If
End Sub


Private Sub t_pu_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER] Acepta - [F6]Dto % - [F7]Dto $"

End Sub

Private Sub t_pu_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF6 Then
  d = InputBox("Ingrese % descuento", "Descuento")
  If Val(d) > 0 Then
     pd = Format(Val(t_pu) * Val(d) / 100, "######0.00")
     t_pu = Val(t_pu) - pd
  End If
End If
  



If KeyCode = vbKeyF7 Then
  d = InputBox("Ingrese descuento en pesos", "Descuento $")
  If Val(d) > 0 Then
     pd = Format(Val(d), "######0.00")
     t_pu = Val(t_pu) - pd
  End If
End If
  
End Sub

Private Sub t_unidad_GotFocus()
Me.StatusBar1.Panels.item(2) = "[ENTER]Acepta"
End Sub
