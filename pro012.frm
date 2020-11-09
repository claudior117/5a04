VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form pro_empaque 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ORDEN DE EMPAQUE"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   255
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   Begin VB.TextBox t_cantlineas 
      Height          =   285
      Left            =   7560
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox t_cl 
      Height          =   285
      Left            =   8520
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   9960
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      Begin VB.TextBox t_funcion 
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label12 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales del Comprobante"
      Height          =   735
      Left            =   240
      TabIndex        =   19
      Top             =   7320
      Width           =   9615
      Begin VB.TextBox t_observaciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Top             =   240
         Width           =   7695
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5175
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   11655
      Begin VB.ComboBox C_OBRA 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   8175
      End
      Begin VB.ComboBox c_sucursal 
         Height          =   315
         ItemData        =   "pro012.frx":0000
         Left            =   7440
         List            =   "pro012.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Height          =   255
         Left            =   7800
         Picture         =   "pro012.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox t_fechavto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10080
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_letra 
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
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1560
         Width           =   375
      End
      Begin VB.ComboBox c_tipocomp 
         Height          =   315
         ItemData        =   "pro012.frx":0376
         Left            =   1680
         List            =   "pro012.frx":0378
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_numcomp 
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
         Height          =   285
         Left            =   3120
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox t_sucursal 
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
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1560
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   720
         Width           =   6015
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Obra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Punto Venta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Ent:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8880
         TabIndex        =   24
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cliente"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   11
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "pro012.frx":037A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "pro012.frx":0BFC
         Style           =   1  'Graphical
         TabIndex        =   12
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
      TabIndex        =   10
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "03/06/2014"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "18:52"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "pro_empaque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Dim calcula_perc_ib As String
Dim alicuota_perc_ib As Single
Dim minimo_perc_ib As Double
Dim gcuit As String
Dim numint As Long
Dim cuentaact As Long
Dim abreviatura As String
Dim cantlineas As Integer
Dim ubicacionctacte As String


Sub iniciacomp()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra], [cant_lineas] from vta_06 where [sucursal] = " & Val(t_sucursal) & " and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex)
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  If rs("imprime_desc_extra") = "S" Then
    Check2 = 1
  Else
    Check2 = 0
  End If
  t_cantlineas = rs("cant_lineas")
Else
  Check2 = 0
  t_cantlineas = 25
End If
Set rs = Nothing

Call mensaje
End Sub
Sub limpia()
   Call armagrid
   t_subtotal = ""
   t_nograbado = ""
   t_perc = ""
   t_iva = ""
   t_total = ""
   Option1 = True
   
End Sub
Sub mensaje()
'activa mensaje de faturacion
tm = c_tipocomp & " [" & t_letra & "]"
If Option2 = True Then
  tm = tm & "  " & "CONTADO"
Else
  tm = tm & "  " & "CUENTA CORRIENTE"
End If
If c_prov.ListIndex = 0 Then
   tm = tm & " **" & vta_clientes.t_cli & "**"
Else
    tm = tm & " **" & c_prov & "**"
End If
Label20 = UCase$(tm)


End Sub
Sub carga()
  Set rs = New ADODB.Recordset
  q = "select [fecha], [fecha_vto], [cotizacion_dolar], [id_cliente], [num_int], [id_vendedor], [subtotal], [impuestos], [total], [perc_ib], [perc_gan], [perc_iva], [iva], " & _
  " [contado], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02] from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = " & c_tipocomp.ItemData(c_tipocomp.ListIndex) & " and [num_comp] = " & Val(t_numcomp)
  rs.MaxRecords = 1
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     t_fecha = rs("fecha")
     t_fechavto = rs("fecha_vto")
     t_cotizacion = rs("cotizacion_dolar")
     
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select [id_producto], [descripcion], [cantidad], [unidad], [pu], [tasaiva], [importe], [pu_final], [tasaib], [num_int], [renglon] from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("unidad") & Chr$(9) & Format$(rs1("pu"), "######0.00") & Chr(9) & rs1("tasaiva") & Chr(9) & rs1("importe") & Chr(9) & rs1("pu_final") & Chr(9) & rs1("tasaib")
        
        Set rs2 = New ADODB.Recordset
        q = "select [desc_ext], [cant_lineas] from vta_015 where [num_int] = " & rs1("num_int") & " and [renglon] = " & rs1("renglon")
        rs2.Open q, cn1
        If Not rs2.EOF And Not rs2.BOF Then
           k = rs2("cant_lineas")
           msf1.AddItem 0 & Chr(9) & "" & Chr(9) & rs2("desc_ext") & Chr(9) & k
           msf1.RowHeight(msf1.Rows - 1) = k * 250
 
        End If
        Set rs2 = Nothing
        rs1.MoveNext
     Wend
     Call renumera
     Set rs1 = Nothing
     
     
     
     
  
  
   
  
  Else
     EXISTE = "N"
  End If
  Set rs = Nothing
  
End Sub

Sub carga2()
  Set rs = New ADODB.Recordset
  q = "select [num_int], [letra], [sucursal], [id_tipocomp] from vta_02 where  [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  ni = 0
  While Not rs.EOF
    If rs("letra") = t_letra And rs("sucursal") = Val(t_sucursal) And rs("id_tipocomp") = 150 Then
      ni = rs("num_int")
      
    End If
    rs.MoveNext
  Wend
  Set rs = Nothing
  
 If ni <> 0 Then
     MsgBox ("Comprobante Existente")
     EXISTE = "S"
     Set rs = New ADODB.Recordset
     q = "select [fecha], [fecha_vto], [cotizacion_dolar], [id_cliente], [num_int], [id_vendedor], [subtotal], [impuestos], [total], [perc_ib], [perc_gan], [perc_iva], [iva], " & _
     " [contado], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [observaciones] from vta_02 where [num_int] = " & ni
     rs.MaxRecords = 1
     rs.Open q, cn1
  
     t_fecha = rs("fecha")
     t_fechavto = rs("fecha_vto")
     t_observaciones = rs("observaciones")
     
     c_prov.ListIndex = buscaindice(c_prov, rs("id_cliente"))
     
     Set rs1 = New ADODB.Recordset
     q = "select [id_producto], [descripcion], [cantidad], [unidad], [pu], [tasaiva], [importe], [pu_final], [tasaib], [num_int], [renglon], [bultos] from vta_03 where [num_int] = " & rs("num_int")
     rs1.Open q, cn1
     Call armagrid
     While Not rs1.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs1("id_producto"), "00000") & Chr(9) & rs1("descripcion") & Chr(9) & rs1("cantidad") & Chr(9) & rs1("unidad") & Chr$(9) & Format$(rs1("bultos"), "######0.00")
        
        Set rs2 = New ADODB.Recordset
        q = "select [desc_ext], [cant_lineas] from vta_015 where [num_int] = " & rs1("num_int") & " and [renglon] = " & rs1("renglon")
        rs2.Open q, cn1
        If Not rs2.EOF And Not rs2.BOF Then
           k = rs2("cant_lineas")
           msf1.AddItem 0 & Chr(9) & "" & Chr(9) & rs2("desc_ext") & Chr(9) & k
           msf1.RowHeight(msf1.Rows - 1) = k * 250
 
        End If
        Set rs2 = Nothing
        rs1.MoveNext
     Wend
     Call renumera
     Set rs1 = Nothing
     
     
     
   Set rs = Nothing
  
   
  
  Else
     EXISTE = "N"
  End If
  
  
End Sub
Private Sub btnacepta_Click()

      Call iniciagraba
  
End Sub


Sub iniciagraba()




J = MsgBox("Confirma Grabar Orden de Empaque ", 4)
If J = 6 Then
  If verificaperiodog(t_fecha) = "A" Then
        Call normal
   Else
        MsgBox ("Periodo Cerrado. Imposible grabar comprobante")
  End If

End If
  
    

End Sub





Sub normal()
  Set rs = New ADODB.Recordset
  q = "select * from vta_02 where [sucursal] = " & Val(t_sucursal) & " and letra = '" & t_letra & "' and [id_tipocomp] = 150 and [num_comp] = " & Val(t_numcomp)
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      EXISTE = "S"
         ni = rs("num_int")
         Set rs = Nothing
         J = MsgBox("Comprobante existente. ¿Desea Modificarlo? ", 4)
         If J = 6 Then
           Set cl_compvta = New comprobantes_venta
           cl_compvta.cargar2 (ni)
           cl_compvta.borrar
           Set cl_compvta = Nothing
           Call graba
         End If
       
  Else
    Set rs = Nothing
    EXISTE = "N"
    Call graba
  End If

End Sub
Private Sub btnsale_Click()
J = MsgBox("Abandona el comprobante (S/N)", 4)
If J = 6 Then
  Unload Me
End If
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 6
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 6700
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1000
msf1.ColWidth(5) = 1000

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Pieza"
msf1.TextMatrix(0, 2) = "Detalle Pieza"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "Nro.Bulto"

End Sub




Private Sub c_obra_LostFocus()
If C_OBRA.ListIndex < 0 Then
  C_OBRA.ListIndex = 0
End If

End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex >= 0 Then
  If Val(c_prov.ListIndex) > 1 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
   Else
    c_prov.ListIndex = 0
  End If
  Call iniciacli
End If

End Sub


Sub inicia()
espere.Show
espere.Label1 = "Inicializando Comprobante....."
espere.Refresh

   t_letra = "O"
   gcuit = vta_clientes.t_cuit
  
   Set cl_compvta = New comprobantes_venta
   cl_compvta.sucursal = Val(c_sucursal)
   cl_compvta.actual (150)
   cl_compvta.letra = t_letra
   cl_compvta.SACANUMCOMP
   t_numcomp = Format$(cl_compvta.numcomp, "00000000")
   cantlineas = cl_compvta.cant_lineas
   Set cl_compvta = Nothing
   t_cotizacion = para.cotizacion

     t_alicuotaib = "0.00"
     t_percib = "0.00"
     'gcuit = "0"
   
   Call armagrid
   Unload espere

  





End Sub

Private Sub c_sucursal_LostFocus()
If c_sucursal.ListIndex < 0 Then
  c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)
End If
t_sucursal = Format$(c_sucursal, "0000")
t_numcomp = ""
End Sub

Private Sub c_tipocomp_GotFocus()
btnacepta.Enabled = False
End Sub

Private Sub c_tipocomp_LostFocus()

Call iniciacli
End Sub









Private Sub Command5_Click()
vta_clientes.Show
End Sub



Sub iniciacli()
 If c_prov.ListIndex >= 0 Then
   vta_clientes.t_id = c_prov.ItemData(c_prov.ListIndex)
   vta_clientes.carga
   t_letra = "O"
   

 End If
 
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
  
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call TabEnter2(Me, 8)
End If


End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call carga_clientes(c_prov)
c_prov.ListIndex = 0
c_prov.RemoveItem 0
c_prov.ListIndex = 0

Call carga_obras(C_OBRA, "E")
C_OBRA.ListIndex = 0

Call carga_SUCURSALES(c_sucursal)
c_sucursal.ListIndex = buscaindice(c_sucursal, glo.sucursal)

c_tipocomp.clear
c_tipocomp.AddItem "Orden de Empaque", 0
c_tipocomp.ListIndex = 0

Call armagrid
Call barraesag(Me)
t_sucursal = Format$(glo.sucursal, "0000")
Load pro_empaque1

Load vta_clientes
vta_clientes.limpia
gcuit = "0"


End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload pro_empaque1

End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F3] Descipcion extra - [F5] Saca Renglon - [F9] Graba "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
Me.KeyPreview = False

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 If msf1.Rows > 1 Then
  If Val(msf1.TextMatrix(msf1.Rows - 1, 0)) > 0 Then
   Load gen_descextra
   gen_descextra.t_modulo = "E"
   gen_descextra.t_funcion = "A"
   gen_descextra.Show
  End If
 End If
End If




If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
   r = msf1.Row
   If r + 1 < msf1.Rows Then
      If Val(msf1.TextMatrix(r + 1, 0)) = 0 Then
        msf1.RemoveItem (r + 1)
      End If
   End If
   If msf1.Rows > 2 Then
     msf1.RemoveItem (r)
   Else
     Call armagrid
   End If
   Call renumera
  Else
   Call armagrid
   
 End If
 
End If


If KeyCode = vbKeyF9 Then
 
  Call renumera
 
  btnacepta.Enabled = True
  
  t_observaciones.SetFocus
End If

If KeyCode = vbKeyInsert Then
   pro_empaque1.t_renglon = ""
   pro_empaque1.t_cantidad = ""
   pro_empaque1.t_pu = ""
  
   If msf1.Rows - 1 < cantlineas Then
     pro_empaque1.Show
   Else
     MsgBox ("Se ha superado la cantidad maxima dde items para este comprobante")
   End If
End If


End Sub

Sub renumera()
r = 1
For i = 1 To msf1.Rows - 1
 If Val(msf1.TextMatrix(i, 0)) <> 0 Then
    msf1.TextMatrix(i, 0) = r
    r = r + 1
 End If
Next i


End Sub
Sub graba()
  'On Error GoTo ERRORGRABA
  
  numint = saca_ultnumero_int_comp("V")
      
  Set cl_compvta = New comprobantes_venta
  cl_compvta.sucursal = Val(t_sucursal)
  cl_compvta.actual (150)
  cl_compvta.letra = "O"
  cl_compvta.numcomp = Val(t_numcomp)
  abreviatura = cl_compvta.abreviatura
         cp = "0000-00000000"
         ep = "S"
         cp = "ctdo"
         contado = "S"
         ssi = 0
      
      If EXISTE = "N" Then
        cl_compvta.ACTUALIZA_NUMERADOR
      End If
      
      moneda = "P"
      
      
      
       codact = 0
       alicuotaib = 0
       cuentaact = para.cuenta_ventas
      
        
              
      tiporespiva = vta_clientes.c_iva.ItemData(vta_clientes.c_iva.ListIndex)
       
      If c_prov.ListIndex = 0 Then
        idcli = 1
      Else
        idcli = c_prov.ItemData(c_prov.ListIndex)
      End If
      
      
        T2 = 0
      
      cn1.BeginTrans
       
       
       QUERY = "INSERT INTO vta_02([num_int], [sucursal], [num_comp], [letra], [id_tipocomp], [id_cliente], [fecha], [id_usuario], [subtotal], [impuestos], [iva], [total]," & _
"[estado], [id_cuenta], [stock], [cta_cte], [grabado], [estado_pago], [recibo_Pago], [observaciones], [cotizacion_dolar], [total_otra_moneda], [moneda], [id_vendedor], " & _
" [VENTA], [CONTADO], [perc_ib], [perc_gan], [perc_iva] , [id_actividad], [alicuota_ib], [alicuota_perc_iva], [canje_cereal], [fecha_vto], [total_bultos],  [valor_declarado], " & _
" [transporte], [direccion_transp], [cuit_transp], [perc_ss], [sucursal_ingreso], [cliente02], [direccion02], [cuit02], [localidad02], [id_tipo_iva02], [chofer02], [dominio02], " & _
" [dominio_acoplado02], [SALDO_IMPAGO02], [num_z])"



QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numcomp) & ", '" & t_letra & "', 150" & _
", " & idcli & ", '" & t_fecha & "', " & para.id_usuario & ", " & Val(0) & ", " & Val(0) & ", " & Val(0) & ", 0" & _
", 'A', " & cuentaact & ", '" & cl_compvta.STOCK & "', '" & cl_compvta.ctacte & "', '" & cl_compvta.grabado & "', '" & ep & "', '" & cp & "', '" & t_observaciones & _
" ', 1, " & T2 & ", '" & moneda & "', 1, '" & cl_compvta.venta & "', '" & contado & "', " & Val(t_perc) & _
", 0, " & Val(t_perciva) & ", " & codact & ", " & Val(t_alicuotaib) & ", " & Val(t_alicuotaperciva) & ", false, '" & t_fechavto & "', 0, 0, ' ', ' ', ' ', 0, " & Val(c_sucursal) & _
", '" & Left$(vta_clientes.t_cli, 50) & "', '" & Left$(vta_clientes.t_direccion, 50) & "', '" & Left$(vta_clientes.t_cuit, 20) & "', '" & Left$(vta_clientes.t_localidad, 50) & _
"', " & tiporespiva & ", ' ', ' ', ' ', " & ssi & ", " & para.z_actual & ")"

                                                                                                                                                                                                                                                            
cn1.Execute QUERY
COSTOINV = 0
Set cl_cli = Nothing
For i = 1 To msf1.Rows - 1
  renglon = Val(msf1.TextMatrix(i, 0))
  If renglon > 0 Then
        
        If Val(msf1.TextMatrix(i, 1)) > 1 Then
          Set cl_prod = New productos
          cl_prod.cargar (Val(msf1.TextMatrix(i, 1)))
          costo = cl_prod.costoreal
          Set cl_prod = Nothing
        Else
          costo = 0
        End If
        
        QUERY = "INSERT INTO vta_03([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [pu], [importe], [tasaiva], [impuesto], [costo], [cantidad_original], [tunidad], [pu_final], [tasaib],[bultos])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & " ', " & Val(msf1.TextMatrix(i, 3)) & ", 0,0,0,0,0, " & Val(msf1.TextMatrix(i, 3)) & ", '" & msf1.TextMatrix(i, 4) & "', 0, 0, " & Val(msf1.TextMatrix(i, 5)) & ")"
        cn1.Execute QUERY

  Else
    'grabo desc extra
    QUERY = "INSERT INTO vta_015([num_int], [RENGLON], [desc_ext], [cant_lineas])"
    QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i - 1, 0)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 3)) & ")"
    cn1.Execute QUERY
  End If


Next i
      
      
     
      
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Emitir Orden de Empaque NI:" & numint & "', " & para.id_usuario & ", 'V', " & numint & ", '" & Now & "', '[150] " & t_letra & " " & Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000") & "', 12, " & idcli & ")"
  
     cn1.Execute QUERY

      
      
      cn1.CommitTrans
      
      
     
        
      
      
          J = MsgBox("Confirma Impresion del Comprobante", 4)
          If J = 6 Then
             Set cl_compvta = New comprobantes_venta
             cl_compvta.cargar2 (numint)
             cl_compvta.imprimir
          End If
      Call INICIALIZA2(Me)
      Call armagrid
      
      
      
      
      c_prov.SetFocus
      t_sucursal = Format$(c_sucursal, "0000")
      
      
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos y vuelva a repetir la operacion")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
     pro_empaque1.t_renglon = msf1.Row
     pro_empaque1.t_basico = msf1.TextMatrix(msf1.Row, 1)
     pro_empaque1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
     pro_empaque1.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
     pro_empaque1.t_unidad = msf1.TextMatrix(msf1.Row, 4)
      pro_empaque1.t_pu = msf1.TextMatrix(msf1.Row, 5)
     
     pro_empaque1.Show
   Else
     Load gen_descextra
     gen_descextra.Text1 = msf1.TextMatrix(msf1.Row, 2)
     gen_descextra.t_modulo = "E"
     gen_descextra.t_funcion = "M"
     gen_descextra.Show
   End If
  End If
End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
Me.KeyPreview = True

End Sub



Private Sub t_fecha_GotFocus()
If glo.sucursalf = Val(t_sucursal) Then
   t_fecha = Format$(Now, "dd/mm/yyyy")
   t_fecha.Locked = True
Else
   t_fecha.Locked = False
End If

End Sub

Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub


Private Sub t_fechavto_LostFocus()
If Not IsDate(t_fechavto) Then
  t_fechavto = Format$(Now, "dd/mm/yyyy")
Else
  t_fechavto = Format$(t_fechavto, "dd/mm/yyyy")
End If

End Sub



Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)

End Sub

Private Sub t_numcomp_LostFocus()
If IsNumeric(t_numcomp) Then
   t_numcomp = Format$(t_numcomp, "00000000")
   
    Call carga2
     
   Call iniciacomp

Else
  t_numcomp.SetFocus
End If
End Sub

Private Sub t_observaciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 btnacepta.SetFocus
 End If
 
 
End Sub

Private Sub t_observaciones_LostFocus()
Call NULOS(t_observaciones)
End Sub


Private Sub t_sucursal_GotFocus()
t_sucursal = Format$(Val(c_sucursal), "0000")
End Sub

Private Sub t_sucursal_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If
Call inicia
End Sub



