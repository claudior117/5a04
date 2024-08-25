VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form ABM_OC 
   BackColor       =   &H00E0E0E0&
   Caption         =   "EMITIR ORDEN DE COMPRA"
   ClientHeight    =   8715
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   12030
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   9000
      TabIndex        =   43
      Top             =   6360
      Width           =   2775
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprime Descripcion Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar"
      Height          =   855
      Left            =   10800
      TabIndex        =   41
      Top             =   1080
      Width           =   1095
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   120
         Picture         =   "Proc001A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Totales"
      Height          =   855
      Left            =   240
      TabIndex        =   34
      Top             =   7320
      Width           =   8415
      Begin VB.TextBox t_subtotal 
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
         Left            =   120
         MaxLength       =   14
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox t_nograbado 
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
         Left            =   1680
         MaxLength       =   14
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox t_iva 
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
         Left            =   3360
         MaxLength       =   14
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox t_total 
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
         Left            =   5040
         MaxLength       =   14
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox t_dolares 
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
         Left            =   6840
         MaxLength       =   14
         TabIndex        =   35
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "No Grabado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3360
         TabIndex        =   38
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5040
         TabIndex        =   37
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00400040&
         Caption         =   "Total U$s"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   10800
      TabIndex        =   30
      Top             =   0
      Width           =   975
      Begin VB.OptionButton Option3 
         BackColor       =   &H8000000A&
         Caption         =   "U$s"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H8000000A&
         Caption         =   "$"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   240
      TabIndex        =   25
      Top             =   6360
      Width           =   8415
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   9
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox t_condiciones 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   8
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Condiciones de Compra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   2040
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7646
      _Version        =   393216
      WordWrap        =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1935
      Left            =   240
      TabIndex        =   19
      Top             =   0
      Width           =   10455
      Begin VB.TextBox t_cotiz 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   9360
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Left            =   7200
         Picture         =   "Proc001A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox t_tecontacto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   7560
         MaxLength       =   25
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox t_contacto 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox t_fechaprob 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6000
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_numoc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox t_sucursal 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Text            =   "c_prov"
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cotiz. Dolar:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   33
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Te:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6840
         TabIndex        =   28
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Contacto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Probable Entrega::"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Orden de Compra Nro.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   16
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Proc001A.frx":040F
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Proc001A.frx":0C91
         Style           =   1  'Graphical
         TabIndex        =   17
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
      TabIndex        =   15
      Top             =   8460
      Width           =   12030
      _ExtentX        =   21220
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14111
            MinWidth        =   14111
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "25/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:50 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ABM_OC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Sub busca(tipo As String)
'tipo = I por id_producto tipo = B por cod_barra
Set rs = New ADODB.Recordset
q = "select * from a2 where "
If tipo = "I" Then
  q = q & "  [id_producto] = " & Val(t_basico)
Else
  q = q & "  [cod_barra] = " & Val(t_basico)
End If
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
Else
  MsgBox ("Producto no Ingresado")
  t_basico.SetFocus
End If
Set rs = Nothing
End Sub


Sub carga_oc()
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar(65, "O", Val(t_sucursal), Val(t_numoc), 0)
  If cl_comp.numint = 0 Then
    EXISTE = "N"
    t_cotiz = para.cotizacion
    
  Else
     EXISTE = "S"
     MsgBox ("La Orden de Compra ya existe en el Sistema")
     Set rs = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & cl_comp.numint
     t_fecha = cl_comp.fecha
     t_fechaprob = cl_comp.fechaprobentrega
     c_prov.ListIndex = buscaindice(c_prov, cl_comp.idproveedor)
     rs.Open q, cn1
     While Not rs.EOF
        r = msf1.Rows
        msf1.AddItem r & Chr(9) & Format$(rs("id_producto"), "00000") & Chr(9) & rs("detalle") & Chr(9) & rs("id_requisicion") & Chr(9) & rs("observaciones") & Chr(9) & rs("cantidad") & Chr(9) & rs("pu") & Chr(9) & rs("unidad") & Chr(9) & rs("tasa_iva") & Chr(9) & rs("importe") & Chr(9) & "" & Chr(9) & rs("id_obra") & Chr(9) & rs("ESTADO") & Chr(9) & de
        r = r + 1
        Set rs2 = New ADODB.Recordset
        q = "select * from a21 where [num_int] = " & rs("num_int") & " and [renglon] = " & rs("renglon")
        rs2.Open q, cn1
        If Not rs2.EOF And Not rs2.BOF Then
          k = rs2("cant_lineas")
          msf1.AddItem 0 & Chr(9) & "" & Chr(9) & rs2("descripcion") & Chr(9) & k
          msf1.RowHeight(msf1.Rows - 1) = k * 250
        End If
        Set rs2 = Nothing
       
       rs.MoveNext
     Wend
     Set rs = Nothing
     Call renumera
  End If
End Sub

Private Sub btnacepta_Click()
 If msf1.Rows > 1 Then
  J = MsgBox("Confirma Grabar Orden de Compra", 4)
  If J = 6 Then
   If verificaperiodog(t_fecha) = "A" Then
     Call graba
   Else
     MsgBox ("El periodo para el cual desea ingresar el comprobante esta CERRADO!!!!")
  End If
 End If
End If
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 14
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 4000
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 2500
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 1000
msf1.ColWidth(7) = 600
msf1.ColWidth(8) = 800
msf1.ColWidth(9) = 1100
msf1.ColWidth(10) = 2000
msf1.ColWidth(11) = 500
msf1.ColWidth(12) = 500
msf1.ColWidth(13) = 5000


msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Referencia"
msf1.TextMatrix(0, 4) = "Observaciones"
msf1.TextMatrix(0, 5) = "Cantidad"
msf1.TextMatrix(0, 6) = "P.U."
msf1.TextMatrix(0, 7) = "Unid."
msf1.TextMatrix(0, 8) = "% Iva"
msf1.TextMatrix(0, 9) = "Importe"
msf1.TextMatrix(0, 10) = "Obra/Destino"
msf1.TextMatrix(0, 11) = "Id.Obra"
msf1.TextMatrix(0, 12) = ""
msf1.TextMatrix(0, 13) = "Descirpcion Extra "
msf1.ColWidth(11) = 0
msf1.ColWidth(12) = 0


End Sub





Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  If Val(c_prov) > 0 Then
    c_prov.ListIndex = buscaindice(c_prov, Val(c_prov))
  Else
    c_prov.ListIndex = 0
  End If
End If

Set cl_prov = New proveedores
cl_prov.carga (c_prov.ItemData(c_prov.ListIndex))
If cl_prov.idprov > 0 Then
   t_contacto = cl_prov.contacto
   t_tecontacto = cl_prov.tecontacto
End If
Set cl_prov = Nothing
End Sub



Private Sub Check2_LostFocus()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra] from g2 where [id_tipocomp] = 65"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
If Not rs.EOF And Not rs.BOF Then
 If Check2 = 1 Then
  rs("imprime_desc_extra") = "S"
 Else
   rs("imprime_desc_extra") = "N"
 End If
 rs.Update
End If
Set rs = Nothing
End Sub

Private Sub Command1_Click()
gen_seleccionarimp.Show
End Sub

Private Sub Command2_Click()
ABM_PROv.Show
End Sub

Private Sub Command2_LostFocus()
c_prov.clear
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
    gen_tools.Show
    
End Select
End Sub
Sub iniciacomp()
Set rs = New ADODB.Recordset
q = "select [imprime_desc_extra] from g2 where  [id_tipo_comp] = 65"
rs.Open q, cn1
If Not rs.EOF And Not rs.BOF Then
  If rs("imprime_desc_extra") = "S" Then
    Check2 = 1
  Else
    Check2 = 0
  End If
Else
  Check2 = 0
End If
Set rs = Nothing
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
Sub sacatotales()
t_subtotal = Format$(Val(t_subtotal), "######0.00")
t_nograbado = Format$(Val(t_nograbado), "######0.00")
t_iva = Format$(Val(t_iva), "######0.00")
T_TOTAL = Format$(Val(t_subtotal) + Val(t_nograbado) + Val(t_perc) + Val(t_iva), "######0.00")
If Option4 = True Then
 If Val(t_cotiz) <= 1 Then
  t_cotiz = para.cotizacion
 End If
  t_dolares = Format$(Val(T_TOTAL) / Val(t_cotiz), "#####0.00")
Else
  t_dolares = Format$(Val(T_TOTAL) * Val(t_cotiz), "#####0.00")
End If
End Sub

Sub sacatotales2()
  s = 0
  v = 0
  For i = 1 To msf1.Rows - 1
      r = Val(msf1.TextMatrix(i, 9))
      s = s + r
      v = v + (r * Val(msf1.TextMatrix(i, 8)) / 100)
  Next i
  t_subtotal = s
  t_iva = v
  Call sacatotales

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 13)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
t_sucursal = Format$(glo.sucursal, "0000")
Call armagrid
Call barraesag(Me)
Option4 = True
Call numera
End Sub
Sub numera()
q = "select * from g2 where [id_tipo_comp] = 65 "
Set rs = New ADODB.Recordset
rs.MaxRecords = 1
rs.Open q, cn1

If Not rs.EOF And Not rs.BOF Then
  t_numoc = rs("ult_num") + 1
Else
  MsgBox ("Error al inicializar comprobante")
  Exit Sub
End If
Set rs = Nothing

End Sub
Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[INS] Agrega Art.- [ENTER] Modifica Art. - [F5] Saca Art. - [F9] Continua - [F3] Sol.Mat.Prod. - [F4] Faltantes - [F6] Desc.Extra"
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  abm_solmat.Show
End If


If KeyCode = vbKeyF6 Then
 If msf1.Rows > 1 Then
  If Val(msf1.TextMatrix(msf1.Rows - 1, 0)) > 0 Then
   Load gen_descextra
   gen_descextra.t_modulo = "O"
   gen_descextra.t_funcion = "A"
   gen_descextra.Show
  End If
 End If
End If


If KeyCode = vbKeyF4 Then
 If EXISTE = "N" Then
  J = MsgBox("Confirma Importar Productos en Registro de Faltantes para este Proveedor", 4)
  If J = 6 Then
     q = "select * from a6 where [num_int] = " & para.numint_regfaltante & " and [envase] = " & c_prov.ItemData(c_prov.ListIndex)
     Set rs = New ADODB.Recordset
     rs.Open q, cn1
     While Not rs.EOF
      Set cl_prod = New productos
      ip = rs("id_producto")
      cl_prod.cargar (ip)
      d = rs("detalle")
      cu = rs("cantidad")
      pu = cl_prod.precio_ult_compra
      u = rs("unidad")
      ti = cl_prod.tasaiva
      im = cu * pu
      r = msf1.Rows
      o = "Stock"
      ido = 1
      ref = rs("renglon")
      If Val(cu) > 0 Then
        ABM_OC.msf1.AddItem r & Chr(9) & Format$(ip, "00000") & Chr(9) & d & Chr(9) & " " & Chr(9) & " " & Chr(9) & cu & Chr(9) & pu & Chr(9) & u & Chr(9) & ti & Chr(9) & im & Chr(9) & o & Chr(9) & ido & Chr(9) & ref
      End If
      Set cl_prod = Nothing
      rs.MoveNext
    Wend
    
  End If
 Else
  MsgBox ("No se puede incorporar items del Reg. de Faltantes a una OC ya emitida. Emita una nueva")
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
  If msf1.Rows > 1 Then
     Call sacatotales2
  End If
  Frame2.Enabled = True
  Frame4.Enabled = True
  t_condiciones.SetFocus
End If



If KeyCode = vbKeyInsert Then
  If msf1.Row < 20 Then
   abm_oc1.t_renglon = ""
   abm_oc1.t_nroreq = ""
   abm_oc1.t_basico = ""
   abm_oc1.t_detalle = ""
   abm_oc1.t_renglonp = ""
   abm_oc1.t_cantunit = ""
   abm_oc1.t_unidad = ""
   abm_oc1.t_importe = ""
   abm_oc1.t_renglonrf = ""
   abm_oc1.Show
 End If
End If
End Sub

Sub graba()
If EXISTE = "S" Then
  Set cl_comp = New COMPROBANTES
  Call cl_comp.cargar(65, "O", Val(t_sucursal), Val(t_numoc), 0)
  If cl_comp.numint <> 0 Then
    J = MsgBox("Comprobante existente, desea modificar", 4)
    If J = 6 Then
      cl_comp.borrar
      EXISTE = "N"
    Else
      EXISTE = "S"
    End If
  End If
  Set cl_comp = Nothing
End If

If EXISTE = "N" Then
   'oc nueva
      'On Error GoTo ERRORGRABA
      numint = saca_ultnumero_int_comp("C")
      t_numoc = Format$(saca_ultnumero_comp(65), "00000000")
      
      Set cl_comp = New COMPROBANTES
      cl_comp.actual (65)
      STOCK = cl_comp.STOCK
      ctacte = cl_comp.ctacte
      If Option4 = True Then
        moneda = "P"
      Else
       moneda = "D"
      End If
      
      infocontacto = Left$(RTrim$(t_contacto) & "  " & RTrim$(t_tecontacto), 80)
     
     Set cl_prov = New proveedores
     cl_prov.carga (c_prov.ItemData(c_prov.ListIndex))
     If cl_prov.idprov > 0 Then
      prov0 = cl_prov.razonsocial
      cuit0 = Val(cl_prov.CUIT)
     End If
     Set cl_prov = Nothing
     cn1.BeginTrans
     
     
QUERY = "INSERT INTO a5([num_int], [sucursal], [num_comprobante], [letra], [id_tipocomp], [id_proveedor], [fecha], [id_usuario], [subtotal], [iva], [no_grabado], [percep_ret], [total], " & _
"[fecha_prob_entrega], [fecha_recepcion], [estado], [ID_CODRETGAN], [ID_CUENTA], [STOCK], [CTACTE], [grabado], [estado_pago], [num_op], [obs], [condiciones], [info_contacto], " & _
"[moneda], [cotiz_dolar], [contado], [fecha_vto], [compra], [proveedor05], [cuit05], [zona], [saldo_impago], [pagos_realizados], [pago_actual], [minimo_no_imp])"
QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & Val(t_numoc) & ", 'O', 65, " & c_prov.ItemData(c_prov.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & _
 ", " & Val(t_subtotal) & ", " & Val(t_iva) & ", " & Val(t_nograbado) & ", 0, " & Val(T_TOTAL) & ",'" & t_fechaprob & "', '" & t_fechaprob & "', 'P', 0, 0, '" & STOCK & "', '" & ctacte & "', '" & _
cl_comp.grabado & "', 'X', '0000-00000000'" & ", '" & Left$(t_obs, 80) & "', '" & Left$(t_condiciones, 80) & "', '" & Left$(infocontacto, 80) & "', '" & moneda & "', " & Val(t_cotiz) & ", 'S', '" & t_fecha & "', 'N', '" & _
Left$(prov0, 50) & "', " & cuit0 & ", 1, 0, 0, 0, 0)"
            
      
      cn1.Execute QUERY
      
      For i = 1 To msf1.Rows - 1
       renglon = Val(msf1.TextMatrix(i, 0))
       If renglon > 0 Then
  
        If Val(msf1.TextMatrix(i, 3)) > 0 Then 'nro ref.
                                        
            Set rs2 = New ADODB.Recordset
            q = "select * from pro_04 where [num_referencia] = " & Val(msf1.TextMatrix(i, 3))
            rs2.MaxRecords = 1
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            If Not rs2.EOF And Not rs2.BOF Then
               totaloc = rs2("total_oc") + Val(msf1.TextMatrix(i, 5))
               If totaloc >= rs2("total_pedido") Then
                  estado = "C"
               Else
                  estado = "I"
               End If
                rs2("total_oc") = totaloc
                rs2("estado_pedido") = estado
                rs2("estado_oc") = "I"
               rs2.Update
             End If
             nr = Val(msf1.TextMatrix(i, 3))
        Else
            'creo una entrada por producto para seguirlo por el sistema
            'num_referencia auto
            Set rs2 = New ADODB.Recordset
            q = "select * from pro_04"
            rs2.MaxRecords = 1
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            rs2.AddNew
            rs2("id_producto") = Val(msf1.TextMatrix(i, 1))
            rs2("detalle") = Left$(msf1.TextMatrix(i, 2), 50)
            rs2("total_pedido") = Val(msf1.TextMatrix(i, 5))
            rs2("total_oc") = Val(msf1.TextMatrix(i, 5))
            rs2("total_recibido") = 0
            rs2("estado_pedido") = "C"
            rs2("estado_oc") = "I"  'INCOMPLETA
            rs2("fecha") = t_fecha
            rs2("id_usuario") = para.id_usuario
            rs2("observaciones") = RTrim$(msf1.TextMatrix(i, 4)) & " "
            rs2("fecha_esperado") = t_fechaprob
            rs2("id_obra") = Val(msf1.TextMatrix(i, 11))
            rs2("tipo04") = 2
            rs2.Update
            nr = rs2("num_referencia")
            Set rs2 = Nothing
        
        
            
        
        End If
        
        'actualizo faltantes y prod en oc
        Set cl_prod = New productos
        cl_prod.borraprodfaltante (Val(msf1.TextMatrix(i, 1)))
        Set cl_prod = Nothing
              
        Set rs2 = New ADODB.Recordset
        q = "select [pedidos] from a2 where [id_producto] = " & Val(msf1.TextMatrix(i, 1))
        rs2.MaxRecords = 1
        rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs2.EOF And Not rs2.BOF Then
                rs2("pedidos") = rs2("pedidos") + Val(msf1.TextMatrix(i, 5))
                rs2.Update
        End If
        Set rs2 = Nothing
        
        
        
        q = "select * from pro_05 where [num_referencia] = " & nr
        Set rs2 = New ADODB.Recordset
        rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
        If Not rs2.EOF And Not rs2.BOF Then
           rs2.MoveLast
           s = rs2("secuencia") + 1
        Else
           s = 1
        End If
        Set rs2 = Nothing
        QUERY = "INSERT INTO pro_05([num_referencia], [secuencia], [modulo], [num_int], [cantidad], [tipo_comprobante], [fecha], [unidad])"
        QUERY = QUERY & " VALUES (" & nr & ", " & s & ", 'C', " & numint & ", " & Val(msf1.TextMatrix(i, 5)) & ", 65, '" & t_fecha & "', '" & msf1.TextMatrix(i, 7) & "')"
        cn1.Execute QUERY
                 
        
        QUERY = "INSERT INTO a6([num_int], [RENGLON], [id_producto], [detalle], [cantidad], [pu], [importe], [envase], [bultos],[id_requisicion],[estado], [tasa_iva], [renglon_requisicion], [observaciones], [num_int_item], [unidad], [id_obra])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 5)) & ", " & Val(msf1.TextMatrix(i, 6)) & ", " & Val(msf1.TextMatrix(i, 9)) & ", 0, 0, 0," & " 'P', " & Val(msf1.TextMatrix(i, 8)) & ", 0,'" & Left$(msf1.TextMatrix(i, 4) & " ", 30) & "', " & nr & ", '" & msf1.TextMatrix(i, 7) & "', " & Val(msf1.TextMatrix(i, 11)) & ")"
        cn1.Execute QUERY
      
             
      Else
       'desc. extra
       QUERY = "INSERT INTO a21([num_int], [RENGLON], [descripcion], [cant_lineas])"
       QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i - 1, 0)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 3)) & ")"
       cn1.Execute QUERY
      
      End If
      
      Next i
      
      
        nc = Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numoc), "00000000")
      
       QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
       QUERY = QUERY & " VALUES ('Emite Orden de Compra:" & numint & "', " & para.id_usuario & ", 'C', " & numint & ", '" & Now & "', '[" & nc & "', 105, " & c_prov.ItemData(c_prov.ListIndex) & ")"
       cn1.Execute QUERY
      
      
      
      cn1.CommitTrans
      Set rs = Nothing
      
      J = MsgBox("Imprime O.C.", 4)
      If J = 6 Then
         Set cl_comp = New COMPROBANTES
         cl_comp.cargar2 (numint)
         If cl_comp.numint > 0 Then
           cl_comp.imprimir
         
             
         End If
      End If
 
      
      Call INICIALIZA2(Me)
      Call armagrid
      t_sucursal = Format$(glo.sucursal, "0000")
      Call numera
      t_numoc.SetFocus
Else
   MsgBox ("No se puede modificar O.C.")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If msf1.Row > 0 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 0 Then
    abm_oc1.limpia
    abm_oc1.t_renglon = msf1.Row
    abm_oc1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    abm_oc1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    abm_oc1.t_cantunit = msf1.TextMatrix(msf1.Row, 5)
    abm_oc1.t_nroreq = msf1.TextMatrix(msf1.Row, 3)
    abm_oc1.t_obs = msf1.TextMatrix(msf1.Row, 4)
    abm_oc1.t_pu = msf1.TextMatrix(msf1.Row, 6)
    abm_oc1.t_unidad = msf1.TextMatrix(msf1.Row, 7)
    abm_oc1.C_OBRA.ListIndex = buscaindice(abm_oc1.C_OBRA, Val(msf1.TextMatrix(msf1.Row, 11)))
    abm_oc1.t_renglonrf = msf1.TextMatrix(msf1.Row, 12)
    abm_oc1.Show
   Else
     Load gen_descextra
     gen_descextra.Text1 = msf1.TextMatrix(msf1.Row, 2)
     gen_descextra.t_modulo = "O"
     gen_descextra.t_funcion = "M"
     gen_descextra.Show
   End If
  End If
End If



End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
End Sub

Private Sub t_condiciones_LostFocus()
t_condiciones = RTrim$(t_condiciones) & " "
End Sub

Private Sub t_contacto_LostFocus()
t_contacto = RTrim$(t_contacto) & " "
End Sub

Private Sub t_fecha_LostFocus()
If Not IsNull(t_fecha) Then
 If Not IsDate(t_fecha) Then
   t_fecha = Format$(Now, "dd/mm/yyyy")
 Else
   t_fecha = Format$(t_fecha, "dd/mm/yyyy")
 End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If

Call verifica_fechacorte(t_fecha)
If verificaperiodo(t_fecha) = "C" Then
   MsgBox ("El periodo para el cual se deseas ingresar el comprobante esta CERRADO!!!!!")
   t_fecha.SetFocus
   t_fecha = ""
End If
End Sub

Private Sub t_fechaprob_LostFocus()
If Not IsNull(t_fechaprob) Then
 If Not IsDate(t_fechaprob) Then
  t_fechaprob = Format$(t_fecha, "dd/mm/yyyy")
 Else
  t_fechaprob = Format$(t_fechaprob, "dd/mm/yyyy")
 End If
Else
 t_fechaprob = Format$(t_fecha, "dd/mm/yyyy")
End If
End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
 Call carga_oc
 Call iniciacomp
End Sub

Private Sub t_obs_LostFocus()
t_obs = RTrim$(t_obs) & " "
End Sub

Private Sub t_tecontacto_LostFocus()
t_tecontacto = RTrim$(t_tecontacto) & " "
End Sub

Private Sub t_total_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If

End Sub
