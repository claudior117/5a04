VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_ABM_cli 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CLIENTES"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12225
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1335
      Left            =   7200
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin VB.ComboBox C_vend 
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox t_localidad 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox t_prov 
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Vendedor:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Localidad:"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social:"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command5 
         Caption         =   "&Enviar Correo"
         Height          =   735
         Left            =   5400
         Picture         =   "vta001A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "vta001A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "vta001A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "vta001A.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "vta001A.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "vta001A.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "vta001A.frx":17B4
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   8415
      Width           =   12225
      _ExtentX        =   21564
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
            TextSave        =   "24/08/2016"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:48 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10186
      _Version        =   393216
      AllowBigSelection=   0   'False
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "vta_ABM_cli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim gquery As String

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
   c_vend.ListIndex = 0
End If
End Sub

Private Sub Command1_Click()
If para.id_grupo_modulo_actual >= 5 Then
Call INICIALIZA2(vta_abm_cli1)
 vta_abm_cli1!t_funcion = "A"
 vta_abm_cli1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 5 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   vta_abm_cli1!t_funcion = "M"
   Call LLENACAMPOS
  End If
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 MsgBox ("Error en Datos")
 Exit Sub
End Sub

Sub LLENACAMPOS()
'On Error GoTo ERROR1
Set rs = New adodb.Recordset
q = "select * from vta_01 where [id_cliente] = " & Val(msf1.TextMatrix(msf1.Row, 0))
rs.Open q, cn1
 vta_abm_cli1!t_id = rs("id_cliente")
 vta_abm_cli1!t_descripcion = rs("denominacion")
 vta_abm_cli1!t_direccion = rs("direccion")
 vta_abm_cli1!t_te = rs("te")
 vta_abm_cli1!t_localidad = rs("localidad")
 vta_abm_cli1!t_cp = rs("cp")
 vta_abm_cli1!c_provincia.ListIndex = buscaindice(vta_abm_cli1!c_provincia, rs("id_prov"))
 vta_abm_cli1!t_email = rs("email")
 vta_abm_cli1!t_cuit = rs("cuit")
 vta_abm_cli1!c_iva.ListIndex = buscaindice(vta_abm_cli1!c_iva, rs("id_tipoiva"))
 vta_abm_cli1!t_credito = rs("limite_credito")
 vta_abm_cli1!c_vend.ListIndex = buscaindice(vta_abm_cli1!c_vend, rs("id_vendedor"))
 vta_abm_cli1!t_operadorgranos = rs("inscripto_operador_granos")
 vta_abm_cli1!T_PERCIB = rs("percive_ib")
 vta_abm_cli1!t_saldoi = rs("saldo_incobrable")
 vta_abm_cli1!t_direccionlocal = rs("direccion_local")
 
 If rs("id_proveedor") > 0 Then
   vta_abm_cli1!c_prov.ListIndex = buscaindice(vta_abm_cli1!c_prov, rs("id_proveedor"))
 Else
   vta_abm_cli1!c_prov.ListIndex = 0
 End If
 vta_abm_cli1!t_observaciones = rs("observaciones")
 vta_abm_cli1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Clientes. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If msf1.Rows > 0 Then
 If para.id_grupo_modulo_actual >= 7 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   vta_abm_cli1!t_funcion = "B"
   Call LLENACAMPOS
  End If
 Else
  Call sinpermisos
 End If
End If

Exit Sub
e1:
 Exit Sub
End Sub

Private Sub Command4_Click()
Call imprime
End Sub

Sub imprime()
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 0
    c(1) = 1
    c(2) = 2
    c(3) = 3
    c(4) = 4
    c(5) = 5
    c(6) = 6
   
    For i = 7 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "                                                                                                LISTADO DE CLIENTES", "", "    Localidad: " & t_localidad, "    Vendedor: " & c_vend, 42, 10, True, False, "H")
      
  End If


End Sub



Private Sub Command5_Click()
Dim RetVal As Long
'On Error GoTo e1
 Call nivel_acceso(1)
 email = ""
 If para.id_grupo_modulo_actual >= 5 Then
   If msf1.Row <= msf1.RowSel Then
         For i = msf1.Row To msf1.RowSel
           If Len(RTrim$(msf1.TextMatrix(i, 7))) > 10 Then
               If Len(email) > 0 Then
                 email = email & ";"
               End If
               email = email & RTrim$(msf1.TextMatrix(i, 7))
           End If
         Next
            ' -- cuando se elecciona desde abajo hacia arriba
    Else
        For i = msf1.RowSel To msf1.Row
           If Len(RTrim$(msf1.TextMatrix(i, 7))) > 10 Then
               If Len(email) > 0 Then
                 email = email & ";"
               End If
               email = email & RTrim$(msf1.TextMatrix(i, 7))
           End If
        Next
    End If
    
    
    If Len(email) > 0 Then
      RetVal = ShellExecute(Me.hWnd, "Open", "mailto:" & email, vbNullString, vbNullString, vbNormalFocus)
    End If
     
 Else
  Call sinpermisos
 End If

Exit Sub
e1:
 Exit Sub

End Sub





Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 10
msf1.FixedCols = 2
msf1.SelectionMode = flexSelectionFree
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 600
msf1.ColWidth(1) = 3000
msf1.ColWidth(2) = 2500
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 2500
msf1.ColWidth(5) = 1500
msf1.ColWidth(6) = 600
msf1.ColWidth(7) = 3000
msf1.ColWidth(8) = 2000
msf1.ColWidth(9) = 1500
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Razon Social"
msf1.TextMatrix(0, 2) = "Direccion"
msf1.TextMatrix(0, 3) = "Localidad"
msf1.TextMatrix(0, 4) = "Te"
msf1.TextMatrix(0, 5) = "Cuit/DNI"
msf1.TextMatrix(0, 6) = "Iva"
msf1.TextMatrix(0, 7) = "Email"
msf1.TextMatrix(0, 8) = "Vendedor"
msf1.TextMatrix(0, 9) = "Te Contacto"

For i = 1 To 9
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(0) = 9 'der
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
Dim q As String
Call armagrid
espere.Show
espere.Label1 = "ESPERE [Leyendo Base de Datos]... "
espere.Refresh

q = "select * from vta_01, g3, vta_05 where vta_01.[id_tipoiva] = g3.[cod_tipoiva] and vta_01.[id_vendedor] = vta_05.[id_vendedor]"
c = " and "
If t_prov <> "" Then
 q = q & c & " [vta_01.denominacion] like '%" & t_prov & "%'"
 c = " and "
End If
If t_localidad <> "" Then
 q = q & c & " [localidad] like '%" & t_localidad & "%'"
 c = " and "
End If

If c_vend.ListIndex > 0 Then
 q = q & c & " vta_01.[id_vendedor] =" & c_vend.ItemData(c_vend.ListIndex)
 c = " and "
End If

If Option2 = True Then
  q = q & " order by [vta_01.denominacion]"
Else
   q = q & " order by [id_cliente]"
End If

Set rs = New adodb.Recordset
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("id_cliente") & Chr$(9) & rs("vta_01.denominacion") & Chr$(9) & rs("direccion") & Chr$(9) & rs("localidad") & Chr$(9) & rs("te") & Chr$(9) & rs("cuit") & Chr$(9) & rs("abreviatura") & Chr$(9) & rs("email") & Chr$(9) & rs("vta_05.denominacion")
 rs.MoveNext
 c = c + 1
Wend
 msf1.AddItem ""
 msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
Call INICIALIZA2(vta_abm_cli1)
Unload espere
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Load vta_abm_cli1
Load vta_clientes

Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0

Option2 = True
Call armagrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload vta_abm_cli1
Unload vta_clientes
End Sub

Private Sub msf1_GotFocus()
StatusBar1.Panels.Item(2) = "[F1] Datos   - [F3] Padron IB -  [F4] Saca - [F7] Imprime - [F11]Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
 If msf1.Rows > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
     vta_clientes.t_id = Val(msf1.TextMatrix(msf1.Row, 0))
     vta_clientes.carga
     vta_clientes.Show

   End If
 End If
End If

If KeyCode = vbKeyF3 Then
 If msf1.Rows > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
     Load gen_consultaib
     gen_consultaib.t_id = Val(msf1.TextMatrix(msf1.Row, 0))
     gen_consultaib.t_tipo = "C"
      gen_consultaib.Show
     gen_consultaib.carga
   End If
 End If
End If


If KeyCode = vbKeyF4 Then
 If msf1.Rows > 0 Then
   msf1.RemoveItem msf1.Row
  End If
End If


If KeyCode = vbKeyF7 Then
 If msf1.Rows > 0 Then
   
   Call imprime
    
 End If
End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If
End Sub

Private Sub t_localidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call limpia
  msf1.SetFocus
End If

End Sub

Private Sub t_prov_GotFocus()
t_prov = ""
End Sub

Private Sub t_localidad_GotFocus()
t_localidad = ""
End Sub

Private Sub t_localidad_LostFocus()
Call limpia
End Sub

Private Sub t_prov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call limpia
  msf1.SetFocus
End If
End Sub
