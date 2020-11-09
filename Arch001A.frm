VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ABM_PROv 
   BackColor       =   &H00E0E0E0&
   Caption         =   "PROVEEDORES"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Transportes"
      Height          =   735
      Left            =   4800
      TabIndex        =   21
      Top             =   7320
      Width           =   3135
      Begin VB.CheckBox Check1 
         BackColor       =   &H8000000A&
         Caption         =   "Ver solo empresas de Tranportes"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenados por"
      Height          =   735
      Left            =   240
      TabIndex        =   18
      Top             =   7320
      Width           =   4455
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razon Social"
         Height          =   255
         Left            =   2160
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id."
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5775
      Left            =   240
      TabIndex        =   17
      Top             =   1440
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10186
      _Version        =   393216
      AllowBigSelection=   0   'False
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtrar"
      Height          =   1335
      Left            =   7080
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin VB.TextBox t_contacto 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   960
         Width           =   2535
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
         Caption         =   "Contacto"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Localidad"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
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
      Left            =   240
      TabIndex        =   4
      Top             =   0
      Width           =   6735
      Begin VB.CommandButton Command5 
         Caption         =   "&Enviar Correo"
         Height          =   735
         Left            =   5400
         Picture         =   "Arch001A.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Listar"
         Height          =   735
         Left            =   4080
         Picture         =   "Arch001A.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Borrar"
         Height          =   735
         Left            =   2760
         Picture         =   "Arch001A.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Modificar"
         Height          =   735
         Left            =   1440
         Picture         =   "Arch001A.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   120
         Picture         =   "Arch001A.frx":0C28
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
         Picture         =   "Arch001A.frx":0F32
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
         Picture         =   "Arch001A.frx":17B4
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
      Width           =   12105
      _ExtentX        =   21352
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "ABM_PROv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 14
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
msf1.ColWidth(10) = 500
msf1.ColWidth(11) = 1500
msf1.ColWidth(12) = 2500
msf1.ColWidth(13) = 0
msf1.TextMatrix(0, 0) = "Id."
msf1.TextMatrix(0, 1) = "Razon Social"
msf1.TextMatrix(0, 2) = "Direccion"
msf1.TextMatrix(0, 3) = "Localidad"
msf1.TextMatrix(0, 4) = "Te"
msf1.TextMatrix(0, 5) = "Cuit"
msf1.TextMatrix(0, 6) = "Iva"
msf1.TextMatrix(0, 7) = "Email"
msf1.TextMatrix(0, 8) = "Contacto"
msf1.TextMatrix(0, 9) = "Te Contacto"
msf1.TextMatrix(0, 10) = ""
msf1.TextMatrix(0, 11) = "Provincia"
msf1.TextMatrix(0, 12) = "Cuenta"

For i = 1 To 9
  msf1.ColAlignment(i) = 1 'izq
Next i
msf1.ColAlignment(0) = 9 'der
End Sub

Private Sub btnacepta_Click()
Call limpia
msf1.SetFocus
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub


Private Sub Command1_Click()
Call nivel_acceso(2)
If para.id_grupo_modulo_actual >= 3 Then
 Call INICIALIZA2(abm_prov1)
 abm_prov1!t_funcion = "A"
 abm_prov1.Show
Else
 Call sinpermisos
End If
End Sub

Private Sub Command2_Click()
'On Error GoTo e1
If msf1.Rows > 0 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 3 Then
  If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
   abm_prov1!t_funcion = "M"
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

Sub LLENACAMPOS()
'On Error GoTo ERROR1
Set rs = New ADODB.Recordset
q = "select * from a1 where [id_proveedor] = " & Val(msf1.TextMatrix(msf1.Row, 0))
rs.Open q, cn1
 abm_prov1!t_id = rs("id_proveedor")
 abm_prov1!t_descripcion = rs("denominacion")
 abm_prov1!t_direccion = rs("direccion")
 abm_prov1!t_te = rs("te")
 abm_prov1!t_localidad = rs("localidad")
 abm_prov1!t_cp = rs("cp")
 abm_prov1!c_provincia.ListIndex = buscaindice(abm_prov1!c_provincia, rs("id_provincia"))
 abm_prov1!t_email = rs("email")
 abm_prov1!t_cuit = rs("cuit")
 abm_prov1!c_iva.ListIndex = buscaindice(abm_prov1!c_iva, rs("cod_tipoiva"))
 abm_prov1!c_ret.ListIndex = buscaindice(abm_prov1!c_ret, rs("id_codretgan"))
 abm_prov1!t_contacto = rs("contacto")
 abm_prov1!t_tecontacto = rs("te_contacto")
 abm_prov1!t_emailcontacto = rs("email_contacto")
 abm_prov1!t_inscgan = rs("inscripto_gan")
 abm_prov1!t_fecha_vto_exepcion = rs("fecha_vto_exepcion_ib")
 abm_prov1!t_numib = rs("num_ib")
 abm_prov1!c_retib.ListIndex = buscaindice(abm_prov1!c_retib, rs("id_codretib"))
 abm_prov1!c_cuenta.ListIndex = buscaindice(abm_prov1!c_cuenta, rs("id_cuenta_a1"))
 abm_prov1.Show

Set rs = Nothing

Exit Sub
ERROR1:
  MsgBox ("Error al Cargar Proveedores. Proc.: LLENACAMPOS")
  Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo e1
If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 7 Then
   abm_prov1!t_funcion = "B"
   Call LLENACAMPOS
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





Private Sub Command5_Click()
Dim RetVal As Long
On Error GoTo e1
 Call nivel_acceso(2)
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
    If t_localidad = "" Then
      l = "<Todas>"
    Else
      l = "<" & t_localidad & ">"
    End If
    Call imprimegrid(msf1, c(), "                                                                                          LISTADO DE PROVEEDORES", "", "   Localidad: " & l, " ", 42, 10, True, False, "H")
  End If



End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub
Sub limpia()
espere.Show
espere.Label1 = "ESPERE [Leyendo Base de Datos]... "
espere.Refresh

Call armagrid
Dim q As String
q = "select * from a1, g3, i_02, g6, c_01 where a1.[cod_tipoiva] = g3.[cod_tipoiva]  and [id_impuesto] = 217 and [id_concepto] = [id_codretgan] and a1.[id_tipoib] = g6.[id_tipoib]  and [id_cuenta_a1] = [id_cuenta]"
c = " and "
If t_prov <> "" Then
 q = q & c & " [denominacion] like '%" & t_prov & "%'"
 c = " and "
End If
If t_localidad <> "" Then
 q = q & c & " [localidad] like '%" & t_localidad & "%'"
 c = " and "
End If

If t_contacto <> "" Then
 q = q & c & " [contacto] like '%" & t_contacto & "%'"
 c = " and "
End If

If Check1 = 1 Then
  q = q & c & " [transporte] = 'S'"
End If
If Option2 = True Then
  q = q & " order by [denominacion]"
Else
   q = q & " order by [id_proveedor]"
End If


Set rs = New ADODB.Recordset
rs.Open q, cn1
c = 0
While Not rs.EOF
 msf1.AddItem rs("id_proveedor") & Chr$(9) & rs("denominacion") & Chr$(9) & rs("direccion") & Chr$(9) & rs("localidad") & Chr$(9) & rs("te") & Chr$(9) & rs("cuit") & Chr$(9) & rs("abreviatura") & Chr$(9) & rs("email") & Chr$(9) & rs("contacto") & Chr$(9) & rs("te_contacto") & Chr$(9) & rs("id_provincia") & Chr$(9) & rs("provincia") & Chr$(9) & rs("c_01.descripcion")
 rs.MoveNext
 c = c + 1
Wend
msf1.AddItem ""
msf1.AddItem "" & Chr$(9) & "Total de Registros : " & c

Set rs = Nothing
'Call conectaradodc(Adodc1, q, cn1)
'DataGrid1.Refresh
Call INICIALIZA2(abm_prov1)
Unload espere
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call armagrid
Load abm_prov1
Option2 = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload abm_prov1
End Sub

Private Sub msf1_GotFocus()
StatusBar1.Panels.Item(2) = "[F1] Datos Prov.  - [F3] Padron IB -  [F4] Saca - [F7] Imprime - [F11]Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
 If msf1.Rows > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
     com_proveedor.t_id = Val(msf1.TextMatrix(msf1.Row, 0))
     com_proveedor.carga
     com_proveedor.Show
   End If
 End If
End If


If KeyCode = vbKeyF7 Then
   Call imprime
End If

If KeyCode = vbKeyF3 Then
 If msf1.Rows > 0 Then
   If Val(msf1.TextMatrix(msf1.Row, 0)) > 1 Then
     Load gen_consultaib
     gen_consultaib.t_id = Val(msf1.TextMatrix(msf1.Row, 0))
     gen_consultaib.t_tipo = "P"
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

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If


End Sub

Private Sub t_contacto_GotFocus()
t_contacto = ""

End Sub

Private Sub t_prov_GotFocus()
t_prov = ""
End Sub

Private Sub t_localidad_GotFocus()
t_localidad = ""
End Sub

Private Sub t_prov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call limpia
  msf1.SetFocus
End If
End Sub
