VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form prod_solmat 
   BackColor       =   &H00E0E0E0&
   Caption         =   "EMITIR SOLICITUD DE MATERIALES"
   ClientHeight    =   8835
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12195
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8835
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   7200
      Width           =   9135
      Begin VB.TextBox t_obs 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   79
         TabIndex        =   5
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4935
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8705
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   9255
      Begin VB.TextBox t_fechaprob 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   6360
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1575
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
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.ComboBox c_obra 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Text            =   "c_obra"
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Esperado:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   6
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
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Solicitud de Materiales Nro:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Obra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10080
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Pro001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Pro001.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   9
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
      TabIndex        =   7
      Top             =   8580
      Width           =   12195
      _ExtentX        =   21511
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
            TextSave        =   "20/11/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:04 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "prod_solmat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double
Sub carga_oc()
End Sub

Private Sub btnacepta_Click()
 If msf1.Rows > 1 Then
  J = MsgBox("Confirma Grabar Solicitud", 4)
  If J = 6 Then
   Call graba
  End If
 End If
End Sub

Private Sub btnsale_Click()
J = MsgBox("Confirma salir", 4)
If J = 6 Then
 Unload Me
End If
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 8
msf1.ColWidth(0) = 500
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 4000
msf1.ColWidth(3) = 1400
msf1.ColWidth(4) = 2500
msf1.ColWidth(5) = 1000
msf1.ColWidth(6) = 500
msf1.ColWidth(7) = 1000

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Detalle"
msf1.TextMatrix(0, 3) = "Fecha Esp."
msf1.TextMatrix(0, 4) = "Observaciones"
msf1.TextMatrix(0, 5) = "Cantidad"
msf1.TextMatrix(0, 6) = "Unidad"
msf1.TextMatrix(0, 7) = "Ref."

End Sub






Private Sub c_obra_LostFocus()
If C_OBRA.ListIndex < 0 Then
  C_OBRA.ListIndex = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
End Select

End Sub

Private Sub Form_Load()

Call carga_obras(C_OBRA, "E")
'c_obra.ListIndex = 0
t_sucursal = Format$(glo.sucursal, "0000")
Call armagrid
Call barraesag(Me)


End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Termina "
If msf1.Rows > 1 Then
  msf1.FocusRect = flexFocusNone
Else
  msf1.FocusRect = flexFocusLight
End If
End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
 If msf1.Rows > 2 Then
    msf1.RemoveItem (msf1.Row)
 Else
   Call armagrid
 End If
End If

If KeyCode = vbKeyF9 Then
  t_obs.Enabled = True
  t_obs.SetFocus
End If


If KeyCode = vbKeyInsert Then
  If msf1.Row < 35 Then
   PROD_SOLMAT1.t_renglon = ""
   PROD_SOLMAT1.t_fechaesperado = ""
   PROD_SOLMAT1.t_basico = ""
   PROD_SOLMAT1.t_detalle = ""
   PROD_SOLMAT1.t_renglonp = ""
   PROD_SOLMAT1.t_cantunit = ""
   PROD_SOLMAT1.t_unidad = ""
   PROD_SOLMAT1.Show
 End If
End If
End Sub

Sub graba()
If EXISTE = "N" Then
   'oc nueva
      'On Error GoTo ERRORGRABA
      numint = saca_ultnumero_int_comp("P")
      
      Set cl_compprod = New comprobantes_produccion
      cl_compprod.sacaultimonumero (1)
      If cl_compprod.numcomp > 0 Then
         cn1.BeginTrans
         QUERY = "INSERT INTO pro_01([num_int], [sucursal], [num_comprobante], [id_tipocomp], [id_obra], [fecha], [id_usuario], [fecha_esperado], [estado], [observaciones])"
         QUERY = QUERY & " VALUES (" & numint & ", " & Val(t_sucursal) & ", " & cl_compprod.numcomp & ", 1, " & C_OBRA.ItemData(C_OBRA.ListIndex) & ", '" & t_fecha & "', " & para.id_usuario & ", '" & t_fechaprob & "', 'P', '" & RTrim$(t_obs) & " " & "')"
         cn1.Execute QUERY
      
         For i = 1 To msf1.Rows - 1
            'creo una entrada por producto para sehguirlo por el sistema
            'num_referencia auto
             
            Set rs2 = New ADODB.Recordset
            q = "select * from pro_04"
            rs2.MaxRecords = 1
            rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
            rs2.AddNew
            rs2("id_producto") = Val(msf1.TextMatrix(i, 1))
            rs2("detalle") = msf1.TextMatrix(i, 2)
            rs2("total_pedido") = Val(msf1.TextMatrix(i, 5))
            rs2("total_oc") = 0
            rs2("total_recibido") = 0
            rs2("estado_pedido") = "I"  'I INCOMPLETO  'C COMPLETO
            rs2("estado_oc") = "I"  'I INCOMPLETA  C COMPLETA
            rs2("fecha") = t_fecha
            rs2("id_usuario") = para.id_usuario
            rs2("observaciones") = RTrim$(msf1.TextMatrix(i, 4)) & " "
            rs2("fecha_esperado") = t_fechaprob
            rs2("id_obra") = C_OBRA.ItemData(C_OBRA.ListIndex)
            rs2("tipo04") = 1
            rs2.Update
            nr = rs2("num_referencia")
            Set rs2 = Nothing
            
            QUERY = "INSERT INTO pro_02([num_int], [RENGLON], [id_producto], [descripcion], [cantidad], [fecha_esperado], [observaciones], [num_referencia], [num_int_oc], [unidad])"
            QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', " & Val(msf1.TextMatrix(i, 5)) & ", '" & msf1.TextMatrix(i, 3) & "', '" & RTrim$(msf1.TextMatrix(i, 4)) & " " & "', " & nr & ", 0, '" & msf1.TextMatrix(i, 6) & " ')"
            cn1.Execute QUERY
         
            QUERY = "INSERT INTO pro_05([num_referencia], [secuencia], [modulo], [num_int], [cantidad], [tipo_comprobante], [fecha], [unidad], [obs])"
            QUERY = QUERY & " VALUES (" & nr & ", 1, 'P', " & numint & ", " & Val(msf1.TextMatrix(i, 5)) & ", 1, '" & t_fecha & "', '" & msf1.TextMatrix(i, 6) & " ', 'Pedido')"
            cn1.Execute QUERY
         
         
         
         Next i
      
      cn1.CommitTrans
      Set rs = Nothing
      
      J = MsgBox("Imprime Comprobante", 4)
      If J = 6 Then
         cl_compprod.cargar2 (numint)
         If cl_compprod.numint > 0 Then
           cl_compprod.imprimir
         End If
      End If
 
      
      Call INICIALIZA2(Me)
      Call armagrid
      t_numoc.SetFocus
   End If
   Set cl_compprod = Nothing
Else
   MsgBox ("No se puede modificar Solicitud")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos Modulo:Graba")
  Exit Sub

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    PROD_SOLMAT1.limpia
    PROD_SOLMAT1.t_renglon = msf1.Row
    PROD_SOLMAT1.t_basico = msf1.TextMatrix(msf1.Row, 1)
    PROD_SOLMAT1.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    PROD_SOLMAT1.t_cantunit = msf1.TextMatrix(msf1.Row, 5)
    PROD_SOLMAT1.t_fechaesperado = msf1.TextMatrix(msf1.Row, 3)
    PROD_SOLMAT1.t_obs = msf1.TextMatrix(msf1.Row, 4)
    PROD_SOLMAT1.t_unidad = msf1.TextMatrix(msf1.Row, 6)

    PROD_SOLMAT1.Show
  End If

End If
End Sub

Private Sub msf1_LostFocus()
Call barraesag(Me)
msf1.FocusRect = flexFocusLight
End Sub



Private Sub t_fecha_LostFocus()
If Not IsDate(t_fecha) Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
Else
  t_fecha = Format$(t_fecha, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
End Sub

Private Sub t_fechaprob_LostFocus()
If Not IsDate(t_fechaprob) Then
  t_fechaprob = Format$(Now, "dd/mm/yyyy")
Else
  t_fechaprob = Format$(t_fechaprob, "dd/mm/yyyy")
End If
End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
If t_numoc = "" Then
    EXISTE = "N"
Else
   Call carga_oc
End If
End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_obs_LostFocus()
t_obs = RTrim$(t_obs) & " "
End Sub

