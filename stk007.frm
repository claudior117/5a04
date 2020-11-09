VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form stk_egreso 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SALIDAS DE MERCADERIA DE  STOCK"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11955
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   8640
      TabIndex        =   16
      Top             =   0
      Width           =   3135
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   $"stk007.frx":0000
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   2895
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   4695
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8281
      _Version        =   393216
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
      Height          =   2175
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   8295
      Begin VB.TextBox t_comp 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1320
         Width           =   3255
      End
      Begin VB.ComboBox c_obra 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox t_detalle 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   49
         TabIndex        =   4
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox t_numoc 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         MaxLength       =   8
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Responsable:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Obra:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Fecha Salida:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Caption         =   "Movimiento Interno Nro.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   8
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "stk007.frx":0095
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
         Picture         =   "stk007.frx":0917
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
      Top             =   8415
      Width           =   11955
      _ExtentX        =   21087
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
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "stk_egreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim gcolumna As Integer
Dim EXISTE As String
Dim cantidadp As Double

Private Sub btnsale_Click()
Unload Me
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 1000
msf1.ColWidth(2) = 4500
msf1.ColWidth(3) = 1200
msf1.ColWidth(4) = 3200

msf1.TextMatrix(0, 0) = "Reng."
msf1.TextMatrix(0, 1) = "Id.Prod."
msf1.TextMatrix(0, 2) = "Descripcion"
msf1.TextMatrix(0, 3) = "Cantidad"
msf1.TextMatrix(0, 4) = "Detalle"


End Sub








Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 5)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()
Call carga_obras(C_OBRA, "A")
C_OBRA.ListIndex = 0
Call armagrid
Call barraesag(Me)


End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [ENTER] Modifica - [F5] Elimina - [F9] Graba"
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
 If msf1.Rows > 1 Then
  J = MsgBox("Confirma Grabar Salida Stock ", 4)
  If J = 6 Then
   If verificaperiodog(t_fecha) = "A" Then
     Call graba
   Else
     MsgBox ("Periodo cerrado. Imposible grabar operacion")
   End If
  End If
 End If
End If

If KeyCode = vbKeyInsert Then
   stk_EGRESO2.t_renglon = ""
   stk_EGRESO2.t_cantidad = ""
   stk_EGRESO2.Show
End If
End Sub

Sub graba()
If EXISTE = "N" Then
   'oc nueva
      'On Error GoTo ERRORGRABA
      
      
      cn1.BeginTrans
      QUERY = "INSERT INTO stk_02([fecha], [letra], [num_comprobante], [id_usuario], [detalle], [sucursal], [tipo_comprobante], [id_proveedor], [id_obra])"
      QUERY = QUERY & " VALUES ('" & t_fecha & "', 'X', " & Val(t_numoc) & ", " & para.id_usuario & ", '" & t_detalle & " ', " & glo.sucursal & ", 30, 0, " & C_OBRA.ItemData(C_OBRA.ListIndex) & ")"
    
      cn1.Execute QUERY
      
      qr = "SELECT @@IDENTITY AS NewID"
      Set rs = cn1.Execute(qr)
      numint = rs.Fields("NewID").Value

      
      
      For i = 1 To msf1.Rows - 1
        QUERY = "INSERT INTO stk_03([num_int], [RENGLON], [id_producto], [descripcion], [unidad], [detalle], [cantidad], [ubicacion])"
        QUERY = QUERY & " VALUES (" & numint & ", " & Val(msf1.TextMatrix(i, 0)) & ", " & Val(msf1.TextMatrix(i, 1)) & ", '" & msf1.TextMatrix(i, 2) & "', ' ', '" & msf1.TextMatrix(i, 4) & "', " & msf1.TextMatrix(i, 3) & " , 'S')"
        cn1.Execute QUERY
      
        QUERY = "INSERT INTO stk_01([fecha], [id_producto], [cantidad], [ubicacion], [comprobante], [descripcion], [num_mov_int], [modulo])"
        QUERY = QUERY & " VALUES ('" & t_fecha & "', " & Val(msf1.TextMatrix(i, 1)) & ", " & Val(msf1.TextMatrix(i, 3)) & ", 'S', 'Salida Stk " & Format$(numint, "00000000") & "', '" & Left$(C_OBRA, 49) & " ', " & numint & ", 'S')"
        cn1.Execute QUERY
  
        QUERY = "update a2 set  [stock]=[stock] - " & Val(msf1.TextMatrix(i, 3))
        QUERY = QUERY & " where [id_producto]= " & Val(msf1.TextMatrix(i, 1))
        cn1.Execute QUERY
      Next i
      
      cn1.CommitTrans
      Set rs = Nothing
      
      J = MsgBox("Imprime Movimiento Interno", 4)
      If J = 6 Then
         Set cl_stock = New STOCK
         cl_stock.imprimir (numint)
         Set cl_stock = Nothing
      End If

      
      Call INICIALIZA2(Me)
      Call armagrid
      t_numoc.SetFocus
Else
   MsgBox ("No se puede modificar Mov. Interno stock")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  

End Sub
Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    stk_ingreso2.t_renglon = msf1.Row
    stk_ingreso2.t_basico = msf1.TextMatrix(msf1.Row, 1)
    stk_ingreso2.t_detalle = msf1.TextMatrix(msf1.Row, 2)
    stk_ingreso2.t_cantidad = msf1.TextMatrix(msf1.Row, 3)
    stk_ingreso2.t_renglonp = msf1.TextMatrix(msf1.Row, 5)
    stk_ingreso2.Show
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

Private Sub t_letra_LostFocus()
If t_letra <> "" Then
  t_letra = Format$(t_letra, ">@")
Else
  t_letra = "X"
End If
End Sub

Private Sub t_numoc_GotFocus()
Call INICIALIZA2(Me)

End Sub

Private Sub t_numoc_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub t_numoc_LostFocus()
 Call carga_oc
End Sub
Sub carga_oc()
q = "SELECT * FROM STK_02, STK_03 WHERE STK_02.[NUM_INT] = " & Val(t_numoc) & "  AND [tipo_comprobante] = 20 and STK_02.[NUM_INT] = STK_03.[NUM_INT]"
Set rs = New ADODB.Recordset
rs.Open q, cn1
r = 0
If Not rs.EOF And Not rs.BOF Then
    EXISTE = "S"
    t_fecha = rs("FECHA")
    t_detalle = rs("stk_02.DETALLE")
    Call armagrid
    r = 1
    While Not rs.EOF
      msf1.AddItem r & Chr$(9) & rs("id_producto") & Chr$(9) & rs("descripcion") & Chr$(9) & rs("cantidad") & Chr$(9) & rs("stk_03.detalle") & Chr$(9) & rs("ubicacion")
      r = r + 1
      rs.MoveNext
    Wend
Else
  EXISTE = "N"
  Call armagrid
  t_fecha = ""
  t_detalle = ""
End If

End Sub

Private Sub t_sucursal_LostFocus()
t_sucursal = Format$(Val(t_sucursal), "0000")

End Sub
