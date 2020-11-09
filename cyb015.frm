VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_cc_detalleb 
   BackColor       =   &H00E0E0E0&
   Caption         =   "DETALLE DE MOVIMIENTOS BANCARIOS"
   ClientHeight    =   8670
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   11895
   End
   Begin VB.Frame CUIT 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11655
      Begin VB.TextBox t_numint 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   10200
         MaxLength       =   10
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Numero Interno"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   10200
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10320
      TabIndex        =   2
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cyb015.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb015.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   3
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
      TabIndex        =   1
      Top             =   8415
      Width           =   12090
      _ExtentX        =   21325
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
            TextSave        =   "09:43"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_cc_detalleb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String
Dim l2 As String


Private Sub btnsale_Click()
Me.Hide
End Sub


Private Sub Form_Activate()
Call carga
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub carga()
List1.clear
l1 = "---------------------------------------------------------------------------------------------------"
l2 = "*************************************"

If t_numint <> "" Then
  q = "select * from cyb_04, cyb_06, cyb_01 where [num_mov_banco] = " & Val(t_numint) & " and cyb_04.[id_tipomov] = cyb_06.[id_tipomov] and cyb_04.[id_banco] = cyb_01.[id_forma_pago] "
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     List1.AddItem l2
     List1.AddItem rs("cyb_06.descripcion")
     List1.AddItem l2
     List1.AddItem Space$(25) & "Numero Interno Mov......:" & Format$(rs("num_mov_banco"), "00000")
     List1.AddItem Space$(25) & "Numero Comprobante......:" & Format$(rs("num_comp"), "00000000")
     List1.AddItem Space$(25) & "Banco...................:" & Format$(rs("id_banco"), "00000") & ") " & rs("cyb_01.descripcion")
     List1.AddItem Space$(25) & "Fecha...................:" & rs("fecha")
     List1.AddItem Space$(25) & "Estado Conc. ¿Entro?....:" & rs("entro")
     List1.AddItem Space$(25) & "Ubicacion...............:" & rs("cyb_04.ubicacion")
     List1.AddItem Space$(25) & "Fecha Diferida..........:" & rs("fecha_dif")
     List1.AddItem Space$(25) & "Fecha Acreditacion......:" & rs("fecha_acreed")
     List1.AddItem Space$(25) & "Detalle.................:" & rs("detalle")
     List1.AddItem Space$(25) & "Generado por Modulo.....:" & rs("modulo")
     List1.AddItem Space$(25) & "Nro.Int. Operacion......:" & rs("num_mov_int")
     
     q = "select * from c_02 where [num_mov_int] = " & Val(t_numint) & " and [modulo] = 'B'"
     Set rs1 = New ADODB.Recordset
     rs1.Open q, cn1
     If Not rs1.EOF And Not rs1.BOF Then
       na = Format$(rs1("num_interno"), "0000000000")
     Else
       na = "0000000000"
     End If
     List1.AddItem Space$(25) & "Asiento Int..............: " & na
     Set rs1 = Nothing
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem Space$(5) & "Importe.................:" & Format$(rs("importe"), "######0.00")
     List1.AddItem ""
     List1.AddItem ""
     Set rs = Nothing
     
     Call movcaja
          
     Call movcht
     
     Call MOVCHP
     
     
 End If
 Set rs = Nothing
End If
End Sub

Sub movcaja()
     i = Space$(10)
     e = Space$(10)
     q = "Select * from cyb_05, cyb_01 where [num_mov_int] = " & Val(t_numint) & " and [modulo] = 'B' and cyb_05.[id_forma_pago] = cyb_01.[id_forma_pago]"
     Set rs = New ADODB.Recordset
     rs.Open q, cn1
     pasada = 0
     While Not rs.EOF
       If pasada = 0 Then
          List1.AddItem "Caja"
          List1.AddItem l1
          List1.AddItem "Tipo         Descripcion                        Ingresos    Egresos   Operacion          "
          List1.AddItem l1
          pasada = 1
       End If
       tipo = "[" & Left$(rs("cyb_01.descripcion"), 10) & "]"
       Desc = Format$(Left$(rs("cyb_05.descripcion"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       If rs("ubicacion") = "D" Then
             RSet i = Format$(rs("importe"), "######0.00")
             RSet e = Format$(0, "######0.00")
       Else
             RSet e = Format$(rs("importe"), "######0.00")
             RSet i = Format$(0, "######0.00")
       End If
       o = Format$(Left$(rs("operacion"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       List1.AddItem tipo & " " & Desc & "  " & i & "  " & e & "  " & o
       rs.MoveNext
     Wend
     If pasada = 1 Then
        List1.AddItem ""
        List1.AddItem ""
     End If
     Set rs = Nothing
End Sub

Sub movcht()
     i = Space$(10)
     q = "Select * from cyb_03 where [num_mov_banco_e] = " & Val(t_numint)
     Set rs = New ADODB.Recordset
     rs.Open q, cn1
     pasada = 0
     While Not rs.EOF
       If pasada = 0 Then
          List1.AddItem "Valores de Terceros"
          List1.AddItem l1
          List1.AddItem "Num.Int.   Num.Ch.     Banco                              Importe  Entregado por "
          List1.AddItem l1
          pasada = 1
       End If
       ni = "[" & Format$(rs("num_interno"), "00000000") & "]"
       nch = Format$(rs("num_cheque"), "0000000000")
       RSet i = Format$(rs("importe"), "######0.00")
       b = Format$(Left$(rs("banco"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
       ep = Format$(Left$(rs("origen"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")

       List1.AddItem ni & " " & nch & "  " & b & "  " & i & "  " & ep
       rs.MoveNext
     Wend
     Set rs = Nothing
     If pasada = 1 Then
        List1.AddItem ""
        List1.AddItem ""
     End If

End Sub

Sub MOVCHP()
     i = Space$(10)
     q = "Select * from cyb_02, CYB_01 where [num_mov_banco] = " & Val(t_numint) & " AND [ID_BANCO] = [ID_FORMA_PAGO]"
     Set rs = New ADODB.Recordset
     rs.Open q, cn1
     pasada = 0
     While Not rs.EOF
       If pasada = 0 Then
          List1.AddItem "Valores ROPIOS"
          List1.AddItem l1
          List1.AddItem "Num.Ch.     Banco                              Importe  "
          List1.AddItem l1
          pasada = 1
       End If
       nch = Format$(rs("num_cheque"), "0000000000")
       RSet i = Format$(rs("importe"), "######0.00")
       b = Format$(Left$(rs("DESCRIPCION"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")

       List1.AddItem nch & "  " & b & "  " & i
       rs.MoveNext
     Wend
     Set rs = Nothing
     
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Private Sub Form_Load()

Call barraesag(Me)


End Sub



Private Sub List1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[F7] Imprime - [F8] Borra Comp. - [ESC] Termina "

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
 'borrar mov.
 
 J = MsgBox("Confirma Borrar Movimiento " & t_numint, 4)
 If J = 6 Then
    ni = Val(t_numint)
    Set cl_banco = New bancos
    cl_banco.borrar (ni)
    Set cl_banco = Nothing
 End If
    
End If

If KeyCode = vbKeyF7 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    k = 0
    Printer.FontName = "Courier New"
    Printer.FontSize = 9
    While k <= List1.ListCount - 1
     Printer.Print List1.List(k)
     k = k + 1
    Wend
    Printer.EndDoc
  End If
End If


End Sub
