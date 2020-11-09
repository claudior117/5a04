VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form exp_prodreintegro 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PRODUCTOS(COMPRAS) PARA REINTEGRO"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registros"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   7200
      Width           =   1335
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   600
         Width           =   6015
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1455
      Left            =   8040
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.TextBox t_numop 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Num. OP:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Factura:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Embarque:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "exp002.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "exp002.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   2
      Top             =   8415
      Width           =   12240
      _ExtentX        =   21590
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
            TextSave        =   "09:40"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9340
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3840
      TabIndex        =   19
      Top             =   7320
      Width           =   4335
   End
End
Attribute VB_Name = "exp_prodreintegro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub carga()
 espere.Show
 espere.Label1 = "Espere mientras se carga Reintegro.."
 espere.Refresh
 
  Call armagrid
 
  
 'busco el producto en las ventas
  q = "select * from exp01 where [num_exp] = " & c_vend.ItemData(c_vend.ListIndex)  '& ' " and exp01.[num_exp] = exp02.[num_exp]"
  'q = q & " order by [renglon]"
  Set rs2 = New ADODB.Recordset
  rs2.Open q, cn1
  ct = 0
  IT = 0
  If Not rs2.EOF And Not rs2.BOF Then
          t_fecha = rs2("fecha_embarque")
          t_fecha2 = rs2("fecha_fact")
          t_numop = rs2("num_exp")
          If rs2("estado") = "T" Then
            Label6 = "TERMINADA"
          Else
            Label6 = "EN PROCESO"
          End If
          q = "select * from exp02, a1 where [num_exp] = " & c_vend.ItemData(c_vend.ListIndex) & " and exp02.[id_proveedor] = a1.[id_proveedor]"
          q = q & " order by [num_int_c], [renglon_c], [renglon]"
          Set rs3 = New ADODB.Recordset
          rs3.Open q, cn1
          nic = 0
          txc = 0
          While Not rs3.EOF
            If nic = 0 Then
               nic = rs3("num_int_c")
            End If
                         
            q = "select * from a5, a6 where a5.[num_int] = a6.[num_int] and a5.[num_int] = " & rs3("num_int_c") & " and [renglon] = " & rs3("renglon_c")
            Set rs4 = New ADODB.Recordset
            rs4.MaxRecords = 1
            rs4.Open q, cn1
            If Not rs4.EOF And Not rs4.BOF Then
              sc = Format$(rs4("subtotal"), "#####0.00")
              ic = Format$(rs4("iva"), "#####0.00")
              ivp = Format$((rs4("pu") * rs4("tasa_iva") / 100) * rs3("cantidad"), "#####0.00")
              
              If nic <> rs3("num_int_c") Then
                'mostrar
               
               
               
               m = Format$(txc, "#####0.00")
               nic = rs3("num_int_c")
               txc = ivp
               msf1.TextMatrix(msf1.Rows - 1, 16) = m
               msf1.TextMatrix(msf1.Rows - 1, 17) = F
               m = ""
              Else
                txc = txc + Val(ivp)
                m = ""
              End If
              
              
              If rs4("estado_pago") <> "N" Then
                 F = sacafechaultimopago(rs4("a5.num_int"))
                 If F = "01/01/2000" Then
                    F = " "
                 End If
              Else
                 F = " "
              End If
      
              
            Else
              sc = 0
              ic = 0
              ivp = 0
               F = " "
            End If
            
            
            
            Set rs4 = Nothing
            r = rs3("renglon")
            
            msf1.AddItem r & Chr(9) & rs3("id_producto") & Chr(9) & rs3("producto") & Chr(9) & rs3("cantidad") & Chr(9) & rs3("unidad") & Chr(9) & rs3("pusiva") & Chr(9) & rs3("pusiva") * rs3("cantidad") & Chr$(9) & rs3("operacion_c") & Chr$(9) & rs3("denominacion") & Chr$(9) & rs3("fecha_compra") & Chr$(9) & rs3("Obs") & Chr$(9) & rs3("num_int_c") & Chr$(9) & rs3("renglon_c") & Chr$(9) & sc & Chr$(9) & ic & Chr$(9) & ivp '& Chr$(9) & m & Chr$(9) & f
            ct = ct + rs3("cantidad")
            IT = IT + (rs3("pusiva") * rs3("cantidad"))
            rs3.MoveNext
          Wend
    
          m = Format$(txc, "#####0.00")
          msf1.TextMatrix(msf1.Rows - 1, 16) = m
          msf1.TextMatrix(msf1.Rows - 1, 17) = F
          Set rs3 = Nothing
          msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "----------------------------" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "----------------------------"
          msf1.AddItem "" & Chr(9) & "" & Chr(9) & "Totales ------>" & Chr(9) & Format$(ct, "######0.00") & Chr(9) & "" & Chr(9) & "" & Chr(9) & Format$(IT, "######0.00")
          
  End If
  Set rs2 = Nothing
  
 Unload espere
End Sub
Private Sub btnacepta_Click()
If c_prov.ListIndex > 0 And c_vend.ListIndex > 0 Then
   Call carga
Else
   MsgBox ("Seleccione Operacion de Exportacion")
End If
  
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub






Private Sub c_prov_LostFocus()
 
 Call carga_exportaciones(c_vend)
End Sub

Private Sub c_vend_LostFocus()
If c_vend.ListIndex < 0 Then
  c_vend.ListIndex = 0
Else
  If c_vend.ListIndex > 0 Then
     Call carga
  End If
End If

End Sub

Private Sub Command1_Click()
exp_lista.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub carga_exportaciones(c As ComboBox)
If c_prov.ListIndex > 0 Then
  Set rs = New ADODB.Recordset
  q = "select [num_exp], [detalle] from exp01 where [id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  rs.Open q, cn1
  Call llena_combo(rs, "detalle", "num_exp", c, True)
  c.AddItem "<Seleccionar Operacion exportacion>", 0
  c.ListIndex = 0
  Set rs = Nothing
  Unload espere
Else
  c.clear
  c.AddItem "<Seleccionar Operacion exportacion>", 0
  c.ListIndex = 0
End If
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 18
msf1.ColWidth(0) = 800
msf1.ColWidth(1) = 800
msf1.ColWidth(2) = 3000
msf1.ColWidth(3) = 1000
msf1.ColWidth(4) = 800
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 2200
msf1.ColWidth(8) = 2000
msf1.ColWidth(9) = 1000
msf1.ColWidth(10) = 2000
msf1.ColWidth(11) = 0
msf1.ColWidth(12) = 0
msf1.ColWidth(13) = 1000
msf1.ColWidth(14) = 1000
msf1.ColWidth(15) = 1000
msf1.ColWidth(16) = 1000
msf1.ColWidth(17) = 1000
msf1.TextMatrix(0, 0) = "Renglon"
msf1.TextMatrix(0, 1) = "Id.Prod"
msf1.TextMatrix(0, 2) = "Producto"
msf1.TextMatrix(0, 3) = "Cant."
msf1.TextMatrix(0, 4) = "Unidad"
msf1.TextMatrix(0, 5) = "Pu s/Iva"
msf1.TextMatrix(0, 6) = "Total"
msf1.TextMatrix(0, 7) = "Op. Compra"
msf1.TextMatrix(0, 8) = "Proveedor"
msf1.TextMatrix(0, 9) = "Fecha Comp."
msf1.TextMatrix(0, 10) = "Obs."
msf1.TextMatrix(0, 11) = "Num.Int.Compra"
msf1.TextMatrix(0, 12) = "Renglon Compra"
msf1.TextMatrix(0, 13) = "Subtotal Comp."
msf1.TextMatrix(0, 14) = "Iva Comp."
msf1.TextMatrix(0, 15) = "Iva Producto"
msf1.TextMatrix(0, 16) = "Iva x Comp."
msf1.TextMatrix(0, 17) = "Fec.Pago"




For i = 0 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 6
  msf1.ColAlignment(i) = 9 'der
Next i
For i = 7 To 10
  msf1.ColAlignment(i) = 1 'izq
Next i


End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call armagrid
Call carga_clientes(c_prov)
c_prov.AddItem "<Seleccionar Cliente>", 0
c_prov.ListIndex = 0

Call carga_exportaciones(c_vend)
c_vend.AddItem "<Seleccionar Operacion>", 0
c_vend.ListIndex = 0


Option1 = True

Load exp_productos

End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload exp_productos
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Agrega - [F8] Borra - [F7] Imprime - [F11] Excel -  "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyInsert Then
 If Val(t_numop) > 0 And Label6 = "EN PROCESO" Then
   exp_productos.Show
 Else
   MsgBox ("Seleccione una Operacion de Exportacion EN PROCESO  antes de agregar productos para reintegro")
 End If
End If


If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 1
    c(1) = 2
    c(2) = 3
    c(3) = 4
    c(4) = 5
    c(5) = 6
    c(6) = 7
    c(7) = 8
    c(8) = 9
    c(9) = 10
    For i = 10 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "PLANILLA PARA REITEGRO DE EXPORTACIONES", "Cliente: " & c_prov, "Nro. Op: [" & Format$(t_numop, "00000") & "]  Operacion: " & c_vend, "Embarque: " & t_fecha, 50, 8, True, False, "H")
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyF8 Then
  If msf1.Rows > 1 And Label6 = "EN PROCESO" Then
    J = MsgBox("Confirma eliminar renglon " & msf1.TextMatrix(msf1.Row, 0), 4)
    If J = 6 Then
      r = msf1.Row
      On Error GoTo errb
      cn1.BeginTrans
          QUERY = "update a6 set  [exportacion]=[exportacion] - " & Val(msf1.TextMatrix(r, 3))
          QUERY = QUERY & " where [num_int]= " & Val(msf1.TextMatrix(r, 11)) & " and [renglon]= " & Val(msf1.TextMatrix(r, 12))
          cn1.Execute QUERY
   
      
      QUERY = "DELETE FROM exp02 WHERE [num_exp] = " & Val(t_numop) & " and [renglon] = " & msf1.TextMatrix(r, 0)
      
      cn1.Execute QUERY
      cn1.CommitTrans
      Call carga
      
    End If
  Else
    MsgBox ("Operacion Inexistente o Terminada")
  End If
End If

Exit Sub
errb:
MsgBox ("No se pudo eliminar el item del registro de exportacion")
Exit Sub
End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"


End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yyyy")
  End If
End If

End Sub
