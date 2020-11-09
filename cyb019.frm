VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_carterach2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cartera Cheque a la fecha"
   ClientHeight    =   8745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8745
   ScaleWidth      =   12150
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por"
      Height          =   615
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Interno"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Difereida"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Caretera Cheques de Terceros a la Fecha"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.TextBox t_fecha5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Height          =   6015
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   5535
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9763
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb019.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cyb019.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salir sin Modificar"
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
      Top             =   8490
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "30/08/2013"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "13:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "cyb_carterach2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call carga
msf1.SetFocus
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub carga()
Call armagrid
Dim i As String
i = Space$(10)

'Call inicia
Set rs = New ADODB.Recordset

q = "select * from cyb_03  "
c = " where "


If t_fecha5 <> "" And IsDate(t_fecha5) Then
  q = q & c & " datevalue([fecha_ingreso]) <  datevalue('" & t_fecha5 & "')"
  c = " and "
 ' q = q & c & " datevalue([fecha_salida]) >  datevalue('" & t_fecha5 & "')"
 ' c = " and "
  'la caretra de cheque a un dia son todos los cheques que entraron ante de ese dia(incluido)
  'ý que ademas salieron despues de ese dia sin incluir
Else
 Exit Sub

End If

If Option1 = True Then
   q = q & " order by [num_interno]"
Else
   q = q & " order by [fecha_dif]"
End If

rs.Open q, cn1
t = 0
c = 0
While Not rs.EOF
  m = 0
  If rs("estado") = "C" Then
    m = 1
  Else
    If DateValue(rs("fecha_salida")) > DateValue(t_fecha5) Then
        m = 1
    End If
  
  End If
  If m = 1 Then
     Label6 = c
     Label6.Refresh
     fs = " "
     nop = " "
     fs = rs("fecha_salida")
     If rs("num_int_op") > 0 Then
         Set rs2 = New ADODB.Recordset
         q = "select * from a5 where [num_int] = " & rs("num_int_op") & " and [id_tipocomp] = 50 "
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           nop = Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
         Else
           nop = " "
         End If
       Else
        nop = " "
     End If
     RSet i = Format$(rs("importe"), "######0.00")
     t = t + Val(i)
     c = c + 1
     Set rs2 = New ADODB.Recordset
     q = "select * from cyb_05 where [num_int_ch_terc] = " & rs("num_interno")
     rs2.Open q, cn1
     If Not rs2.EOF And Not rs2.BOF Then
      ope = rs2("operacion")
     Else
      ope = "Sin Detallar"
     End If
     Set rs2 = Nothing
     msf1.AddItem Format$(rs("num_interno"), "00000") & Chr$(9) & Format$(rs("num_cheque"), "0000000000") & Chr$(9) & Format$(rs("fecha_dif"), "dd/mm/yyyy") & Chr$(9) & rs("banco") & Chr$(9) & rs("origen") & Chr$(9) & rs("destino") & Chr$(9) & i & Chr$(9) & rs("estado") & Chr$(9) & rs("sucursal") & Chr$(9) & rs("titular") & Chr$(9) & rs("fecha_emision") & Chr$(9) & fs & Chr$(9) & nop & Chr$(9) & ope
   End If
   rs.MoveNext
Wend
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "___________________" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & ""
msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & "Cheques: " & c & Chr$(9) & "" & Chr$(9) & "" & Chr$(9) & Format$(t, "########0.00")
          
Set rs = Nothing

End Sub

Private Sub Form_Load()
Call INICIALIZA2(Me)
Call barraesag(Me)
Option2 = True
Call armagrid


End Sub

  
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 15
msf1.ColWidth(0) = 900
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 2400
msf1.ColWidth(5) = 2400
msf1.ColWidth(6) = 1200
msf1.ColWidth(7) = 600
msf1.ColWidth(8) = 1200
msf1.ColWidth(9) = 2000
msf1.ColWidth(10) = 1100
msf1.ColWidth(11) = 1100
msf1.ColWidth(12) = 1100
msf1.ColWidth(13) = 1400
msf1.ColWidth(14) = 2000
msf1.TextMatrix(0, 0) = "Num.Int."
msf1.TextMatrix(0, 1) = "Num.Ch."
msf1.TextMatrix(0, 2) = "Fecha Dif."
msf1.TextMatrix(0, 3) = "Banco"
msf1.TextMatrix(0, 4) = "Origen"
msf1.TextMatrix(0, 5) = "Destino"
msf1.TextMatrix(0, 6) = "Importe"
msf1.TextMatrix(0, 7) = "Estado"
msf1.TextMatrix(0, 8) = "Sucursal"
msf1.TextMatrix(0, 9) = "Titular"
msf1.TextMatrix(0, 10) = "Fecha Emis."
msf1.TextMatrix(0, 10) = "Fecha Entrada"
msf1.TextMatrix(0, 11) = "Fecha Salida"
msf1.TextMatrix(0, 12) = "Nro.O.P"
msf1.TextMatrix(0, 13) = "Op. Entrada"
For i = 0 To 13
    msf1.ColAlignment(i) = 1
Next i
msf1.ColAlignment(6) = 7




End Sub






Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = " [F7] Imprime -[F11] Excel "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 12
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    c(6) = 5
    c(7) = 6
    c(8) = 7
    c(9) = 8
    c(10) = 9
    c(11) = 10
    c(12) = 11
    

    For i = 13 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "CARTERA CH. TERCERO", "", "Al...: " & t_fecha6, " ", 55, 7, True, False, "H")
  End If
    
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

If KeyCode = vbKeyInsert Then
 Call nivel_acceso(4)
 If para.id_grupo_modulo_actual > 8 Then
  op_fp1_1.t_funcion = "A"
  op_fp1_1.Show
  
 Else
  Call sinpermisos
 End If
End If

If KeyCode = vbKeyF8 Then
 Call nivel_acceso(4)
 If para.id_grupo_modulo_actual > 8 Then
   'borra ch
   J = MsgBox("Atencion. La eliminacion del cheque implica cambios en la CAJA pero No se eliminan mov. asociados como OP o RBO. Confirma Eliminar Cheque Nro. Interno : " & msf1.TextMatrix(msf1.Row, 0), 4)
   If J = 6 Then
       QUERY = "DELETE FROM cyb_03 WHERE [num_interno] = " & Val(msf1.TextMatrix(msf1.Row, 0))
       cn1.BeginTrans
       cn1.Execute QUERY
    
       QUERY = "DELETE FROM cyb_05 WHERE [num_int_ch_terc] = " & Val(msf1.TextMatrix(msf1.Row, 0))
       cn1.Execute QUERY
    
       cn1.CommitTrans
    
    
    Call carga
    msf1.SetFocus

   
   End If
 Else
  Call sinpermisos
 End If
End If



End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call nivel_acceso(4)
  If para.id_grupo_modulo_actual > 8 Then
    Set cl_chterc = New chterceros
    cl_chterc.cargar (Val(msf1.TextMatrix(msf1.Row, 0)))
    If cl_chterc.numinterno > 0 And (cl_chterc.tiposalida = "M" Or cl_chterc.tiposalida = "C") Then
        Load cyb_salidachterc
        cyb_salidachterc.t_NUMINT = Val(msf1.TextMatrix(msf1.Row, 0))
        cyb_salidachterc.T_NUMCH = msf1.TextMatrix(msf1.Row, 1)
        cyb_salidachterc.t_fechad = msf1.TextMatrix(msf1.Row, 2)
        cyb_salidachterc.t_banco = msf1.TextMatrix(msf1.Row, 3)
        cyb_salidachterc.t_origen = msf1.TextMatrix(msf1.Row, 4)
        cyb_salidachterc.t_destino = " "
        cyb_salidachterc.t_importe = msf1.TextMatrix(msf1.Row, 6)
        cyb_salidachterc.t_sucursal = msf1.TextMatrix(msf1.Row, 8)
        cyb_salidachterc.t_titular = msf1.TextMatrix(msf1.Row, 9)
        cyb_salidachterc.t_fechae = msf1.TextMatrix(msf1.Row, 10)
        cyb_salidachterc.c_tipo.ListIndex = 0
        cyb_salidachterc.c_cuenta.ListIndex = 0
            
        If cl_chterc.tiposalida = "M" Then
            'solo se editan los ch. salida manual
            cyb_salidachterc.t_destino = msf1.TextMatrix(msf1.Row, 5)
            cyb_salidachterc.t_fecha = msf1.TextMatrix(msf1.Row, 11)
            cyb_salidachterc.t_funcion = "M"
            
        Else
            'nueva salida
            cyb_salidachterc.t_destino = " "
            cyb_salidachterc.t_fecha = " "
            cyb_salidachterc.t_funcion = "A"
        End If
        cyb_salidachterc.Show
    Else
        MsgBox ("Solo se pueden editar cheques en Cartera o con Salida Manual")
    End If
     Set cl_chterc = Nothing
  Else
    Call sinpermisos
  End If
End If
End Sub


Private Sub t_fecha5_GotFocus()
t_fecha5 = ""
End Sub

Private Sub t_fecha5_LostFocus()
If Not IsDate(t_fecha5) Then
  t_fecha5 = ""
End If
End Sub

Private Sub t_fecha6_GotFocus()
t_fecha6 = ""
End Sub
Sub inicia()

End Sub
Private Sub t_fecha6_LostFocus()
If Not IsDate(t_fecha6) Then
  t_fecha6 = ""
End If

End Sub
