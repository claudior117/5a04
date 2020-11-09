VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_carterach 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE CHEQUES DE TERCERO"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8730
   ScaleWidth      =   12165
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Operacion (Ent. y Sal. Ch. del Sistema"
      Height          =   975
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   5055
      Begin VB.TextBox t_fecha7 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   34
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha8 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   33
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha5 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha6 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00800000&
         Caption         =   "Entrada(Desde-Hasta):"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Salida(Desde-Hasta):"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0C0C0&
      Height          =   4455
      Left            =   120
      TabIndex        =   25
      Top             =   2760
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid msf1 
         Height          =   4095
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado"
      Height          =   615
      Left            =   6360
      TabIndex        =   22
      Top             =   1920
      Width           =   3375
      Begin VB.ComboBox c_estados 
         Height          =   315
         ItemData        =   "cyb005.frx":0000
         Left            =   1440
         List            =   "cyb005.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "Estado Actual:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenado por"
      Height          =   615
      Left            =   6360
      TabIndex        =   19
      Top             =   1200
      Width           =   4215
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Difereida"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Interno"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   10200
      TabIndex        =   16
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb005.frx":0083
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cyb005.frx":0905
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   6360
      TabIndex        =   11
      Top             =   120
      Width           =   4215
      Begin VB.TextBox t_cedido 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   15
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox t_entregado 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Cedido a :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Entregado por"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Diferida Cheque"
      Height          =   975
      Left            =   3360
      TabIndex        =   6
      Top             =   1560
      Width           =   2895
      Begin VB.TextBox t_fecha3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox t_fecha4 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   7
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800000&
         Caption         =   "Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Emision Cheque"
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   3015
      Begin VB.TextBox t_fecha2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800000&
         Caption         =   "Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Caption         =   "Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8475
      Width           =   12165
      _ExtentX        =   21458
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
            TextSave        =   "13:43"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   10200
      TabIndex        =   27
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "cyb_carterach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Private Sub btnacepta_Click()
Call carga
msf1.SetFocus
End Sub
Sub imprime()


End Sub
Private Sub btnsale_Click()
Unload Me
End Sub


Sub inicia()
List1.clear
l1 = "============================================================================================================="
List1.AddItem "Cartera de Valores de Terceros"
List1.AddItem ""
List1.AddItem "Fecha de Ingreso entre:      "
List1.AddItem "Fecha de Acreditacion entre: "
List1.AddItem ""
List1.AddItem l1
List1.AddItem "Num.Int Num.Ch.     Fecha Dif.  Banco            Entregado por    Cedido a         Importe"
List1.AddItem l1
End Sub



Private Sub carga()
Call armagrid
Dim i As String
i = Space$(10)

'Call inicia
Set rs = New ADODB.Recordset

q = "select * from cyb_03  "
c = " where "

If t_fecha1 <> "" And IsDate(t_fecha1) Then
  q = q & c & " datevalue([fecha_emision]) >= datevalue('" & t_fecha1 & "')"
  c = " and "
End If

If t_fecha2 <> "" And IsDate(t_fecha2) Then
  q = q & c & " datevalue([fecha_emision]) <= datevalue('" & t_fecha2 & "')"
  c = " and "
End If

If t_fecha3 <> "" And IsDate(t_fecha3) Then
  q = q & c & " datevalue([fecha_dif]) >= datevalue('" & t_fecha3 & "')"
  c = " and "
End If

If t_fecha4 <> "" And IsDate(t_fecha4) Then
  q = q & c & " datevalue([fecha_dif]) <= datevalue('" & t_fecha4 & "')"
  c = " and "
End If

If t_fecha5 <> "" And IsDate(t_fecha5) Then
  q = q & c & " datevalue([fecha_ingreso]) >= datevalue('" & t_fecha5 & "')"
  c = " and "
End If

If t_fecha6 <> "" And IsDate(t_fecha6) Then
  q = q & c & " datevalue([fecha_ingreso]) <= datevalue('" & t_fecha6 & "')"
  c = " and "
End If

If t_fecha7 <> "" And IsDate(t_fecha7) Then
  q = q & c & " datevalue([fecha_salida]) >= datevalue('" & t_fecha7 & "')"
  c = " and "
End If

If t_fecha8 <> "" And IsDate(t_fecha8) Then
  q = q & c & " datevalue([fecha_salida]) <= datevalue('" & t_fecha8 & "')"
  c = " and "
End If

If t_entregado <> "" Then
  q = q & c & " [origen] like '%" & t_entregado & "%'"
  c = " and "
End If

If t_cedido <> "" Then
  q = q & c & " [destino] like '%" & t_cedido & "%'"
  c = " and "
End If

If c_estados.ListIndex > 0 Then
   q = q & c & " [estado] = '" & Mid$(c_estados, 1, 1) & "'"
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
     Label6 = c
     Label6.Refresh
     If rs("estado") = "C" Then
       fs = " "
       nop = " "
     Else
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
c_estados.ListIndex = 3
Call armagrid
Load op_fp1_1
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

  







Private Sub Form_Unload(Cancel As Integer)
Unload op_fp1_1
End Sub

Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.Item(2) = "[INS] Ingresa Ch. - [F8] Elimina - [F7] Imprime - [ENTER] Salida Ch.[F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
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
    c(7) = 7
    c(8) = 11
    c(9) = 12
    c(10) = 13

    For i = 14 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "INFORME CH. TERCERO", "", "Periodo...: " & t_fecha1 & " " & t_fecha2, "Estado..:" & c_estados, 55, 7, True, False, "H")
  End If
    
End If

If KeyCode = vbKeyInsert Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 7 Then
  op_fp1_1.t_funcion = "A"
  op_fp1_1.Show
  
 Else
  Call sinpermisos
 End If
End If

If KeyCode = vbKeyF8 Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 8 Then
   'borra ch
   J = MsgBox("Confirma Eliminar Cheque Nro. Interno : " & msf1.TextMatrix(msf1.Row, 0), 4)
   If J = 6 Then
       Set rs = New ADODB.Recordset
       q = "select * from cyb_03 where [num_interno] = " & Val(msf1.TextMatrix(msf1.Row, 0))
       rs.Open q, cn1
       If Not rs.EOF And Not rs.BOF Then
         If estadocaja(rs("fecha_ingreso")) = "A" Then
           If rs("num_int_rbo") = 0 Then
            If rs("Estado") = "C" Then
             QUERY = "DELETE FROM cyb_03 WHERE [num_interno] = " & Val(msf1.TextMatrix(msf1.Row, 0))
             cn1.BeginTrans
             cn1.Execute QUERY
    
             QUERY = "DELETE FROM cyb_05 WHERE [num_int_ch_terc] = " & Val(msf1.TextMatrix(msf1.Row, 0))
             cn1.Execute QUERY
    
             cn1.CommitTrans
             
             Call borracontabilidad(Val(msf1.TextMatrix(msf1.Row, 0)), "H")
            Else
             MsgBox ("Imposible borrar cheques que no esten en cartera")
            End If
           Else
             MsgBox ("Este valor fue ingresado por Recibos, debe eliminar el documento para darlo de baja")
           End If
         Else
           MsgBox ("Caja Cerrada. Imposible realizar operacion")
         End If
       End If
       Call carga
       msf1.SetFocus
   End If
 Else
  Call sinpermisos
 End If
End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If


End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call nivel_acceso(4)
  If para.id_grupo_modulo_actual >= 7 Then
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

