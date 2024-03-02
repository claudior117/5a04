VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cc_detalle 
   BackColor       =   &H00E0E0E0&
   Caption         =   "COMPROBANTE DE COMPRA"
   ClientHeight    =   9435
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   17760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   17760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Archivo Electronico"
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   8040
      Width           =   13335
      Begin VB.CommandButton Command1 
         Caption         =   "Ver o Modificar"
         Height          =   375
         Left            =   11400
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_path 
         Height          =   405
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   10935
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6180
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   17295
   End
   Begin VB.Frame CUIT 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   1335
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   17295
      Begin VB.TextBox t_numint 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   15240
         MaxLength       =   10
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox t_tipocomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14400
         MaxLength       =   6
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10800
         MaxLength       =   8
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox t_letra 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8880
         MaxLength       =   6
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox t_sucursal 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9720
         MaxLength       =   6
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox t_prov 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   960
         MaxLength       =   81
         TabIndex        =   2
         Top             =   720
         Width           =   7095
      End
      Begin VB.TextBox t_idprov 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800080&
         Caption         =   "Numero Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   15240
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800080&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   14400
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800080&
         Caption         =   "Comprobante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8880
         TabIndex        =   11
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   15600
      TabIndex        =   6
      Top             =   7920
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "Proc004.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Proc004.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   5
      Top             =   9180
      Width           =   17760
      _ExtentX        =   31327
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cc_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim l1 As String
Dim l2 As String


Private Sub btnsale_Click()
Unload Me
End Sub









Private Sub Command1_Click()
gen_path.t_id = t_numint
gen_path.t_modulo = "Compras"
gen_path.t_origen = "C"
gen_path.t_path = t_path
gen_path.Show

End Sub

Private Sub Form_Activate()
Call carga
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     gen_tools.Show
End Select
End Sub
Sub electronico()
Set rs2 = New ADODB.Recordset
q = "select * from a20 where [num_int] = " & Val(t_numint)
rs2.Open q, cn1
If Not rs2.EOF And Not rs2.BOF Then
  t_path = rs2("path")
Else
 t_path = ""
End If
Set rs2 = Nothing

End Sub
Sub carga()
List1.clear
l1 = "----------------------------------------------------------------------------------------------------------------"
l2 = "*************************************"

If t_numint <> "" Then
  q = "select * from a5, a1, g2, g1 where a5.[id_proveedor] = a1.[id_proveedor] and [id_tipocomp] = [id_tipo_comp] and a5.[id_usuario] = g1.[id_usuario]  and a5.[num_int] = " & Val(t_numint)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.EOF And Not rs.BOF Then
     t_tipocomp = rs("id_tipocomp")
     List1.AddItem Space$(60) & "Numero  :" & rs("ABREVIATURA") & " " & rs("letra") & " " & Format$(rs("sucursal"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
     List1.AddItem Space$(60) & "Fecha   :" & rs("fecha")
     List1.AddItem Space$(60) & "Vto.    :" & rs("fecha_vto")
     List1.AddItem ""
     List1.AddItem "Proveedor  :(" & Format$(rs("a5.id_proveedor"), "00000") & ") " & rs("denominacion")
     List1.AddItem "Direccion  : " & rs("direccion") & " (" & rs("cp") & ") " & rs("localidad")
     List1.AddItem "Cuit       : " & rs("cuit") & Space(10) & " Zona: " & rs("zona")
     List1.AddItem "TE         : " & rs("TE")
     List1.AddItem "Email      : " & rs("email")
     List1.AddItem "Contacto   : " & rs("info_contacto")
     q = "select * from c_02 where [num_mov_int] = " & Val(t_numint) & " and [modulo] = 'C'"
     Set rs1 = New ADODB.Recordset
     rs1.Open q, cn1
     If Not rs1.EOF And Not rs1.BOF Then
       na = Format$(rs1("num_interno"), "0000000000")
     Else
       na = "0000000000"
     End If
     Set rs1 = Nothing
      
     
     Call electronico
     
     Select Case Val(t_tipocomp)
     Case Is = 50
       Call PAGOS
     
       List1.AddItem ""
       List1.AddItem ""
       List1.AddItem "Total a Ingresar en Cuenta prov..:" & Format$((rs("ret_gan") + rs("ret_ib") + rs("total")), "######0.00")
       List1.AddItem "Total Sujeto a Retenciones.......:" & Format$(rs("subtotal"), "######0.00")
       List1.AddItem "Total Retetncion Ganancias.......:" & Format$(rs("ret_gan"), "######0.00")
       List1.AddItem "Total Retetencion Ing. Brutos....:" & Format$(rs("ret_ib"), "######0.00") & "  (" & rs("a5.IVA") & ")"
       List1.AddItem "Total de la Orden de Pago........:" & Format$(rs("total"), "######0.00") & "    U$s " & Format$(rs("total_d"), "######0.00")
       List1.AddItem ""

    Case Is = 65
       Call oc
       List1.AddItem ""
       List1.AddItem ""
       List1.AddItem l2
       List1.AddItem "Totales"
       List1.AddItem l2
     
       List1.AddItem "Subtotal  : " & Format$(rs("subtotal"), "######0.00")
       List1.AddItem "Iva       : " & Format$(rs("a5.iva"), "######0.00")
       List1.AddItem "No Grabado: " & Format$(rs("no_grabado"), "######0.00")
       List1.AddItem "Perc.y Ret: " & Format$(rs("percep_ret"), "######0.00")
       List1.AddItem "       ----------------- "
       If rs("MONEDA") = "P" Then
         List1.AddItem "Total $   : " & Format$(rs("total"), "######0.00")
         List1.AddItem ""
         List1.AddItem "Total U$s : " & Format$(rs("total_d"), "######0.00")
         List1.AddItem ""
       Else
         List1.AddItem "Total U$s : " & Format$(rs("total"), "######0.00")
         List1.AddItem ""
         List1.AddItem "Total $   : " & Format$(rs("total_d"), "######0.00")
         List1.AddItem ""
       End If
     Case Is = 70
       Call cotiz
    
     
     
     Case Else
       Call COMPROBANTES
       List1.AddItem ""
       List1.AddItem ""
       List1.AddItem l2
       List1.AddItem "Totales"
       List1.AddItem l2
     
       List1.AddItem "Subtotal  : " & Format$(rs("subtotal"), "######0.00")
       List1.AddItem "Iva       : " & Format$(rs("a5.iva"), "######0.00")
       List1.AddItem "No Grabado: " & Format$(rs("no_grabado"), "######0.00")
       List1.AddItem "Perc.y Ret: " & Format$(rs("percep_ret"), "######0.00")
       List1.AddItem "       ----------------- "
       If rs("MONEDA") = "P" Then
         List1.AddItem "Total $   : " & Format$(rs("total"), "######0.00")
         List1.AddItem ""
         List1.AddItem "Total U$s : " & Format$(rs("total_d"), "######0.00")
         List1.AddItem ""
       Else
         List1.AddItem "Total U$s : " & Format$(rs("total"), "######0.00")
         List1.AddItem ""
         List1.AddItem "Total $   : " & Format$(rs("total_d"), "######0.00")
         List1.AddItem ""
       End If
     
       List1.AddItem "Detalle IVA"
       List1.AddItem "-------------------------------------------------------------"
       List1.AddItem "Neto           Alicuota        Iva    "
       List1.AddItem "-------------------------------------------------------------"
       q = "select * from a23 where [num_int] = " & rs("num_int")
       Set rs3 = New ADODB.Recordset
       rs3.Open q, cn1
       n = Space$(10)
       a = Space$(10)
       v = Space$(10)
       While Not rs3.EOF
           RSet n = Format$(rs3("neto"), "######0.00")
           RSet a = Format$(rs3("tasa_iva"), "######0.00")
           RSet v = Format$(rs3("iva"), "######0.00")
           List1.AddItem n & "  " & a & "%    " & v
           
           rs3.MoveNext
       Wend
       Set rs3 = Nothing
       
       List1.AddItem ""
       List1.AddItem ""
       
       If rs("id_tipocomp") = 1 Then
         e = Space$(10)
         q = "select * from a15, a5 where [num_int_comp] = " & rs("num_int") & " and [num_int_op] = [num_int]"
         Set rs3 = New ADODB.Recordset
         rs3.Open q, cn1
             If rs("moneda") = "P" Then
               t = rs("total")
             Else
               t = rs("total_d")
             End If
             RSet e = Format$(t, "#####0.00")
             List1.AddItem "Cancelacion en $"
             List1.AddItem "-------------------------------------------------------------"
             List1.AddItem "Fecha       O.P.                   Importe         Saldo    "
             List1.AddItem "                                   Cancelado     Pendiente"
             List1.AddItem "-------------------------------------------------------------"
             List1.AddItem "            Deuda Original                    " & e
         s = Space$(10)
         While Not rs3.EOF
           F = rs3("fecha")
           r = Format$(rs3("sucursal"), "0000") & "-" & Format$(rs3("num_comprobante"), "00000000")
           RSet e = Format$(rs3("importe_pagado"), "######0.00")
           RSet s = Format$(rs3("saldo_comprobante") - rs3("importe_pagado"), "######0.00")
           List1.AddItem F & "  " & r & "    " & e & "       " & s
           rs3.MoveNext
         Wend
         Set rs3 = Nothing
     End If
     List1.AddItem " "
     List1.AddItem " "
     
     End Select
     List1.AddItem "Condiciones: " & rs("condiciones")
     List1.AddItem ""
     List1.AddItem "Observaciones: " & rs("obs")
     List1.AddItem ""
     List1.AddItem "SALDO IMPAGO: " & Format$(rs("saldo_impago"), "######0.00")
     
     List1.AddItem Space$(60) & "***** DATOS DE AUDITORIA *****"
     List1.AddItem ""
     List1.AddItem Space$(56) & "Cod.ret.Gan...: " & rs("a5.id_codretgan")
     List1.AddItem Space$(56) & "Cod.ret.IB....: " & rs("a5.id_codretib")
     List1.AddItem Space$(60) & "Esperado......:" & rs("fecha_prob_entrega")
     List1.AddItem Space$(60) & "Estado........:" & rs("estado")
     List1.AddItem Space$(60) & "Cta.Cte.......:" & rs("a5.ctacte")
     List1.AddItem Space$(60) & "Stock.........:" & rs("a5.stock")
     List1.AddItem Space$(60) & "Iva...........:" & rs("grabado")
     List1.AddItem Space$(60) & "Num.Int.......:" & rs("num_int")
     If rs("id_cuenta") > 0 Then
       Set rs2 = New ADODB.Recordset
       q = "SELECT [descripcion] FROM C_01 WHERe [id_cuenta] = " & rs("id_cuenta")
       rs2.MaxRecords = 1
       rs2.Open q, cn1
       If Not rs2.EOF And Not rs2.BOF Then
          cuenta = rs2("descripcion")
       Else
          cuenta = "Inexistente"
       End If
       List1.AddItem Space$(60) & "Cuenta........:" & rs("id_cuenta") & " - " & cuenta
       Set rs2 = Nothing
     End If
     List1.AddItem Space$(60) & "Usuario.......:" & rs("usuario")
     List1.AddItem Space$(60) & "Moneda........:" & rs("Moneda")
     List1.AddItem Space$(60) & "Cotizacion....:" & rs("cotiz_dolar")
     List1.AddItem Space$(56) & "Estado Pago...:" & rs("estado_pago") & "  " & rs("num_op")
     List1.AddItem Space$(56) & "Asiento Int...: " & na
     

 
 End If
 Set rs = Nothing
End If
End Sub
Sub COMPROBANTES()
     List1.AddItem ""
     List1.AddItem l1
     List1.AddItem "Id.   Detalle                                           Cant.  Unid. Env.  PU     %Iva   %Dto  PU c/dto   Importe"
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     c$ = Space$(8)
     p$ = Space$(8)
     i$ = Space$(9)
     v$ = Space$(4)
     dt$ = Space$(5)
     psd$ = Space$(8)
     ev$ = Space$(4)
     While Not rs1.EOF
         RSet c$ = Format$(rs1("cantidad"), "####0.00")
         RSet p$ = Format$(rs1("pu"), "####0.00")
         RSet psd$ = Format$(rs1("pusindto"), "####0.00")
         b = Format$(rs1("id_producto"), "00000")
         d = Format$(Left$(rs1("detalle"), 47), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
         RSet i$ = Format$(rs1("importe"), "#####0.00")
         RSet v$ = Format$(rs1("tasa_iva"), "#0.0")
         RSet dt$ = Format$(rs1("descuento"), "#0.00")
         u = Format$(rs1("unidad06"), "@@@@@!")
         ev$ = Format$(rs1("envase"), "###0")
         List1.AddItem b & " " & d & " " & c$ & " " & u & " " & ev$ & " " & psd$ & " (" & v$ & ") " & dt$ & " " & p$ & " " & i$
         
         rs1.MoveNext
     Wend
     Set rs1 = Nothing

     List1.AddItem ""
     List1.AddItem l2
     List1.AddItem "Percepciones"
     List1.AddItem l2
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a13, a12 where [num_int] = " & Val(t_numint) & " and a13.[id_percepcion] = a12.[id_percepcion]"
     rs1.Open q, cn1
     While Not rs1.EOF
         p = Format$(Left$(rs1("descripcion"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         i = Format$(rs1("importe"), "######0.00")
         List1.AddItem p & "   " & i
         rs1.MoveNext
     Wend
     Set rs1 = Nothing


End Sub


Sub oc()
  Dim a(5) As String
     l22 = l1 & "---------------------------"
     List1.FontSize = 8
     List1.AddItem ""
     List1.AddItem l22
     List1.AddItem "Id.   Detalle                                  Obs.                     Cant.   PU   %Iva  Unidad    Importe Obra"
     List1.AddItem l22
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     c$ = Space$(8)
     p$ = Space$(8)
     i$ = Space$(9)
     v$ = Space$(4)
     dt$ = Space$(5)
     While Not rs1.EOF
         RSet c$ = Format$(rs1("cantidad"), "####0.00")
         RSet p$ = Format$(rs1("pu"), "####0.00")
         b = Format$(rs1("id_producto"), "00000")
         d = Left$(Format$(rs1("detalle"), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 40)
         RSet i$ = Format$(rs1("importe"), "#####0.00")
         RSet v$ = Format$(rs1("tasa_iva"), "#0.0")
         dt$ = Format$(rs1("unidad"), "@@@@@")
         o = Left$(Format$(rs1("observaciones"), "@@@@@@@@@@@@@@@@@@@@!"), 20)
         Set rs2 = New ADODB.Recordset
         q = "select * from a4 where [id_obra] = " & rs1("id_obra")
         rs2.MaxRecords = 1
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           obra = Format$(Left$(rs2("descripcion"), 25), "@@@@@@@@@@@@@@@@@@@@@@@@!")
         Else
           obra = "Dada de baja"
         End If
         Set rs2 = Nothing
         
         List1.AddItem b & " " & d & " " & o & " " & c$ & " " & p$ & " (" & v$ & ") " & dt$ & " " & i$ & " " & obra
       
        
        'muestra descripcion extendida
          Set rs2 = New ADODB.Recordset
          q = "select * from a21 where [num_int] = " & Val(t_numint) & " and [renglon] = " & rs1("renglon")
          rs2.Open q, cn1
          If Not rs2.EOF And Not rs2.BOF Then
             'imprimo lineas
             Call lee_desc_extra(a, rs2("descripcion"))
             For k = 0 To 4
              If a(k) <> "%%" Then
               d = Format$(Left$(a(k), 50), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
               List1.AddItem "     " & " " & d
             End If
             Next k
          End If
          Set rs2 = Nothing
          
        
             
         
         rs1.MoveNext
     Wend
     Set rs1 = Nothing

     List1.AddItem ""
     List1.AddItem l2
     List1.AddItem "Percepciones"
     List1.AddItem l2
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a13, a12 where [num_int] = " & Val(t_numint) & " and a13.[id_percepcion] = a12.[id_percepcion]"
     rs1.Open q, cn1
     While Not rs1.EOF
         p = Format$(Left$(rs1("descripcion"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         i = Format$(rs1("importe"), "######0.00")
         List1.AddItem p & "   " & i
         rs1.MoveNext
     Wend
     Set rs1 = Nothing



End Sub

Sub cotiz()
     l22 = l1 & "---------------------------"
     List1.Font.Size = 8
     List1.AddItem ""
     List1.AddItem l22
     List1.AddItem "Id.   Detalle                                  Obs.                   Cant.  Unidad  Obra"
     List1.AddItem l22
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a6 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     c$ = Space$(8)
     p$ = Space$(8)
     i$ = Space$(9)
     v$ = Space$(4)
     dt$ = Space$(5)
     While Not rs1.EOF
         RSet c$ = Format$(rs1("cantidad"), "####0.00")
         b = Format$(rs1("id_producto"), "00000")
         d = Left$(Format$(rs1("detalle"), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!"), 40)
         dt$ = Format$(rs1("unidad"), "@@@@@")
         o = Left$(Format$(rs1("observaciones"), "@@@@@@@@@@@@@@@@@@@@!"), 20)
         Set rs2 = New ADODB.Recordset
         q = "select * from a4 where [id_obra] = " & rs1("id_obra")
         rs2.MaxRecords = 1
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
           obra = Format$(Left$(rs2("descripcion"), 25), "@@@@@@@@@@@@@@@@@@@@@@@@!")
         Else
           obra = "Dada de baja"
         End If
         Set rs2 = Nothing
         
         List1.AddItem b & " " & d & " " & o & " " & c$ & " " & dt$ & " " & obra
         rs1.MoveNext
     Wend
     Set rs1 = Nothing

     List1.AddItem ""
     List1.AddItem l2
     List1.AddItem "Percepciones"
     List1.AddItem l2
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a13, a12 where [num_int] = " & Val(t_numint) & " and a13.[id_percepcion] = a12.[id_percepcion]"
     rs1.Open q, cn1
     While Not rs1.EOF
         p = Format$(Left$(rs1("descripcion"), 30), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
         i = Format$(rs1("importe"), "######0.00")
         List1.AddItem p & "   " & i
         rs1.MoveNext
     Wend
     Set rs1 = Nothing



End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 3)
  'Case Is = 27
  '      Me.Hide
End Select

End Sub

Sub PAGOS()
     List1.FontSize = 9
     l1 = "---------------------------------------------------------------------------------------------------------"
     l2 = "----------------------------------------------"
     List1.AddItem ""
    
     List1.AddItem "COMPROBANTES APLICADOS "
     List1.AddItem l2
     List1.AddItem "Fecha       Comprobante      Importe     Importe"
     List1.AddItem "                             a Cancelar  Cancelado "
     List1.AddItem l2
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a15 where [num_int_op] = " & Val(t_numint) & ""
     rs1.Open q, cn1
     i = Space$(10)
     ia = Space$(10)
     While Not rs1.EOF
         Set rs2 = New ADODB.Recordset
         q = "select * from a5 where [num_int] = " & rs1("num_int_comp")
         rs2.Open q, cn1
         If Not rs2.EOF And Not rs2.BOF Then
          F = Format$(rs2("fecha"), "dd/mm/yyyy")
          nc = Format$(rs2("letra"), ">@") & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comprobante"), "00000000")
          RSet ia = Format$(rs1("importe_pagado"), "######0.00")
          RSet i = Format$(rs1("saldo_comprobante"), "######0.00")
          List1.AddItem F & "  " & nc & "  " & i & "  " & ia
         End If
         Set rs2 = Nothing
         rs1.MoveNext
     Wend
     Set rs1 = Nothing
     
     
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem ""
     List1.AddItem "FORMA PAGO "
     List1.AddItem l1
     List1.AddItem "Forma Pago      Num.Ch.     Fecha dif.  Banco/Detalle             Titular                    Importe"
     List1.AddItem l1
     
     Set rs1 = New ADODB.Recordset
     q = "select * from a7 where [num_int] = " & Val(t_numint)
     rs1.Open q, cn1
     While Not rs1.EOF
         F = Format$(rs1("fecha_dif"), "dd/mm/yyyy")
         FP = Format$(rs1("id_formapago"), "000") & " " & Format$(Left$(rs1("formapago"), 10), "@@@@@@@@@@!")
         nch = Format$(rs1("num_ch"), "0000000000")
         b = Format$(Left$(rs1("detalle_banco"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!")
         t = Format$(Left$(rs1("titular"), 25), ">@@@@@@@@@@@@@@@@@@@@@@@@@!")
         i = Format$(rs1("importe"), "######0.00")
         List1.AddItem FP & "  " & nch & "  " & F & "  " & b & "  " & t & "  " & i
         rs1.MoveNext
     Wend
     Set rs1 = Nothing
     
 


End Sub
Private Sub Form_Load()

Call barraesag(Me)
'List1.FontName = "Courier new"
'List1.FontSize = 9
Load gen_path
Load vta_listaprecios4

End Sub



Private Sub Form_Unload(Cancel As Integer)
Unload gen_path
Unload vta_listaprecios4
End Sub

Private Sub List1_GotFocus()

Me.StatusBar1.Panels.item(1) = "[F3] Modifica - [F4]Historial Prod. [F5] Imprime Comp. - [F6] Exporta Texto - [F7] Imprime - [F8] Borra  - [ESC] Termina "

End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    k = 0
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.FontSize = 9
    While k <= List1.ListCount - 1
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print List1.List(k)
     k = k + 1
    Wend
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
    Printer.EndDoc
  End If
End If

If KeyCode = vbKeyF4 Then 'historial producto
 Call nivel_acceso(2)
 item = Val(Mid$(List1.List(List1.ListIndex), 1, 5))
 
 If para.id_grupo_modulo_actual >= 5 Then
  If item > 1 Then
     
     vta_listaprecios4.t_idprod = item
     vta_listaprecios4.Option1 = True
     vta_listaprecios4.Show
   End If
 Else
   Call sinpermisos
 End If
End If

If KeyCode = vbKeyF5 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 6 Then
   J = MsgBox("Imprime Comprobante.", 4)
       If J = 6 Then
         Set cl_comp = New COMPROBANTES
         cl_comp.cargar2 (Val(t_numint))
         If cl_comp.numint > 0 Then
            cl_comp.imprimir
         End If
         Set cl_comp = Nothing
       End If
 Else
   Call sinpermisos
 End If
End If

If KeyCode = vbKeyF6 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 5 Then
   Select Case Val(t_tipocomp)
   Case Is = 65
     Set cl_comp = New COMPROBANTES
     cl_comp.exporta_oc (Val(t_numint))
     Set cl_comp = Nothing
   Case Else
    Call exportalist(List1, True, True, para.archivo_exportacion)
   End Select
 Else
   Call sinpermisos
 End If
End If

If KeyCode = vbKeyF8 Then
  Call nivel_acceso(2)
  If para.id_grupo_modulo_actual >= 8 Then
    cc = Val(t_numint)
    Set cl_comp = New COMPROBANTES
    cl_comp.cargar2 (cc)
    If cl_comp.numint > 0 Then
      cl_comp.borrar
    End If
    Set cl_comp = Nothing
    MsgBox ("Operacion Terminada")
  Else
    Call sinpermisos
  End If
End If

If KeyCode = vbKeyF3 Then
     cc = Val(t_numint)
     Call nivel_acceso(2)
     If para.id_grupo_modulo_actual >= 7 Then
       Set cl_comp = New COMPROBANTES
       cl_comp.cargar2 (cc)
        
       If cl_comp.numint > 0 Then
        Load cambia_estado_pago
        cambia_estado_pago.t_id = cl_comp.numint
        cambia_estado_pago.t_descripcion = cl_comp.abreviatura
        cambia_estado_pago.t_estado = cl_comp.estado_pago
        cambia_estado_pago.t_prov = ""
        cambia_estado_pago.t_idprov = cl_comp.idproveedor
        cambia_estado_pago.t_obs = cl_comp.observaciones
        cambia_estado_pago.t_fechaa = cl_comp.fecha
        cambia_estado_pago.t_fecha = cl_comp.fecha
        cambia_estado_pago.t_moneda = cl_comp.moneda
        cambia_estado_pago.t_cotizacion = cl_comp.cotizacion
        cambia_estado_pago.t_subtotal = cl_comp.subtotal
        cambia_estado_pago.t_nograv = cl_comp.nograbado
        cambia_estado_pago.t_iva = cl_comp.iva
        cambia_estado_pago.T_TOTAL = cl_comp.total
        cambia_estado_pago.T_total2 = cl_comp.total_om
        cambia_estado_pago.c_zona = cl_comp.zona
        cambia_estado_pago.c_cuenta.ListIndex = buscaindice(cambia_estado_pago.c_cuenta, cl_comp.idcuenta)
        cambia_estado_pago.Show
       End If
       Set cl_comp = Nothing
     End If
End If


End Sub
Private Sub IMPRIMERETG(ByVal n As Double)
'imprime retencion de ganancia
'n es el num-int
Dim gf1 As String
Set rs = New ADODB.Recordset
q = "select * from a5, a1, g1 where [num_int] = " & n & " and a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_usuario] = g1.[id_usuario]"
rs.Open q, cn1
If Not rs.BOF And Not rs.EOF Then
     copias = 1
     gf1 = "Courier New"
     ip = Space$(9)
     ret = Space$(9)
     For h = 1 To copias
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Comprobante Retencion Ganancia - R.G.(AFIP) 830"
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Regimen de Retencion del IMPUESTO A LAS GANANCIAS"
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(50); "Nro.: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(3); Format$(rs("SUCURSAL"), "0000") & "-" & Format$(rs("num_comprobante"), "00000000")
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(50); "Fecha:";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print Spc(3); rs("fecha")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Agente de Retencion:"; glo.nombrecli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "                    "; glo.direccioncli
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "           c.u.i.t :"; glo.CUIT
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
 
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Provedor : "; Spc(1);
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   ("; Format$(rs("a5.id_proveedor"), "00000"); ")    "; rs("denominacion")
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Direccion: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("direccion")
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "     CUIT: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("cuit")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "__________________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "N", 10)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Regimen: " & rs("a5.id_codretgan")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Pagos a la fecha   : "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Pago a Efectuar    : "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Minimo No Imponible: "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Monto Suj.a Reten. : "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Alicuota           : "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Ret. Anteriores    : "
       RSet ip = Format$(rs("total"), "####0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "RETENCION          : " & ip
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       Call definefont(gf1, "S", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Son pesos: ";
       Call definefont(gf1, "N", 12)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print convierte(Format$(rs("total"), "0.00"))
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Ingresa en DDJJ periodo: "; Mid$(rs("fecha"), 4, 2) & "/" & Mid$(rs("fecha"), 7, 2)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Orden de Pago"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "Numero                        "
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "------------------------------------------------"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print rs("obs")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print
       
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "   ________________________                     ___________________________"
       Call definefont(gf1, "N", 9)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.Print "       Por " & glo.nombrecli & "                                                   Por " & rs("denominacion")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
       Printer.NewPage
     Next h
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.EndDoc
End If
Set rs = Nothing

End Sub


