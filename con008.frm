VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cambia_estado_pago 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODIFICA COMPROBANTE DE COMPRA"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7425
   ScaleWidth      =   8535
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mofifica Totales"
      Height          =   1335
      Left            =   120
      TabIndex        =   38
      Top             =   4680
      Width           =   7455
      Begin VB.TextBox t_total2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   5880
         MaxLength       =   8
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_total 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_iva 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3000
         MaxLength       =   8
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_nograv 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   8
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox t_cotizacion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   3480
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "mayor o igual a  1"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox t_moneda 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   6
         ToolTipText     =   "[P] pesos  -  [D] Dolares"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_subtotal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   120
         MaxLength       =   8
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Modifica Totales"
         Height          =   255
         Left            =   4680
         TabIndex        =   39
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Total 2"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   46
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Total"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Iva"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3000
         TabIndex        =   44
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "No Gravado"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Cotizacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Moneda:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Subtotal"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modifica Datos Generales"
      Height          =   1455
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   8295
      Begin VB.CommandButton Command4 
         Caption         =   "Busca Cuenta"
         Height          =   255
         Left            =   6960
         TabIndex        =   51
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   960
         Width           =   5055
      End
      Begin VB.ComboBox c_zona 
         Height          =   315
         ItemData        =   "con008.frx":0000
         Left            =   4440
         List            =   "con008.frx":000D
         TabIndex        =   49
         Text            =   "Combo1"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modifca Datos Generales"
         Height          =   255
         Left            =   6120
         TabIndex        =   37
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox t_fecha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox t_obs 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1800
         MaxLength       =   70
         TabIndex        =   1
         ToolTipText     =   "Ingrese los digitos del 4 al 7 del cod. de barra de algun articulo de la marca"
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Imputacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modifica estado de pago"
      Height          =   1095
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Cambiar Estado pago"
         Height          =   255
         Left            =   2760
         TabIndex        =   33
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox t_numcomp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   5
         ToolTipText     =   "Ingrese los digitos del 4 al 7 del cod. de barra de algun articulo de la marca"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox t_sucursal 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   4
         ToolTipText     =   "Ingrese los digitos del 4 al 7 del cod. de barra de algun articulo de la marca"
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox T_newestado 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   3
         ToolTipText     =   "[N] Sin Pagar  -  [P] Pagado"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Num. Orden Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         Caption         =   "Estado Pago"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estado Actual del Comprobante"
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   8295
      Begin VB.TextBox t_fechaa 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox T_IDPROV 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox t_prov 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox t_estado 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "[N] Sin Pagar  -  [P] Pagado "
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox t_id 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox t_descripcion 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Actual"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "[N] Sin Pagar     -    [P] Pagado"
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Proveedor"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Estado Pago Actual"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Num. Interno"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Comprobante"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   6840
      TabIndex        =   16
      Top             =   6120
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "con008.frx":0032
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
         Picture         =   "con008.frx":08B4
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
      Top             =   7170
      Width           =   8535
      _ExtentX        =   15055
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
            TextSave        =   "11/10/2012"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:16"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label18 
      BackColor       =   &H0080FFFF&
      Caption         =   $"con008.frx":1136
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   47
      Top             =   6480
      Width           =   6255
   End
End
Attribute VB_Name = "cambia_estado_pago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String



Sub graba()
J = MsgBox("Confirma Modifcar Estado de Pago", 4)
If J = 6 Then
    If T_newestado = "P" Then
      'verifico existencia de o.p.
      Set rs = New ADODB.Recordset
      q = "select * from a5 where [sucursal] = " & Val(t_sucursal) & " and [num_comprobante] = " & Val(t_numcomp) & " and [letra] = 'O' and [id_tipocomp] = 50  and [id_proveedor] = " & Val(T_IDPROV)
      rs.Open q, cn1
      If Not rs.BOF And Not rs.EOF Then
             niop = rs("num_int")
             Call cambia(niop)
      Else
         Y = MsgBox("La O.P. No existe o No pertenece al proveedor. Desea continuar con el cambio de estado", 4)
         If Y = 6 Then
           niop = 0
           Call cambia(niop)
         End If
      End If
      Set rs = Nothing
    Else
      T_newestado = "N"
      t_sucursal = "0000"
      t_numcomp = "00000000"
      Call cambia(0)
    End If
End If

Exit Sub
ERRORGRABA:
MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
End Sub
Sub graba2()
J = MsgBox("Confirma Modificar Datos Genereales del Comprobante", 4)
If J = 6 Then
    m = 0
    If t_fecha <> t_fechaa Then
       If verificaperiodo(t_fecha) = "A" Then
          If verificaperiodo(t_fechaa) = "A" Then
              'modifico datos
              m = 1
          Else
            MsgBox ("El periodo de origen esta Cerrado")
          End If
       Else
          MsgBox ("El periodo Destino esta Cerrado")
       End If
    End If
      
    Set rs = New ADODB.Recordset
    q = "select [fecha], [obs], [zona], [id_cuenta] from a5 where [num_int] = " & Val(t_id)
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs.BOF And Not rs.EOF Then
       rs("obs") = t_obs & " "
       If m = 1 Then
         rs("fecha") = t_fecha
       End If
       If c_zona.ListIndex > 0 Then
         rs("zona") = c_zona.ListIndex
       End If
       
       If c_cuenta.ListIndex >= 0 Then
         rs("id_cuenta") = c_cuenta.ItemData(c_cuenta.ListIndex)
       End If
       rs.Update
    End If
    Set rs = Nothing
         
    If m = 1 Then
      Call grabausuario
    End If
    MsgBox ("Tarea Finalizada")
   
    
End If

End Sub
Sub cambia(ByVal nio)
'nio = num int op
Set rs1 = New ADODB.Recordset
q = "select * from a5 where [num_int] = " & Val(t_id)
rs1.Open q, cn1, adOpenDynamic, adLockOptimistic, 1
If Not rs1.BOF And Not rs1.EOF Then
           rs1("estado_pago") = T_newestado
           rs1("num_op") = Format$(Val(t_sucursal), "0000") & "-" & Format$(Val(t_numcomp), "00000000")
          
           If T_newestado = "N" Then
             'pone el comprobante como no pago
             If rs1("moneda") = "P" Then
                t = rs1("total")
             Else
                t = rs1("total_d")
             End If
             rs1("saldo_impago") = t
             
             'limpio todas la op del comp a15
             Set rs2 = New ADODB.Recordset
             q = "select * from a15 where [num_int_comp] = " & rs1("num_int")
             rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
             While Not rs2.EOF
               rs2.Delete
               rs2.MoveNext
             Wend
             Set rs2 = Nothing
           Else
                          
              Set rs2 = New ADODB.Recordset
              q = "select * from a15 where  [num_int_comp] = " & rs1("num_int") & " and [num_int_op] = " & nio
              rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
              If Not rs2.EOF And Not rs2.BOF Then
                ' rs2.ew
                ipa = rs2("importe_pagado") + rs1("saldo_impago")
              Else
                 ipa = rs1("saldo_impago")
                 rs2.AddNew
              End If
                rs2("num_int_comp") = rs1("num_int")
                rs2("num_int_op") = nio
                rs2("importe_pagado") = ipa
                rs2("saldo_comprobante") = ipa
                rs2.Update
              Set rs2 = Nothing
           
              rs1("saldo_impago") = 0
           End If
          
          rs1.Update
          
          Call grabausuario
          
  End If
 Set rs1 = Nothing
 MsgBox ("Operacion Terminada")
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub Command1_Click()
Call graba
End Sub

Private Sub Command2_Click()
Call graba2

End Sub

Private Sub Command3_Click()
Call graba3

End Sub

Private Sub Command4_Click()
cgr_buscacuenta.Show
End Sub

Private Sub Form_Activate()
If t_moneda = "P" Then
  Label16 = "Total $"
  Label17 = "Total U$$"
Else
  Label17 = "Total $"
  Label16 = "Total U$$"

End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   Case Is = vbKeyF9
     Call graba
         
End Select

End Sub
Sub graba3()
J = MsgBox("Confirma Modificar TOTALES del Comprobante", 4)
If J = 6 Then
    m = 0
   If verificaperiodo(t_fechaa) = "A" Then
      'modifico datos
       m = 1
   Else
      MsgBox ("El periodo del comprobante esta Cerrado")
   End If
   If m = 1 Then
   
    Set rs = New ADODB.Recordset
    q = "select [moneda], [cotiz_dolar], [subtotal], [iva], [no_grabado], [total], [total_d]  from a5 where [num_int] = " & Val(t_id)
    rs.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs.BOF And Not rs.EOF Then
       rs("moneda") = t_moneda
       rs("cotiz_dolar") = Val(t_cotizacion)
       rs("iva") = Val(t_iva)
       rs("no_grabado") = Val(t_nograv)
       rs("total") = Val(t_total)
       rs("total_d") = Val(t_total2)
       rs.Update
    End If
    Set rs = Nothing
    
    
     Call grabausuario
    
    MsgBox ("Tarea Finalizada")
   
   End If
End If

End Sub
Sub grabausuario()
     QUERY = "INSERT INTO g11([detalle], [id_usuario], [modulo], [num_int_comp], [fecha_hora], [obs], [id_operacion], [id_clipro])"
     QUERY = QUERY & " VALUES ('Modif. Comp. Compra NI " & Val(t_id) & "' , " & para.id_usuario & ", 'C'," & Val(t_id) & ", '" & Now & "', '" & t_descripcion & "', 100, " & Val(T_IDPROV) & ")"
     cn1.BeginTrans
     cn1.Execute QUERY
     cn1.CommitTrans
     
     
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 12)
  Case Is = 27
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
c_zona.ListIndex = 0
Call carga_cuentas_cont(c_cuenta, "C", "D")
End Sub


Private Sub t_cotizacion_LostFocus()
If Val(t_cotizacion) < 1 Then
  t_cotizacion = 1
End If
End Sub

Private Sub t_descripcion_LostFocus()
If t_descripcion = "" Then
  t_descripcion = "Null"
End If
End Sub



Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(t_fecha, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)

End Sub

Private Sub t_moneda_LostFocus()
t_moneda = Format$(t_moneda, ">@")
Select Case t_moneda
 Case Is = "P", Is = "D"
 Case Else
   t_moneda = "P"
End Select
If t_moneda = "P" Then
  Label16 = "Total $"
  Label17 = "Total U$$"
Else
  Label17 = "Total $"
  Label16 = "Total U$$"
End If

End Sub

Private Sub T_newestado_LostFocus()
'FIXIT: Reemplazar la función 'UCase' con la función 'UCase$'.                             FixIT90210ae-R9757-R1B8ZE
T_newestado = UCase(T_newestado)
If T_newestado <> "N" And T_newestado <> "P" Then
   T_newestado = "N"
End If

End Sub

Private Sub t_numcomp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  btnacepta.SetFocus
End If
End Sub

Private Sub t_numcomp_LostFocus()
t_numcomp = Format$(Val(t_numcomp), "00000000")

End Sub

Private Sub t_sucursal_LostFocus()
t_sucursal = Format$(Val(t_sucursal), "0000")

End Sub
