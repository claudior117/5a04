VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form op_fp1_1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INGRESO CHEQUES DE TERCERO"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   5775
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9255
      Begin VB.TextBox t_fechai 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox c_cuenta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   8
         Top             =   4680
         Width           =   4335
      End
      Begin VB.TextBox t_funcion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   8400
         MaxLength       =   8
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_origen 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   4200
         Width           =   5775
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   5
         Top             =   3240
         Width           =   5775
      End
      Begin VB.TextBox t_fechad 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox t_fechae 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox T_NUMCH 
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
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox t_importe 
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
         Left            =   1680
         MaxLength       =   21
         TabIndex        =   9
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox t_titular 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   3720
         Width           =   5775
      End
      Begin VB.TextBox t_banco 
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
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2760
         Width           =   5775
      End
      Begin VB.TextBox t_NUMINT 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Ingreso"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cuenta Entrada"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7080
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Num.Int."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Entregado por"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Sucursal"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Dif."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Emision"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Titular"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Banco"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Num.Ch."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   5970
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   53
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
            TextSave        =   "29/02/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "04:35 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "op_fp1_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
   
     
         
End Select

End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 9)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "Sin Imputacion", 0
c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_valores_terceros)


End Sub



Sub graba()
  'On Error GoTo ERRORGRABA
  If t_funcion = "A" Then
      q = "select * from cyb_01 where [id_forma_pago] = 3" 'cheque de terceros
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        If rs("CAJA") = "S" Then
           ctach = rs("id_cuenta_cont")
        Else
           ctach = 0
         End If
         cta = rs("id_cuenta_cont") 'para asiento
      Else
         ctach = 0
      End If
      Set rs = Nothing

  
      q = "select * from cyb_03"
      Set rs = New ADODB.Recordset
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      rs.AddNew
      rs("fecha_emision") = t_fechae
      rs("num_cheque") = Val(T_NUMCH)
      rs("banco") = t_banco
      rs("sucursal") = t_sucursal
      rs("titular") = t_titular
      rs("importe") = Val(t_importe)
      rs("estado") = "C"
      rs("fecha_dif") = t_fechad
      rs("origen") = t_origen
      rs("destino") = " "
      rs("num_mov_banco_i") = 0
      rs("num_mov_banco_e") = 0
      rs("num_int_op") = 0
      rs("num_int_rbo") = 0
      rs("fecha_salida") = Format$(t_fechai, "dd/mm/yyyy")
      rs("fecha_ingreso") = Format$(t_fechai, "dd/mm/yyyy")
      rs("tipo_salida") = "C"
      rs.Update
      
      numintch = rs("num_interno")
      
      
      cn1.BeginTrans
      QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
      QUERY = QUERY & " VALUES (" & ctach & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & t_origen & "', " & Val(t_importe) & ", 'D', '" & Format$(t_fechai, "DD/MM/YYYY") & "', " & numintch & ", 'H', 'Ing.Ch.Manual ', 3, " & numintch & ", " & para.id_usuario & ")"
      cn1.Execute QUERY
      
      
     'contabilidad
     If c_cuenta.ListIndex > 0 Then
         numintcgr = saca_ultnumero_int_comp("G")
         u1 = "D"
         u2 = "H"
                  
         Set rs = New ADODB.Recordset
         q = "select * from c_01 where [id_cuenta] = " & cta
         rs.Open q, cn1
         If Not rs.EOF And Not rs.BOF Then
           dcta = rs("descripcion")
         Else
           dcta = "Cuenta Inexistente"
         End If
         Set rs = Nothing
         
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fechae & "', '[Ing.CHT] " & "', 'H', " & numintch & ", " & Val(t_importe) & ", " & Val(t_importe) & ", " & para.id_usuario & ", '" & Left$(RTrim$(t_origen), 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre valores
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u1 & "', " & Val(t_importe) & ", 'Ing.Ch.Terc. Nro." & Format$(Val(T_NUMCH), "00000000") & "')"
         cn1.Execute QUERY
      
         'cuenta contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 2, " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_importe) & ", 'Ing.Ch.Terc. Nro." & Format$(Val(T_NUMCH), "00000000") & "')"
         cn1.Execute QUERY
      
      End If
      
      cn1.CommitTrans
            
     
Else
   MsgBox ("No se puede modificar  ")
End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
  
   End Sub
 
  
Sub limpia()
T_NUMCH = ""
t_fechai = ""
t_fechad = ""
t_banco = ""
t_sucursal = ""
t_titular = ""
t_importe = ""
T_NUMCH.SetFocus
End Sub


Private Sub t_banco_LostFocus()
t_banco = RTrim$(t_banco) & " "

End Sub

Private Sub t_fechad_LostFocus()
If Not IsNull(t_fechad) Then
  If Not IsDate(t_fechad) Then
     t_fechad = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fechad = Format$(Now, "dd/mm/yyyy")
End If

End Sub

Private Sub t_fechae_LostFocus()
If Not IsNull(t_fechae) Then
  If Not IsDate(t_fechae) Then
     t_fechae = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fechae = Format$(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub t_fechai_LostFocus()
If Not IsNull(t_fechai) Then
  If Not IsDate(t_fechai) Then
     t_fechai = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fechai = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fechai)
  

End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If estadocaja(t_fechai) = "A" Then
   If verificaperiodog(t_fechai) = "A" Then
    J = MsgBox("Graba Movimiento", 4)
    If J = 6 Then
     Call graba
     Call limpia
    End If
   Else
    MsgBox ("Periodo Cerrado. Imposible realizar operacion")
   End If
  Else
   MsgBox ("Caja Cerrada. Imposible realizar esta operacion")
  End If
  Me.Hide
 Else
  Call solonum(KeyAscii, 1)
 End If
End Sub

Private Sub T_NUMCH_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 0)
End Sub

Private Sub T_NUMCH_LostFocus()
T_NUMCH = Format$(Val(T_NUMCH), "0000000000")
End Sub

Private Sub t_origen_LostFocus()
t_origen = RTrim$(t_origen) & " "

End Sub

Private Sub t_sucursal_LostFocus()
t_sucursal = RTrim$(t_sucursal) & " "

End Sub

Private Sub t_titular_LostFocus()
t_titular = RTrim$(t_titular) & " "

End Sub
