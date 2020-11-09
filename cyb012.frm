VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_salidachterc 
   BackColor       =   &H00C0C0C0&
   Caption         =   "SALIDA MANUAL DE VALORES DE TERCEROS"
   ClientHeight    =   8490
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   7005
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8490
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5160
      TabIndex        =   31
      Top             =   7440
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cyb012.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cyb012.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Salida"
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   6615
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox t_destino 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   1
         Top             =   720
         Width           =   4935
      End
      Begin VB.ComboBox c_tipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "cyb012.frx":1104
         Left            =   1560
         List            =   "cyb012.frx":1117
         TabIndex        =   0
         Top             =   240
         Width           =   3615
      End
      Begin VB.ComboBox c_cuenta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Fecha Salida"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Entregado a:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Tipo Salida"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Cuenta Salida"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Entrada"
      Height          =   4815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   6615
      Begin VB.TextBox t_funcion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   6000
         MaxLength       =   8
         TabIndex        =   24
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox t_origen 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3720
         Width           =   4335
      End
      Begin VB.TextBox t_sucursal 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   8
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox t_fechad 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox t_fechae 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox T_NUMCH 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   11
         Top             =   4200
         Width           =   975
      End
      Begin VB.TextBox t_titular 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   9
         Top             =   3240
         Width           =   4335
      End
      Begin VB.TextBox t_banco 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2280
         Width           =   4335
      End
      Begin VB.TextBox t_NUMINT 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   405
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Funcion"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4680
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Num.Int."
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Entregado por"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Sucursal"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Fecha Dif."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Fecha Emision"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Importe"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Titular"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Banco"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Num.Ch."
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   8460
      Width           =   7005
      _ExtentX        =   12356
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
            TextSave        =   "20/03/2014"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "18:44"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "cyb_salidachterc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984



Private Sub btnsale_Click()
Me.Hide
End Sub

Private Sub c_cuenta_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  If c_cuenta.ListIndex < 0 Then
  c_cuenta.ListIndex = 0
  End If
  If estadocaja(t_fecha) = "A" Then
   If verificaperiodog(t_fecha) = "A" Then
   J = MsgBox("Graba Movimiento", 4)
   If J = 6 Then
    Call graba
    Call limpia
   End If
  Else
   MsgBox ("Periodo cerrado. Imposible grabar operacion")
  End If
  Else
   MsgBox ("Caja cerrada. Imposible realizar esta operacion")
  End If
  Me.Hide
 End If

End Sub

Private Sub c_cuenta_LostFocus()
If c_cuenta.ListIndex < 0 Then
  If Val(c_cuenta) > 0 Then
    c_cuenta.ListIndex = buscaindice(c_cuenta, Val(c_cuenta))
  Else
    c_cuenta.ListIndex = 0
  End If
End If

End Sub

Private Sub c_tipo_LostFocus()
If c_tipo.ListIndex < 0 Then
  c_tipo.ListIndex = 0
End If
If c_tipo.ListIndex = 2 Then
  c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_caja)
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
    Call TabEnter2(Me, 3)
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "Sin Imputacion", 0
c_cuenta.ListIndex = 0
End Sub



Sub graba()
      'On Error GoTo ERRORGRABA
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

  
            
      QUERY = "update cyb_03 set  [estado]='" & Mid$(c_tipo, 2, 1) & "' , [destino]='" & t_destino & "' , [tipo_salida]='M' , [fecha_salida]='" & t_fecha & "'"
      QUERY = QUERY & " where [num_interno]= " & Val(t_numint)
      cn1.BeginTrans
      cn1.Execute QUERY
      
      If t_funcion = "M" Then
       'borro los mov. de caja
       QUERY = "DELETE FROM cyb_05 WHERE [num_mov_int] = " & Val(t_numint) & " and [modulo] = 'I'"
        cn1.Execute QUERY

       Set rs2 = New ADODB.Recordset
       q = "select * from c_02 WHERE [num_mov_int] = " & Val(t_numint) & " and [modulo] = 'I'"
       rs2.Open q, cn1
       If Not rs2.EOF And Not rs2.BOF Then
              QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & rs2("num_interno")
              cn1.Execute QUERY
       End If
       Set rs2 = Nothing
       
       QUERY = "DELETE FROM c_02 WHERE [num_mov_int] = " & Val(t_numint) & " and [modulo] = 'I'"
       cn1.Execute QUERY
       
       
       
       
      End If
       
       
       'salida de caja del ch.
       
       If Mid$(c_tipo, 2, 1) <> "C" Then
        QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
        QUERY = QUERY & " VALUES (" & ctach & ", " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & t_destino & "', " & Val(t_importe) & ", 'H', '" & Format$(t_fecha, "DD/MM/YYYY") & "', " & Val(t_numint) & ", 'I', 'Sal.ChT.Man.', 3, " & Val(t_numint) & ", " & para.id_usuario & ")"
        cn1.Execute QUERY
       End If
      
      'entrada de ef. si es cobrado
      If Mid$(c_tipo, 2, 1) = "J" Then
         QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
         QUERY = QUERY & " VALUES (" & para.cuenta_caja & ", " & cta & ", '" & t_destino & "', " & Val(t_importe) & ", 'D', '" & Format$(t_fecha, "DD/MM/YYYY") & "', " & Val(t_numint) & ", 'I', 'Cobranza ChT.',1, " & Val(t_numint) & ", " & para.id_usuario & ")"
         cn1.Execute QUERY
         
         c_cuenta.ListIndex = buscaindice(c_cuenta, para.cuenta_caja)
         
      End If
        
      
      
     'contabilidad
     
    If Generaasientosauto Then
     If c_cuenta.ListIndex > 0 Then
         numintcgr = saca_ultnumero_int_comp("G")
         u1 = "H"
         u2 = "D"
                  
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
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Sal.CHT.Man.] " & "', 'I', " & Val(t_numint) & ", " & Val(t_importe) & ", " & Val(t_importe) & ", " & para.id_usuario & ", '" & Left$(RTrim$(t_destino), 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre valores
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & cta & ", '" & u1 & "', " & Val(t_importe) & ", 'Sal.ChT.Man. Nro." & Format$(Val(T_NUMCH), "00000000") & "')"
         cn1.Execute QUERY
      
         'cuenta contrapartida
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 2, " & c_cuenta.ItemData(c_cuenta.ListIndex) & ", '" & u2 & "', " & Val(t_importe) & ", 'Sal.ChT.Man. Nro." & Format$(Val(T_NUMCH), "00000000") & "')"
         cn1.Execute QUERY
      
      End If
    End If
    
    cn1.CommitTrans
            
     
'Else
'   MsgBox ("No se puede modificar  ")
'End If

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos")
  
  
   End Sub
 
  
Sub limpia()
Me.Hide
End Sub


Private Sub t_banco_LostFocus()
t_banco = RTrim$(t_banco) & " "

End Sub

Private Sub t_destino_LostFocus()
t_destino = RTrim$(t_destino) & " "
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
Else
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
Call verifica_fechacorte(t_fecha)
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
Call verifica_fechacorte(t_fecha)
End Sub

Private Sub t_importe_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
  J = MsgBox("Graba Movimiento", 4)
  If J = 6 Then
    Call graba
    Call limpia
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
