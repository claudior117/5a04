VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cyb_retiroef 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "RETIRO EFECTIVO POR CAJA O CAJERO"
   ClientHeight    =   5250
   ClientLeft      =   1725
   ClientTop       =   2160
   ClientWidth     =   10650
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5250
   ScaleWidth      =   10650
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   8640
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "CYB020.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CYB020.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proveedor"
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   10095
      Begin VB.TextBox Detalle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   49
         TabIndex        =   4
         Top             =   2160
         Width           =   7335
      End
      Begin VB.TextBox t_ingresado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         MaxLength       =   14
         TabIndex        =   3
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox t_numcomp 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox c_banco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   7335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Importe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Nro. Tique"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         Caption         =   "Banco:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox t_numint 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4995
      Width           =   10650
      _ExtentX        =   18785
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
            TextSave        =   "27/02/2015"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "09:41"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Nro. Interno:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "cyb_retiroef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim EXISTE As String
Dim phoy As Double
Dim crg As Integer
Dim pagomes As Double
Dim retmes As Double
Dim impnosujret As Double
Dim rg_alicuota As Double
Dim excedente As Double
Dim gnumintop As Long








Private Sub btnacepta_Click()
J = MsgBox("Confirma Retiro Efectivo", 4)
If J = 6 Then
 If verifica Then
  If verificaperiodog(t_fecha) = "A" Then
    Call graba
    Call pi3
  Else
    MsgBox ("Periodo Cerrado. Imposible grabar Operacion")
  End If
 End If
End If
End Sub

Private Sub btnsale_Click()
inicio_bancos.Show
Unload Me
End Sub

Private Sub c_banco_GotFocus()
Call pi3
End Sub

Private Sub c_banco_LostFocus()
If c_banco.ListIndex < 0 Then
  c_banco.ListIndex = 0
End If
End Sub

Private Sub detalle_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Detalle = Detalle & " "
   If Val(t_ingresado) > 0 Then
     btnacepta.Enabled = True
     btnacepta.SetFocus
   Else
     MsgBox ("El importe debe ser mayor a cero")
   End If
End If



End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 4)
  'Case Is = 27
  '      Me.Hide
End Select
End Sub


Function verifica() As Boolean
  v = True
  If Val(t_ingresado) <= 0 Then
    MsgBox ("El importe debe ser mayor quer cero")
    v = False
  End If
  verifica = v
  
  
End Function

Sub graba()
            
     Set rs = New ADODB.Recordset
     q = "select * from cyb_01 where [id_forma_pago] = " & c_banco.ItemData(c_banco.ListIndex)
     rs.Open q, cn1
     ctabanco = rs("id_CUENTA_CONT")
     Set rs = Nothing

      cn1.BeginTrans
      
      Set rs = New ADODB.Recordset
      q = "select * from cyb_04"
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      rs.AddNew
        rs("id_banco") = c_banco.ItemData(c_banco.ListIndex)
        rs("fecha") = t_fecha
        rs("importe") = Val(t_ingresado)
        rs("id_tipomov") = 70
        rs("fecha_dif") = t_fecha
        rs("ubicacion") = "D"
        rs("entro") = "S"
        rs("fecha_acreed") = t_fecha
        rs("num_comp") = Val(t_numcomp)
        rs("detalle") = Detalle & " "
        rs("Modulo") = "B"
        rs("num_mov_int") = rs("num_mov_banco")
        rs("id_tipodbcr") = 1
        rs("num_mov_int_compras") = 0
      rs.Update
      
      numintb = rs("num_mov_banco")
      Set rs = Nothing
      
      
      
      'mueve caja efectivo
      Set rs = New ADODB.Recordset
      q = "select * from cyb_01 where [id_forma_pago] = 1"
      rs.Open q, cn1
      If Not rs.BOF And Not rs.EOF Then
        If rs("caja") = "S" Then
          'grabo mov caja
           QUERY = "INSERT INTO cyb_05([id_cuenta_caja], [id_cuenta_contra], [descripcion], [importe], [ubicacion], [fecha], [num_mov_int], [modulo], [operacion], [id_forma_pago], [num_int_ch_terc], [id_usuario])"
           QUERY = QUERY & " VALUES (" & rs("id_cuenta_cont") & ", " & ctabanco & ", '" & Left$("Retiro Efectivo Nro." & t_numcomp, 49) & " ', " & Val(t_ingresado) & ", 'D', '" & t_fecha & "', " & numintb & ", 'B', 'Ret. " & Format$(numintb, "00000000") & "', 1, 0, " & para.id_usuario & ")"
           cn1.Execute QUERY
        End If
      End If
      Set rs = Nothing
      
      
'grabo cntabilidad
If Generaasientosauto Then
      Set rs = New ADODB.Recordset
      q = "select * from cyb_06 where [id_tipomov] = 70"
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
        contab = rs("contabilidad")
      Else
        contab = "N"
      End If
      Set rs = Nothing
      
      If contab <> "N" Then
         numintcgr = saca_ultnumero_int_comp("G")
         u1 = contab
         If u1 = "D" Then
           u2 = "H"
         Else
           u2 = "D"
         End If
         
          
         'grabo asiento
         QUERY = "INSERT INTO c_02([num_interno], [fecha], [descripcion], [modulo], [num_mov_int], [debe], [haber], [id_USUARIO], [observaciones])"
         QUERY = QUERY & " VALUES (" & numintcgr & " ,'" & t_fecha & "', '[Retiro Ef.] N.I." & Format$(numintb, "00000000") & "', 'B', " & numintb & ", " & Val(t_ingresado) & ", " & Val(t_ingresado) & ", " & para.id_usuario & ", '" & Left$(Detalle & " Ret.Nro." & t_numcomp, 50) & "')"
         cn1.Execute QUERY
      
         'cuenta madre bancos
         QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
         QUERY = QUERY & " VALUES (" & numintcgr & ", 1, " & ctabanco & ", '" & u1 & "', " & Val(t_ingresado) & ", 'Ret.N.I" & Format$(numintb, "00000000") & "')"
         cn1.Execute QUERY
      
         'formas de pago
         ic = 2
         Set rs = New ADODB.Recordset
         q = "select * from cyb_01 where [id_forma_pago] = 1"
         rs.Open q, cn1
         If Not rs.BOF And Not rs.EOF Then
              d = "Retiro ef. banco " & Format$(numintb, "00000000")
              cta = rs("id_cuenta_cont")
              QUERY = "INSERT INTO c_03([num_interno], [renglon], [id_cuenta], [ubicacion], [importe], [descripcion])"
              QUERY = QUERY & " VALUES (" & numintcgr & ", " & ic & ", " & cta & ", '" & u2 & "', " & Val(t_ingresado) & ", '" & d & "')"
              cn1.Execute QUERY
          End If
        
       End If
 End If
     
     
cn1.CommitTrans
 

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos. Proc:Retiro ef.  Funcion:Graba")
  


End Sub

Private Sub Form_Load()
c_banco.clear
Call carga_formas_pago(c_banco, "B")


End Sub





Private Sub pi3()
   Call INICIALIZA2(Me)
   btnacepta.Enabled = False
   c_banco.SetFocus
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


Private Sub t_ingresado_KeyPress(KeyAscii As Integer)
Call solonum(KeyAscii, 1)

End Sub
