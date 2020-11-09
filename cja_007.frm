VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form cja_transf_caja 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "TRASNFERENCIAS INTERNA ENTRE CAJAS"
   ClientHeight    =   5490
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
   ScaleHeight     =   5490
   ScaleWidth      =   10650
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   8640
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cja_007.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Renueva Lista de Clientes"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cja_007.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   3735
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   10455
      Begin VB.ComboBox c_concepto 
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
         Top             =   360
         Width           =   3615
      End
      Begin VB.ComboBox c_caja 
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
         TabIndex        =   1
         Top             =   840
         Width           =   7335
      End
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
         TabIndex        =   6
         Top             =   3240
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
         TabIndex        =   5
         Top             =   2760
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
         TabIndex        =   4
         Top             =   2280
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
         TabIndex        =   3
         Top             =   1800
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
         TabIndex        =   2
         Top             =   1320
         Width           =   7335
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Concepto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Entrada"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9360
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "Salida"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9360
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Caja Salida:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Importe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Nro. Comprobante:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   "Caja Entrada:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   1320
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
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   5235
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
            TextSave        =   "09:39"
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
      BackColor       =   &H00FF0000&
      Caption         =   "Nro. Interno:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "cja_transf_caja"
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
J = MsgBox("Confirma Transferencia Interna de Caja a Caja", 4)
If J = 6 Then
 If verifica Then
  If verificaperiodog(t_fecha) = "A" Then
   If estadocaja(t_fecha) = "A" Then
      
       Call graba
       Call pi3
    
   Else
     MsgBox ("Caja Cerrada. Imposible realizar operacion")
   End If
  Else
    MsgBox ("Periodo Cerrado. Imposible realizar Operacion")
  End If
 End If
End If
End Sub

Private Sub btnsale_Click()
inicio_caja.Show
Unload Me
End Sub

Private Sub c_banco_GotFocus()
Call pi3
End Sub

Private Sub c_banco_LostFocus()
If c_banco.ListIndex < 0 Then
  If Val(c_banco) > 0 Then
    c_banco.ListIndex = buscaindice(c_banco, Val(c_banco))
  Else
    c_banco.ListIndex = 0
  End If
End If
End Sub

Private Sub c_caja_LostFocus()
If c_caja.ListIndex < 0 Then
  c_caja.ListIndex = 0
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
    Call TabEnter2(Me, 6)
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
  
  If c_caja.ItemData(c_caja.ListIndex) = c_banco.ItemData(c_banco.ListIndex) Then
    MsgBox ("La caja de entrada y salida no pueden ser las mismas")
    v = False
  End If
  
  verifica = v
  
  
End Function

Sub graba()
            
     Set rs = New ADODB.Recordset
     q = "select * from cyb_01 where [id_forma_pago] = " & c_caja.ItemData(c_caja.ListIndex)
     rs.Open q, cn1
     ctacaja = rs("id_CUENTA_CONT")
     Set rs = Nothing

      'salida
      Set rs = New ADODB.Recordset
      q = "select * from cyb_05"
      rs.Open q, cn1, adOpenDynamic, adLockOptimistic
      rs.AddNew
        rs("id_cuenta_caja") = ctacaja
        rs("id_cuenta_contra") = ctacaja
        rs("Descripcion") = Left$(Detalle, 49)
        rs("fecha") = t_fecha
        rs("importe") = Val(t_ingresado)
        rs("ubicacion") = "H"
        rs("Modulo") = "J"
        rs("num_mov_int") = rs("num_mov_caja")
        rs("operacion") = "TICC(S)" & Format$(rs("num_mov_caja"), "0000000")
        rs("id_forma_pago") = c_concepto.ItemData(c_concepto.ListIndex)
        rs("num_int_ch_terc") = 0
        rs("id_usuario") = c_caja.ItemData(c_caja.ListIndex)
        
      rs.Update
     
      rs.AddNew
        rs("id_cuenta_caja") = ctacaja
        rs("id_cuenta_contra") = ctacaja
        rs("Descripcion") = Left$(Detalle, 49)
        rs("fecha") = t_fecha
        rs("importe") = Val(t_ingresado)
        rs("ubicacion") = "D"
        rs("Modulo") = "J"
        rs("num_mov_int") = rs("num_mov_caja")
        rs("operacion") = "TICC(E)" & Format$(rs("num_mov_caja"), "0000000")
        rs("id_forma_pago") = c_concepto.ItemData(c_concepto.ListIndex)
        rs("num_int_ch_terc") = 0
        rs("id_usuario") = c_banco.ItemData(c_banco.ListIndex)
        
      rs.Update
      Set rs = Nothing
      
      
      
      
        

Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  MsgBox ("Error de Actualizacion. Verifique los datos o sus permisos. Proc:Transf. Caja a Banco.  Funcion:Graba")
  


End Sub

Private Sub Form_Load()
c_banco.clear
Call carga_usuarios(c_banco)
Call carga_usuarios(c_caja)
Call carga_formas_pago(c_concepto, "C")


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
