VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form calcula_ret 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALCULO DE RETENCION"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   10905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Resultado"
      Height          =   1095
      Left            =   120
      TabIndex        =   28
      Top             =   2400
      Width           =   3975
      Begin VB.TextBox t_ret 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "Retencion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Complementarios"
      Height          =   3375
      Left            =   7080
      TabIndex        =   11
      Top             =   120
      Width           =   3495
      Begin VB.TextBox t_minimo 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox t_usatabla 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   26
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox t_inscgan 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   1
         TabIndex        =   24
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox t_impnosujret 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox t_alicuota 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox t_retmes 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_pagomes 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   13
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Minimo a Ret.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Usa Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Insc. Gan.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Imp. no suj. ret:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Alicuota:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Ret. del mes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Pagos del mes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Fecha Op:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   6735
      Begin VB.TextBox t_importe 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1680
         Width           =   1575
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox c_concepto 
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.ComboBox c_impuesto 
         Height          =   315
         Left            =   2160
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label14 
         Caption         =   "Importe Neto deducidos impuestos y descuentos"
         Height          =   375
         Left            =   3960
         TabIndex        =   33
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Importe sujeto a retencion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Impuesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         Caption         =   "Concepto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   5280
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Height          =   615
         Left            =   840
         Picture         =   "Arch011.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "Arch011.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   4
      Top             =   3690
      Width           =   10905
      _ExtentX        =   19235
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
Attribute VB_Name = "calcula_ret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String

Sub limpia()
t_ret = ""
t_pagosmes = ""
t_retmes = ""
t_alicuota = ""
t_impnosujret = ""
t_minimo = ""
t_inscgan = ""
t_usatabla = ""

End Sub

Private Sub btnacepta_Click()
If Val(t_importe) > 0 Then
 espere.Show
 espere.Label1 = "Realizando Consulta...."
 espere.Refresh
 Call limpia
 Select Case c_impuesto.ItemData(c_impuesto.ListIndex)
   Case Is = 217
   Call sacadatos
   Call calculag
     Case Is = 50
   Call calculaib
   End Select
 Unload espere
Else
 MsgBox ("El importe sujeto a retencion debe ser mayor a cero")
End If
End Sub


Private Sub btnsale_Click()
Me.Hide
End Sub



Private Sub c_concepto_LostFocus()
If c_concepto.ListIndex < 0 Then
 c_concepto.ListIndex = 0
End If
'Select Case c_impuesto.ItemData(c_impuesto.ListIndex)
' Case Is = 217
'  Call sacadatos
' Case Is = 50
'  Call sacadatosib
'End Select
End Sub

Private Sub c_impuesto_LostFocus()
If c_impuesto.ListIndex < 0 Then
  c_impuesto.ListIndex = 0
End If
Call carga_impuestos(c_concepto, c_impuesto.ItemData(c_impuesto.ListIndex))
c_concepto.ListIndex = 0

End Sub

Private Sub c_prov_LostFocus()
If c_prov.ListIndex < 0 Then
  c_prov.ListIndex = 0
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
       
End Select

End Sub
Sub calculaib()
Set cl_prov2 = New proveedores
cl_prov2.carga (c_prov.ItemData(c_prov.ListIndex))
If cl_prov2.idprov > 0 Then
   Set cl_padronib = New padron_ib
   cl_padronib.cuit_texto = cl_prov2.CUIT
   cl_padronib.buscar
   t_alicuota = Format$(cl_padronib.tasa_retib, "#0.00")
   Set cl_padronib = Nothing
   
   q = "select * from i_01 where [id_impuesto] = 50 "
   Set rs2 = New ADODB.Recordset
   rs2.Open q, cn1
   impmin = rs2("importe_minimo_sujeto_ret")
   retmin = rs2("retencion-minima")
   Set rs2 = Nothing
   
   t_impnosujret = impmin
   t_minimo = retmin
   If Val(t_importe) >= impmin Then
     t_ret = Format$(Val(t_importe) * Val(t_alicuota) / 100, "#####0.00")
     If Val(t_ret) < retmin Then
       t_ret = "0.00"
     Else
       t_ret = Format$(Val(t_importe) * Val(t_alicuota) / 100, "#####0.00")
     End If
   Else
     t_ret = "0.00"
   End If
Else
  t_ret = "0.00"
  MsgBox ("Error al cargar el Proveedor")
End If
Set cl_prov2 = Nothing
End Sub
Sub calculag()
total = Val(t_pagomes) + Val(t_importe)
 excedente = Val(t_pagomes) + Val(t_importe) - Val(t_impnosujret)
  t_ret = 0
  Set rs = New ADODB.Recordset
  q = "select * from i_01 where [id_impuesto] = 217"
  rs.Open q, cn1
  t_minimo = rs("retencion-minima")
  impmin = rs("importe_minimo_sujeto_ret")
  Set rs = Nothing
  
If total >= impmin Then
  If t_usatabla = "N" Then
    t_ret = (excedente * Val(t_alicuota) / 100) - Val(t_retmes)
  Else
   'por tabla
    Set rs = New ADODB.Recordset
    q = "select * from i_03 where [id_impuesto] = 217 order by [secuencia]"
    rs.Open q, cn1, adOpenStatic, adLockReadOnly
    While Not rs.EOF
       If excedente >= rs("minimo") And excedente <= rs("maximo") Then
           'ENCONTRE LA RETENCION
           R1 = rs("importe_retenido")
           R2 = (excedente - rs("sobre_EXcEDENTE")) * (rs("porcentaje_extra") / 100)
           t_ret = R1 + R2 - retmes
           rs.MoveLast
       End If
       rs.MoveNext
   Wend
   Set rs = Nothing
  End If
  If Val(t_ret) >= Val(t_minimo) Then
    t_ret = Format$(t_ret, "#####0.00")
  Else
    t_ret = "0.00"
  End If
Else
 t_ret = "0.00"
End If
    
End Sub
Sub sacadatos()
'saco pago mes y ret mes
m = Val(Mid$(t_fecha, 4, 2))
a = Val(Mid$(t_fecha, 7, 4))
t_retmes = 0
t_pagomes = 0
t_alicuota = 0
t_impnosujret = 0

Set cl_prov2 = New proveedores
cl_prov2.carga (c_prov.ItemData(c_prov.ListIndex))
t_inscgan = cl_prov2.inscriptogan
If cl_prov2.idprov > 0 Then

  Set rs = New ADODB.Recordset
  q = "select * from ret_01 where [id_proveedor] = " & cl_prov2.idprov & " and [id_retgan] = " & c_concepto.ItemData(c_concepto.ListIndex) & " and [mes] = " & m & " and [año] = " & a
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
    t_retmes = rs("ret_mes")
    t_pagomes = rs("pagos_mes")
  Else
    t_retmes = 0
    t_pagomes = 0
  End If
  Set rs = Nothing


  q = "select * from i_02 where [id_impuesto] = 217 and [id_concepto] = " & c_concepto.ItemData(c_concepto.ListIndex)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
    If cl_prov2.inscriptogan = "S" Then
      t_impnosujret = rs("importe_noretenido")
      t_usatabla = rs("porescala_i")
      t_alicuota = rs("alicuota_i")
    Else
      t_impnosujret = rs("importe_noretenido_n")
      t_usatabla = rs("porescala_n")
      t_alicuota = rs("alicuota_n")
    End If
  End If
  Set rs = Nothing
End If
Set cl_prov2 = Nothing
End Sub

Sub sacadatosib()
  q = "select * from i_02 where [id_impuesto] = 50 and [id_concepto] = " & c_concepto.ItemData(c_concepto.ListIndex)
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  If Not rs.BOF And Not rs.EOF Then
      t_impnosujret = rs("importe_noretenido")
      t_alicuota = rs("alicuota_i")
  End If
  Set rs = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 27
        Me.Hide
End Select
End Sub

Private Sub Form_Load()
Call barraesag(Me)
Call carga_proveedores(c_prov)
c_prov.ListIndex = 0
Call carga_impuesto(c_impuesto)
 c_impuesto.ListIndex = 0
 
Call carga_impuestos(c_concepto, c_impuesto.ItemData(c_impuesto.ListIndex))
c_concepto.ListIndex = 0
t_fecha = Format$(Now, "dd/mm/yyyy")

End Sub



Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
   If Not IsDate(t_fecha) Then
     t_fecha = Format$(Now, "dd/mm/yyyy")
   End If
Else
   t_fecha = Format$(Now, "dd/mm/yyyy")
End If
End Sub

Private Sub t_importe_LostFocus()
Select Case c_impuesto.ItemData(c_impuesto.ListIndex)
 Case Is = 217
  Call calculag
 Case Is = 50
  Call calculaib
End Select

End Sub

