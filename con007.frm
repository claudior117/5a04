VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form con_saldosprov 
   BackColor       =   &H00E0E0E0&
   Caption         =   "SALDOS PROVEEDORES"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17760
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9435
   ScaleWidth      =   17760
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   29
      Top             =   2040
      Width           =   3975
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra Saldo a favor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   360
      TabIndex        =   26
      Top             =   8040
      Width           =   6495
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha vencimiento"
         Height          =   495
         Left            =   3480
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha Comprobante"
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   360
      TabIndex        =   21
      Top             =   1320
      Width           =   11055
      Begin VB.ComboBox c_zona 
         Height          =   315
         Left            =   1920
         TabIndex        =   25
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox t_cliente 
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
         TabIndex        =   23
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Zona:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Proveedor(texto):"
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
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Orden"
      Height          =   735
      Left            =   6600
      TabIndex        =   17
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton Option5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Razon Social"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Id. Prov"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13200
      TabIndex        =   14
      Top             =   1440
      Width           =   3975
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Muestra Saldo en U$s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3615
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   38666241
      CurrentDate     =   38803
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestra Saldos en Cero"
      Height          =   615
      Left            =   13200
      TabIndex        =   10
      Top             =   840
      Width           =   3975
      Begin VB.OptionButton O_cero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Tag             =   "P"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton O_nocero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Tag             =   "D"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha Desde - Hasta"
      Height          =   1215
      Left            =   360
      TabIndex        =   9
      Top             =   120
      Width           =   4575
      Begin VB.TextBox t_fecha2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox t_fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha hasta(*):"
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
         TabIndex        =   33
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha desde(*):"
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
         Height          =   495
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moneda"
      Height          =   735
      Left            =   13200
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton O_dolares 
         BackColor       =   &H00E0E0E0&
         Caption         =   "U$s"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Tag             =   "D"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton O_pesos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Tag             =   "P"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   15600
      TabIndex        =   3
      Top             =   7920
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "con007.frx":0000
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
         Picture         =   "con007.frx":0882
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
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Top             =   9030
      Width           =   17760
      _ExtentX        =   31327
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   19403
            MinWidth        =   19403
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "28/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:57 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   8916
      _Version        =   393216
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   16560
      TabIndex        =   20
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "con_saldosprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Private EXISTE As String
Private saldoant As Double
Private saldoact As Double
'FIXIT: Declare 'saf' and 'df' and 'hf' and 'sf' con un tipo de datos de enlace en tiempo de compilación     FixIT90210ae-R1672-R1B8ZE
Dim saf, df, hf, sf, sof As Double




Private Sub btnacepta_Click()
If verifica Then
 If Option4 = True Then
   Call carga
 Else
   Call carga2
 End If
End If
End Sub
Function verifica() As Boolean
  verifica = True
  If t_fecha <> "" Then
    If Not IsDate(t_fecha) Then
      verifica = False
    End If
  Else
    verifica = False
  End If
  
  If t_fecha2 <> "" Then
    If Not IsDate(t_fecha2) Then
      verifica = False
    End If
  Else
    verifica = False
  End If
  
  If verifica = False Then
    MsgBox ("Error en las Fechas Ingresadas")
  End If
  
End Function
Sub carga()
Dim r As Integer
Call armagrid

Load espere
espere.Show
espere.Refresh
pb = 1
Set rs1 = New ADODB.Recordset
QUERY = "select * from A1 where [id_proveedor] > 1"
X = " and "

If t_cliente <> "" Then
  QUERY = QUERY & X & " [denominacion]  like '%" & t_cliente & "%'"
  X = " and "
End If

If Option1 = True Then
  QUERY = QUERY & " order by [id_proveedor]"
Else
  QUERY = QUERY & " order by [denominacion]"
End If
  
rs1.Open QUERY, cn1, adOpenStatic, adLockOptimistic, 1
If Not rs1.EOF And Not rs1.BOF Then
  saf = 0
  df = 0
  hf = 0
  sf = 0
  sof = 0
  r = 0
  While Not rs1.EOF
   Set cl_prov = New proveedores
   cl_prov.carga (rs1("id_proveedor"))
   saldoant = cl_prov.saldo(False, t_fecha, O_pesos, c_zona.ListIndex)
   saldoact = cl_prov.saldoentrefechas(t_fecha, t_fecha2, O_pesos, c_zona.ListIndex)
   If Val(Format$(saldoact, "######0.00")) = 0 And Val(Format$(saldoant, "######0.00")) = 0 And cl_prov.DEBE = 0 And cl_prov.HABER = 0 Then
     If O_cero = True Then
        Call agrega(r)
        r = r + 1
     End If
   Else
     If Val(Format$(saldoact, "######0.00")) > 0 Then
      If Check2 = 1 Then
        Call agrega(r)
        r = r + 1
      End If
     Else
      Call agrega(r)
      r = r + 1
     End If
   End If
   Set cl_prov = Nothing
   rs1.MoveNext
   pb = pb + 1
   Label5 = pb
   Label5.Refresh
   espere.Label1 = "Calculndo Saldo proveedor: " & r
   espere.Label1.Refresh
  Wend
  
  If Option5 = True Then
     
    msf1.col = 5 'Desde que columna iniciar la ordenacion
    msf1.ColSel = 5 'Hasta que columna terminar la ordenacion

    msf1.Row = 1 'Primer renglon del MsFlex a sortear
    msf1.RowSel = msf1.Rows - 1 'Ultimo renglon del msflex a sortear

    msf1.Sort = 3 'metodo de sorteo deseado Numerico ascendente
  End If
  
  msf1.Refresh
  linea = "____________________________________________________________"
  
  If Check1 = 0 Then
      msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & linea & Chr$(9) & linea & Chr$(9) & linea & Chr$(9) & linea
      msf1.AddItem "" & Chr$(9) & "Total Proveedores: " & r & Chr$(9) & Format$(saf, para.formato_numerico) & Chr$(9) & Format$(df, para.formato_numerico) & Chr$(9) & Format$(hf, para.formato_numerico) & Chr$(9) & Format$(sf, para.formato_numerico)
  Else
      msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & linea & Chr$(9) & linea & Chr$(9) & linea & Chr$(9) & linea & Chr$(9) & linea
      msf1.AddItem "" & Chr$(9) & "Total Proveedores: " & r & Chr$(9) & Format$(saf, para.formato_numerico) & Chr$(9) & Format$(df, para.formato_numerico) & Chr$(9) & Format$(hf, para.formato_numerico) & Chr$(9) & Format$(sf, para.formato_numerico) & Chr$(9) & Format$(sof, para.formato_numerico)
  End If

End If
Unload espere
Set rs1 = Nothing
Set rs2 = Nothing


End Sub

Sub carga2()
Dim r As Integer
Call armagrid

Load espere
espere.Show
espere.Refresh
pb = 1
Set rs1 = New ADODB.Recordset
QUERY = "select * from A1 where [id_proveedor] > 1 "
X = " and "

If t_cliente <> "" Then
  QUERY = QUERY & X & " [denominacion]  like '%" & t_cliente & "%'"
  X = " and "
End If

If Option1 = True Then
  QUERY = QUERY & " order by [id_proveedor]"
Else
  QUERY = QUERY & " order by [denominacion]"
End If
  
rs1.Open QUERY, cn1, adOpenStatic, adLockOptimistic, 1
If Not rs1.EOF And Not rs1.BOF Then
  saf = 0
  df = 0
  hf = 0
  sf = 0
  sof = 0
  r = 0
  While Not rs1.EOF
   Set cl_prov = New proveedores
   cl_prov.carga (rs1("id_proveedor"))
   saldoant = cl_prov.saldov(False, t_fecha, O_pesos, c_zona.ListIndex)
   saldoact = cl_prov.saldoentrefechasv(t_fecha, t_fecha2, O_pesos, c_zona.ListIndex)
   If Val(Format$(saldoact, "######0.00")) = 0 Then
     If O_cero = True Then
        Call agrega2(r)
        r = r + 1
     End If
   Else
     Call agrega2(r)
      r = r + 1
   End If
   Set cl_prov = Nothing
   rs1.MoveNext
   pb = pb + 1
   Label5 = pb
   Label5.Refresh
   espere.Label1 = "Calculndo Saldo proveedor: " & r
   espere.Label1.Refresh
  Wend
  If Check1 = 0 Then
      msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________"
      msf1.AddItem "" & Chr$(9) & "Total Proveedores: " & r & Chr$(9) & Format$(saf, "#####0.00") & Chr$(9) & Format$(df, "######0.00") & Chr$(9) & Format$(hf, "######0.00") & Chr$(9) & Format$(sf, "######0.00")
  Else
      msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________" & Chr$(9) & "________________"
      msf1.AddItem "" & Chr$(9) & "Total Proveedores: " & r & Chr$(9) & Format$(saf, "#####0.00") & Chr$(9) & Format$(df, "######0.00") & Chr$(9) & Format$(hf, "######0.00") & Chr$(9) & Format$(sf, "######0.00") & Chr$(9) & Format$(sof, "######0.00")
  End If

End If
Unload espere
Set rs1 = Nothing
Set rs2 = Nothing


End Sub
Sub agrega(r As Integer)
    saf = saf + saldoant
    df = df + cl_prov.DEBE
    hf = hf + cl_prov.HABER
    sf = sf + cl_prov.Saldoprov
    If Check1 = 0 Then
      msf1.AddItem cl_prov.idprov & Chr$(9) & cl_prov.razonsocial & Chr$(9) & Format$(saldoant, para.formato_numerico) & Chr$(9) & Format$(cl_prov.DEBE, para.formato_numerico) & Chr$(9) & Format$(cl_prov.HABER, para.formato_numerico) & Chr$(9) & Format$(cl_prov.Saldoprov, para.formato_numerico)
    Else
        d = cl_prov.DEBE
        h = cl_prov.HABER
        s = cl_prov.Saldoprov
        If O_pesos Then
          so = cl_prov.saldo(True, t_fecha2, False, c_zona.ItemData(c_zona.ListIndex))
        Else
          so = cl_prov.saldo(True, t_fecha2, True, c_zona.ItemData(c_zona.ListIndex))
        End If
        msf1.AddItem cl_prov.idprov & Chr$(9) & cl_prov.razonsocial & Chr$(9) & Format$(saldoant, para.formato_numerico) & Chr$(9) & Format$(d, para.formato_numerico) & Chr$(9) & Format$(h, para.formato_numerico) & Chr$(9) & Format$(s, para.formato_numerico) & Chr$(9) & Format$(so, para.formato_numerico)
        sof = sof + so
    End If
End Sub

Sub agrega2(r As Integer)
    saf = saf + saldoant
    df = df + cl_prov.DEBE
    hf = hf + cl_prov.HABER
    sf = sf + cl_prov.saldoprovv
    If Check1 = 0 Then
      msf1.AddItem cl_prov.idprov & Chr$(9) & cl_prov.razonsocial & Chr$(9) & Format$(saldoant, para.formato_numerico) & Chr$(9) & Format$(cl_prov.DEBE, para.formato_numerico) & Chr$(9) & Format$(cl_prov.HABER, para.formato_numerico) & Chr$(9) & Format$(cl_prov.saldoprovv, para.formato_numerico)
    Else
        d = cl_prov.DEBE
        h = cl_prov.HABER
        s = cl_prov.saldoprovv
        If O_pesos Then
          so = cl_prov.saldov(True, t_fecha2, False, c_zona)
        Else
          so = cl_prov.saldov(True, t_fecha2, True, c_zona)
        End If
        msf1.AddItem cl_prov.idprov & Chr$(9) & cl_prov.razonsocial & Chr$(9) & Format$(saldoant, para.formato_numerico) & Chr$(9) & Format$(d, para.formato_numerico) & Chr$(9) & Format$(h, para.formato_numerico) & Chr$(9) & Format$(s, para.formato_numerico) & Chr$(9) & Format$(so, para.formato_numerico)
        sof = sof + so
    End If
End Sub


Private Sub btnsale_Click()

Unload Me
End Sub





Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyUp
     Call tabup(Me)
End Select

End Sub
Sub armagrid()
'armar grilla
If Check1 = 0 Then
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 7
  msf1.ColWidth(0) = 1000
  msf1.ColWidth(1) = 6000
  msf1.ColWidth(2) = 2100
  msf1.ColWidth(3) = 2100
  msf1.ColWidth(4) = 2100
  msf1.ColWidth(5) = 2100
  msf1.ColWidth(6) = 700
    msf1.TextMatrix(0, 0) = "Id."
  msf1.TextMatrix(0, 1) = "Proveedor"
  msf1.TextMatrix(0, 2) = "Saldo Ant."
  If O_pesos Then
   msf1.TextMatrix(0, 3) = "Debe($)"
   msf1.TextMatrix(0, 4) = "Haber($)"
   msf1.TextMatrix(0, 5) = "Saldo($)"
  Else
   msf1.TextMatrix(0, 3) = "Debe(U$s)"
   msf1.TextMatrix(0, 4) = "Haber(U$s)"
   msf1.TextMatrix(0, 5) = "Saldo(U$s)"
  End If
  msf1.TextMatrix(0, 6) = " "
  For i = 0 To 6
    msf1.ColAlignment(i) = 9
  Next i
  msf1.ColAlignment(1) = 1
    
  
Else
   
  msf1.clear
  msf1.Rows = 1
  msf1.Cols = 8
  msf1.ColWidth(0) = 500
  msf1.ColWidth(1) = 5000
  msf1.ColWidth(2) = 2100
  msf1.ColWidth(3) = 2100
  msf1.ColWidth(4) = 2100
  msf1.ColWidth(5) = 2100
  msf1.ColWidth(6) = 2100
  msf1.ColWidth(7) = 500
  msf1.TextMatrix(0, 0) = "Id."
  msf1.TextMatrix(0, 1) = "Proveedor"
  msf1.TextMatrix(0, 2) = "Saldo Ant."
  If O_pesos Then
   msf1.TextMatrix(0, 3) = "Debe($)"
   msf1.TextMatrix(0, 4) = "Haber($)"
   msf1.TextMatrix(0, 5) = "Saldo($)"
   msf1.TextMatrix(0, 6) = "Saldo(U$s)"
  Else
   msf1.TextMatrix(0, 3) = "Debe(U$s)"
   msf1.TextMatrix(0, 4) = "Haber(U$s)"
   msf1.TextMatrix(0, 5) = "Saldo(U$s)"
   msf1.TextMatrix(0, 6) = "Saldo($)"
  End If
  msf1.TextMatrix(0, 7) = " "
  For i = 0 To 7
    msf1.ColAlignment(i) = 9
  Next i
  msf1.ColAlignment(1) = 1
 End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Is = 13
    Call TabEnter2(Me, 2)
  
End Select


End Sub

Private Sub Form_Load()
Load con_estadocuenta
Call barra(Me)

O_pesos = Checked
O_nocero = Checked

cal1.Visible = False
Check1 = 0
Call armagrid
Option1 = True
Option4 = True
Call carga_zonas(c_zona)
c_zona.AddItem "<Todas>", 0

End Sub




Private Sub Form_Unload(Cancel As Integer)
Unload con_estadocuenta
End Sub



Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1
  t_fecha.SetFocus
Else
  t_fecha2 = cal1
  t_fecha2.SetFocus
End If
cal1.Visible = False
End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel"

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    If O_pesos Then
      m = "Pesos ($)"
    Else
      m = "Dolares (U$s)"
    End If
    If Check1 = 0 Then
      c(0) = 6
      c(1) = 0
      c(2) = 1
      c(3) = 2
      c(4) = 3
      c(5) = 4
      c(6) = 5

      For i = 7 To 14
        c(i) = -1
      Next i
      Call imprimegrid(msf1, c(), "SALDOS PROVEEDORES", "Periodo: " & t_fecha & " - " & t_fecha2, "Moneda: " & m, " ", 72, 8, True, False)
  
  Else
      c(0) = 7
      c(1) = 0
      c(2) = 1
      c(3) = 2
      c(4) = 3
      c(5) = 4
      c(6) = 5
      c(7) = 6
      For i = 8 To 14
        c(i) = -1
      Next i
      Call imprimegrid(msf1, c(), "SALDOS PROVEEDORES", "Periodo: " & t_fecha & " - " & t_fecha2, "Moneda: " & m, " ", 72, 8, True, False)
  
   End If
  End If

End If

If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And msf1.Rows > 1 Then
   c = Val(msf1.TextMatrix(msf1.Row, 0))
   If c > 0 Then
      Load con_estadocuenta
      con_estadocuenta.c_prov.ListIndex = buscaindice(con_estadocuenta.c_prov, c)
      con_estadocuenta.Show
   End If
End If
End Sub

Private Sub O_dolares_Click()
Check1.Caption = "Muestra Saldo en $"
End Sub

Private Sub O_pesos_Click()
Check1.Caption = "Muestra Saldo en U$s"
End Sub

Private Sub t_cliente_GotFocus()
t_cliente = ""
End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = 1
cal1.SetFocus
End Sub

Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = 2
cal1.SetFocus

End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

'FIXIT: t_fecha2_LinkOpen event no tiene equivalente en Visual Basic .NET y no se actualizará.     FixIT90210ae-R7593-R67265
Private Sub t_fecha2_LinkOpen(Cancel As Integer)
Call solofecha(t_fecha2)
End Sub
