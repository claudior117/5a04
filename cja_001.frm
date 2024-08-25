VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form cja_cajadiaria 
   BackColor       =   &H00E0E0E0&
   Caption         =   "CAJA DIARIA  CONTROL DE INGRESOS Y EGRESOS POR DIA"
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Caja de:"
      Height          =   735
      Left            =   8520
      TabIndex        =   34
      Top             =   6480
      Width           =   3255
      Begin VB.ComboBox c_usuario 
         Height          =   315
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SALDO DEL DIA"
      Height          =   855
      Left            =   5040
      TabIndex        =   32
      Top             =   0
      Width           =   2655
      Begin VB.TextBox T_SALDODIA 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Opciones"
      Height          =   855
      Left            =   4440
      TabIndex        =   29
      Top             =   7200
      Width           =   3855
      Begin VB.CheckBox Check3 
         Caption         =   "Agrupa Cobranzas en Cta.Cte"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   480
         Width           =   3615
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Agrupa Ventas Contado"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtros"
      Height          =   735
      Left            =   4440
      TabIndex        =   26
      Top             =   6480
      Width           =   3855
      Begin VB.ComboBox c_n1 
         Height          =   315
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rubros"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Filtros"
      Height          =   975
      Left            =   7800
      TabIndex        =   20
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox C_subrubro 
         Height          =   315
         Left            =   960
         TabIndex        =   24
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox C_rubro 
         Height          =   315
         Left            =   960
         TabIndex        =   23
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Concepto"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Saldos a la Fecha"
      Height          =   1695
      Left            =   240
      TabIndex        =   11
      Top             =   6480
      Width           =   3975
      Begin VB.TextBox t_salf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox t_e 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox t_i 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox t_sa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "Saldo a la Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Egresos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         Caption         =   "Ingresos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Saldo Anterior:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSComCtl2.MonthView Mv_1 
      Height          =   2370
      Left            =   4440
      TabIndex        =   8
      Top             =   2040
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      StartOfWeek     =   179306497
      CurrentDate     =   39157
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Fecha"
      Height          =   855
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cal"
         Height          =   195
         Left            =   2880
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   873
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox t_fecha 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5415
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   1
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "cja_001.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salir sin Modificar"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.CommandButton btnacepta 
         Height          =   615
         Left            =   120
         Picture         =   "cja_001.frx":0882
         Style           =   1  'Graphical
         TabIndex        =   2
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
      TabIndex        =   0
      Top             =   8235
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Cliente"
            TextSave        =   "Cliente"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "25/08/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "11:15 a.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   5415
      Left            =   6240
      TabIndex        =   10
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9551
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "cja_cajadiaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984


Private Sub btnacepta_Click()

Call limpia

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub



Private Sub C_RUBRO_LostFocus()
If C_rubro.ListIndex < 0 Then
   C_rubro.ListIndex = 0
End If
End Sub

Private Sub C_subrubro_LostFocus()
If C_subrubro.ListIndex < 0 Then
  If Val(C_subrubro) > 0 Then
    C_subrubro.ListIndex = buscaindice(C_subrubro, Val(C_subrubro))
  Else
    C_subrubro.ListIndex = 0
  End If
End If

End Sub


Private Sub Check1_Click()
If Check1 = 1 Then
  mv_1.Visible = True
Else
  mv_1.Visible = False
End If
End Sub

Private Sub Form_Activate()
If Check1 = 1 Then
  mv_1.Visible = True
Else
  mv_1.Visible = False
End If
 
If t_fecha = "" Then
  t_fecha = Format$(Now, "dd/mm/yyyy")
End If
Call armagrid
Call armagrid2
Call limpia


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF1
        'carga ingresos
   Call nivel_acceso(3)
   If para.id_grupo_modulo_actual >= 7 Then
      cyb_movcaja.limpia
      cyb_movcaja.c_tipo.ListIndex = 0
      cyb_movcaja.cargacuentas
      cyb_movcaja.t_fecha = t_fecha
      cyb_movcaja.Show
   Else
      Call sinpermisos
   End If
   Case Is = vbKeyF5
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 7 Then
     
     'carga engresos
      cyb_movcaja.limpia
      cyb_movcaja.t_fecha = t_fecha
      cyb_movcaja.c_tipo.ListIndex = 1
      cyb_movcaja.cargacuentas
      cyb_movcaja.Show
     Else
      Call sinpermisos
     End If
   
   Case Is = vbKeyF7
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 5 Then
      J = MsgBox("Prepare Impresora y Confirme", 4)
      If J = 6 Then
            Call imprime
      End If
    Else
      Call sinpermisos
     End If
   
  Case Is = vbKeyF9
    Call nivel_acceso(3)
    If para.id_grupo_modulo_actual >= 8 Then
      J = MsgBox("Cambia estado(Abre/Cierra) caja", 4)
      If J = 6 Then
         Call cambiaestado
            
      End If
    Else
      Call sinpermisos
     End If
   
   
   
   
   End Select



End Sub
Sub cambiaestado()
 q = "select * from  cyb_09 where datevalue([fecha]) = datevalue('" & t_fecha & "')"
 Set rs = New ADODB.Recordset
 rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 If Not rs.EOF And Not rs.BOF Then
   'existe
   If rs("Estado") = "A" Then
       rs("estado") = "C"
   Else
       rs("estado") = "A"
   End If
   rs.Update
 Else
   rs.AddNew
   rs("fecha") = Format$(DateValue(t_fecha), "dd/mm/yyyy")
   rs("estado") = "C"
   rs.Update
 End If
 Call limpia
End Sub
Sub cabecera()
  t = "________________________________________________________________________________"
  
  'Call imprimeempresa(14)
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.FontName = "Courier New"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.FontSize = 14
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print "Planilla de Caja Diaria del " & t_fecha
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.FontSize = 8
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print Tab(110); "Rubro.......:" & C_rubro
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print Tab(110); "Cuenta......:" & C_subrubro
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print t & "    " & t
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print "INGRESOS                      Detalle              Cuenta               Importe     EGRESOS                       Detalle              Cuenta               Importe"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Print t & "    " & t
  
End Sub
Sub imprime()
  'On Error GoTo erriMP
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
  Printer.Orientation = 2
  Call cabecera
  nh = 1
  lph = 45
  Fila = 1
  linea = 10
  t = "____________________________________________________________________________________________________"
  fi = msf1.Rows
  fe = msf2.Rows
  If fi >= fe Then
     tf = fi
  Else
     tf = fe
  End If
  While Fila < tf
   If linea <= lph Then
     Text = ""
     i = Space$(10)
     If Fila < fi Then
        r = Format$(Left$(msf1.TextMatrix(Fila, 4) & " ", 19), "@@@@@@@@@@@@@@@@@@@!")
        s = Format$(Left$(msf1.TextMatrix(Fila, 1) & " ", 19), "@@@@@@@@@@@@@@@@@@@!")
        d = Format$(Left$(msf1.TextMatrix(Fila, 0) & " ", 29), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
        RSet i = Format$(Val(msf1.TextMatrix(Fila, 2)), "######0.00")
        If Val(i) = 0 Then
          If msf1.TextMatrix(Fila, 4) <> "*" Then
            i = Space$(10)
          Else
            i = "----------"
          End If
        End If
     Else
        r = Space$(19)
        s = Space$(19)
        d = Space$(29)
     End If
     Text = d & "|" & r & "|" & s & "|" & i
     
     i = Space$(10)
     If Fila < fe Then
        r = Format$(Left$(msf2.TextMatrix(Fila, 4) & " ", 19), "@@@@@@@@@@@@@@@@@@@!")
        s = Format$(Left$(msf2.TextMatrix(Fila, 1) & " ", 19), "@@@@@@@@@@@@@@@@@@@!")
        d = Format$(Left$(msf2.TextMatrix(Fila, 0) & " ", 29), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
        RSet i = Format$(Val(msf2.TextMatrix(Fila, 2)), "######0.00")
        If Val(i) = 0 Then
          If msf2.TextMatrix(Fila, 4) <> "*" Then
            i = Space$(10)
          Else
            i = "----------"
          End If
        End If
     
     Else
        r = Space$(19)
        s = Space$(19)
        d = Space$(28)
     End If
     Text = Text & " || " & d & "|" & r & "|" & s & "|" & i
     Call imprimelinea(Text, 8, False, False, 1)
     Fila = Fila + 1
     linea = linea + 1
  Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print "________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.Print "Fecha Imp." & Format$(Now, "dd/mm/yyyy") & "   Nro.Hoja: " & Format$(nh, "000") & "     Emitido por: " & glo.usuario
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
     Printer.NewPage
     nh = nh + 1
     linea = 7
     Call cabecera
  End If

 Wend
 For J = Fila To lph
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.Print
 Next J
 RSet i = Format$(Val(t_sa), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(100); "========================================"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(100); "Saldo anterior............: " & i
 RSet i = Format$(Val(t_i), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(100); "  INGRESOS................: " & i
 RSet i = Format$(Val(t_e), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(100); "  EGRESOS.................: " & i
 RSet i = Format$(Val(t_salf), "######0.00")
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(100); "Saldo a la Fecha..........: " & i
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print Tab(100); "========================================"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print "________________________________________________________________"
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.Print "Fecha Imp." & Format$(Now, "dd/mm/yyyy") & "   Nro.Hoja: " & Format$(nh, "000") & "     Emitido por: " & glo.usuario
 
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
 Printer.EndDoc

Exit Sub
erriMP:
g = MsgBox("Error de Impresion. Continua?", 4)
If g = 6 Then
   Resume
Else
'FIXIT: El Asistente para actualización no actualiza el objeto Printer ni la colección Printers a Visual Basic .NET.     FixIT90210ae-R5481-H1984
   Printer.KillDoc
   Exit Sub
End If

End Sub
Sub limpia()


  Call armagrid
  Call armagrid2
  Label7 = estadocaja(t_fecha)
  If Label7 = "A" Then
    Label7.BackColor = &HC000&
    t_fecha.ForeColor = &HC000&
  Else
    Label7.BackColor = &HFF&
    t_fecha.ForeColor = &HFF&

  End If
  
  If C_rubro.ListIndex = 0 Then
    r = 0
  Else
    r = C_rubro.ItemData(C_rubro.ListIndex)
  End If
  
  If C_subrubro.ListIndex = 0 Then
    sr = 0
  Else
    sr = C_subrubro.ItemData(C_subrubro.ListIndex)
  End If
  
  
       Set rs2 = New ADODB.Recordset
       k = "select * from c_01 where [id_cuenta] = " & c_n1.ItemData(c_n1.ListIndex)
       rs2.Open k, cn1
       If Not rs2.EOF And Not rs2.BOF Then
      ci = Format$(rs2("pos1"), "0")
      cf = Format$(rs2("pos1"), "0")
      If rs2("pos2") > 0 Then
         ci = ci & Format$(rs2("pos2"), "0")
         cf = cf & Format$(rs2("pos2"), "0")
         If rs2("pos3") > 0 Then
           ci = ci & Format$(rs2("pos3"), "00")
           cf = cf & Format$(rs2("pos3"), "00")
         Else
           ci = ci & "00"
           cf = cf & "99"
         End If
      Else
         ci = ci & "000"
         cf = cf & "999"
      End If
    End If
    ci = ci & "00"
    cf = cf & "99"
    Set rs2 = Nothing
  
  c = 0

  If c_usuario.ListIndex > 0 Then
    t_sa = Format$(saldoanterior(t_fecha, c_usuario.ItemData(c_usuario.ListIndex), c), "######0.00")
    t_salf = Format$(saldoalafecha(t_fecha, c_usuario.ItemData(c_usuario.ListIndex), c), "######0.00")
  Else
    t_sa = Format$(saldoanterior(t_fecha, 0, c), "######0.00")
     t_salf = Format$(saldoalafecha(t_fecha, 0, c), "######0.00")
  End If
  
  'sub ingresos
  q = "SELECT * FROM cyb_01 where [caja] = 'S'"
  Set rs2 = New ADODB.Recordset
  rs2.Open q, cn1
  ti = 0
  te = 0
  While Not rs2.EOF
       
       q = "SELECT * FROM Cyb_05, cyb_01 WHERE DATEVALUE([FECHA]) = DATEVALUE('" & t_fecha & "') AND cyb_05.[ID_forma_pago] = " & rs2("ID_forma_pago") & " and  cyb_05.[id_forma_pago] = cyb_01.[id_forma_pago]"
       If C_rubro.ListIndex > 0 Then
           q = q & " and cyb_05.[id_forma_pago] = " & C_rubro.ItemData(C_rubro.ListIndex)
       End If
       
       If C_subrubro.ListIndex > 0 Then
           q = q & " and [id_cuenta_contra] = " & C_subrubro.ItemData(C_subrubro.ListIndex)
       End If
       
         If c_n1.ListIndex > 0 Then
       
            q = q & " and [id_cuenta_contra] > " & ci & " and [id_cuenta_contra] <= " & cf
  
         End If

       If c_usuario.ListIndex > 0 Then
         q = q & " and [id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
       End If
       
       q = q & " order by [num_mov_caja]"
         
       Set rs = New ADODB.Recordset
       rs.Open q, cn1
       tci = 0
       tce = 0
       PASADAi = 0
       pasadae = 0
       totalventasctdo = 0
       totalctacte = 0
       While Not rs.EOF
          d = rs("cyb_05.descripcion")
          o = rs("operacion")
          Set rs3 = New ADODB.Recordset
          q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta_contra")
          rs3.Open q, cn1
          If Not rs3.EOF And Not rs3.BOF Then
            cc = rs3("descripcion")
          Else
            cc = "Cuenta Inexistente"
          End If
          Set rs3 = Nothing
          i = Format$(rs("IMPORTE"), "######0.00")
          If rs("ubicacion") = "D" Then
            If PASADAi = 0 Then
              msf1.AddItem UCase(rs2("descripcion"))
              tci = 0
              PASADAi = 1
             End If
               
            If rs("modulo") = "V" Then
              If Mid$(rs("operacion"), 1, 3) <> "Rbo" Then
                'vta contado
                If Check2 = 1 Then
                   totalventasctdo = totalventasctdo + rs("Importe")
                Else
                   msf1.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
                End If
             Else
                 'vta ctacte
                If Check3 = 1 Then
                   totalctacte = totalctacte + rs("Importe")
                Else
                   msf1.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
                End If
             End If
            Else
               msf1.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
            End If
                
             tci = tci + rs("IMPORTE")
          
          Else
            If pasadae = 0 Then
             msf2.AddItem UCase(rs2("descripcion"))
             tce = 0
             pasadae = 1
            End If
            msf2.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
            tce = tce + rs("IMPORTE")
          End If
          rs.MoveNext
       Wend
       If totalventasctdo <> 0 Then
          msf1.AddItem "Ventas Contado Acumuladas" & Chr$(9) & "Ventas" & Chr$(9) & Format$(totalventasctdo, "#####0.00") & Chr$(9) & 0 & Chr$(9) & "Ventas Contado"
       End If
       If totalctacte <> 0 Then
          msf1.AddItem "Cobranzas en Cta. Cte. Acum." & Chr$(9) & "Deudores" & Chr$(9) & Format$(totalctacte, "#####0.00") & Chr$(9) & 0 & Chr$(9) & "Recibos"
       End If
       
       If tci > 0 Then
          msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "---------------" & Chr$(9) & Chr$(9) & "*"
          msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & Format$(tci, "######0.00")
          msf1.AddItem ""
          ti = ti + tci
       End If
             
       If tce > 0 Then
          msf2.AddItem "" & Chr$(9) & "" & Chr$(9) & "---------------" & Chr$(9) & Chr$(9) & "*"
          msf2.AddItem "" & Chr$(9) & "" & Chr$(9) & Format$(tce, "######0.00")
          msf2.AddItem ""
          te = te + tce
       End If
       
       t_i = Format$(ti, "#####0.00")
       t_e = Format$(te, "#####0.00")
       
       rs2.MoveNext
  Wend
  T_SALDODIA = Format$(Val(t_i) - Val(t_e), "######0.00")
  
  
  msf1.SetFocus
  
  
End Sub

Sub limpia2()
  Call armagrid
  Call armagrid2
  Label7 = estadocaja(t_fecha)
  If Label7 = "A" Then
    Label7.BackColor = &HC000&
    t_fecha.ForeColor = &HC000&
  Else
    Label7.BackColor = &HFF&
    t_fecha.ForeColor = &HFF&

  End If
  
  If C_rubro.ListIndex = 0 Then
    r = 0
  Else
    r = C_rubro.ItemData(C_rubro.ListIndex)
  End If
  
  If C_subrubro.ListIndex = 0 Then
    sr = 0
  Else
    sr = C_subrubro.ItemData(C_subrubro.ListIndex)
  End If
  
  
       Set rs2 = New ADODB.Recordset
       k = "select * from c_01 where [id_cuenta] = " & c_n1.ItemData(c_n1.ListIndex)
       rs2.Open k, cn1
       If Not rs2.EOF And Not rs2.BOF Then
      ci = Format$(rs2("pos1"), "0")
      cf = Format$(rs2("pos1"), "0")
      If rs2("pos2") > 0 Then
         ci = ci & Format$(rs2("pos2"), "0")
         cf = cf & Format$(rs2("pos2"), "0")
         If rs2("pos3") > 0 Then
           ci = ci & Format$(rs2("pos3"), "00")
           cf = cf & Format$(rs2("pos3"), "00")
         Else
           ci = ci & "00"
           cf = cf & "99"
         End If
      Else
         ci = ci & "000"
         cf = cf & "999"
      End If
    End If
    ci = ci & "00"
    cf = cf & "99"
    Set rs2 = Nothing
  
  
  If c_usuario.ListIndex > 0 Then
    t_sa = Format$(saldoanterior(t_fecha, 0, c, c_usuario.ItemData(c_usuario.ListIndex)), "######0.00")
    t_salf = Format$(saldoalafecha(t_fecha, 0, c, c_usuario.ItemData(c_usuario.ListIndex)), "######0.00")
  Else
    t_sa = Format$(saldoanterior(t_fecha, 0, c), "######0.00")
    t_salf = Format$(saldoalafecha(t_fecha, 0, c), "######0.00")
  End If
  
  'sub ingresos
  q = "SELECT * FROM cyb_01 where [caja] = 'S'"
  Set rs2 = New ADODB.Recordset
  rs2.Open q, cn1
  ti = 0
  te = 0
  While Not rs2.EOF
       
       q = "SELECT * FROM Cyb_05, cyb_01 WHERE DATEVALUE([FECHA]) = DATEVALUE('" & t_fecha & "') AND cyb_05.[ID_forma_pago] = " & rs2("ID_forma_pago") & " and  cyb_05.[id_forma_pago] = cyb_01.[id_forma_pago]"
       If C_rubro.ListIndex > 0 Then
           q = q & " and cyb_05.[id_forma_pago] = " & C_rubro.ItemData(C_rubro.ListIndex)
       End If
       
       If C_subrubro.ListIndex > 0 Then
           q = q & " and [id_cuenta_contra] = " & C_subrubro.ItemData(C_subrubro.ListIndex)
       End If
       
         If c_n1.ListIndex > 0 Then
       
            q = q & " and [id_cuenta_contra] > " & ci & " and [id_cuenta_contra] <= " & cf
  
         End If

       If c_usuario.ListIndex > 0 Then
           q = q & " and [id_usuario] = " & c_usuario.ItemData(c_usuario.ListIndex)
       End If
       
       q = q & " order by [id_cuenta_contra], [num_mov_caja]"
         
       Set rs = New ADODB.Recordset
       rs.Open q, cn1
       tci = 0
       tce = 0
       PASADAi = 0
       pasadae = 0
       totalventasctdo = 0
       totalctacte = 0
       
       While Not rs.EOF
          d = rs("cyb_05.descripcion")
          o = rs("operacion")
          Set rs3 = New ADODB.Recordset
          q = "select * from c_01 where [id_cuenta] = " & rs("id_cuenta_contra")
          rs3.Open q, cn1
          If Not rs3.EOF And Not rs3.BOF Then
            cc = rs3("descripcion")
          Else
            cc = "Cuenta Inexistente"
          End If
          Set rs3 = Nothing
          i = Format$(rs("IMPORTE"), "######0.00")
          If rs("ubicacion") = "D" Then
            If PASADAi = 0 Then
              msf1.AddItem UCase(rs2("descripcion"))
              tci = 0
              PASADAi = 1
             End If
               
            If rs("modulo") = "V" Then
              If Mid$(rs("operacion"), 1, 3) <> "Rbo" Then
                'vta contado
                If Check2 = 1 Then
                   totalventasctdo = totalventasctdo + rs("Importe")
                Else
                   msf1.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
                End If
             Else
                 'vta ctacte
                If Check3 = 1 Then
                   totalctacte = totalctacte + rs("Importe")
                Else
                   msf1.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
                End If
             End If
            End If
                
             tci = tci + rs("IMPORTE")
          
          Else
            If pasadae = 0 Then
             msf2.AddItem UCase(rs2("descripcion"))
             tce = 0
             pasadae = 1
            End If
            msf2.AddItem o & Chr$(9) & cc & Chr$(9) & i & Chr$(9) & rs("num_mov_caja") & Chr$(9) & d
            tce = tce + rs("IMPORTE")
          End If
          rs.MoveNext
       Wend
       If totalventasctdo > 0 Then
          msf1.AddItem "Ventas Contado Acumuladas" & Chr$(9) & "Ventas" & Chr$(9) & Format$(totalventasctdo, "#####0.00") & Chr$(9) & 0 & Chr$(9) & "Ventas Contado"
       End If
       If totalctacte > 0 Then
          msf1.AddItem "Cobranzas en Cta. Cte. Acum." & Chr$(9) & "Deudores" & Chr$(9) & Format$(totalctacte, "#####0.00") & Chr$(9) & 0 & Chr$(9) & "Recibos"
       End If
       
       If tci > 0 Then
          msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "---------------"
          msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & Format$(tci, "######0.00")
          msf1.AddItem ""
          ti = ti + tci
       End If
             
       If tce > 0 Then
          msf2.AddItem "" & Chr$(9) & "" & Chr$(9) & "---------------"
          msf2.AddItem "" & Chr$(9) & "" & Chr$(9) & Format$(tce, "######0.00")
          msf2.AddItem ""
          te = te + tce
       End If
       
       t_i = Format$(ti, "#####0.00")
       t_e = Format$(te, "#####0.00")
       
       rs2.MoveNext
  Wend
  T_SALDODIA = Format$(Val(t_i) - Val(t_e), "######0.00")
  
  
  msf1.SetFocus
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 5
msf1.FixedCols = 0
'msf1.SelectionMode = flexSelectionByRow
'msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 2400
msf1.ColWidth(1) = 1600
msf1.ColWidth(2) = 1100
msf1.ColWidth(3) = 900
msf1.ColWidth(4) = 2000
msf1.TextMatrix(0, 0) = "INGRESOS  "
msf1.TextMatrix(0, 1) = "Cuenta"
msf1.TextMatrix(0, 2) = "Importe"
msf1.TextMatrix(0, 3) = "Num.Mov."
msf1.TextMatrix(0, 4) = "Detalle"
msf1.ColAlignment(0) = 1 'izq
msf1.ColAlignment(1) = 1
msf1.ColAlignment(2) = 9 'der
msf1.ColAlignment(3) = 9 'der
msf1.ColAlignment(4) = 1
End Sub

Sub armagrid2()
'armar grilla
msf2.clear
msf2.Rows = 1
msf2.Cols = 5
msf2.FixedCols = 0
'msf2.SelectionMode = flexSelectionByRow
'msf2.FocusRect = flexFocusNone
msf2.ColWidth(0) = 2400
msf2.ColWidth(1) = 1600
msf2.ColWidth(2) = 1100
msf2.ColWidth(3) = 900
msf2.ColWidth(4) = 2000
msf2.TextMatrix(0, 0) = "EGRESOS  "
msf2.TextMatrix(0, 1) = "Cuenta"
msf2.TextMatrix(0, 2) = "Importe"
msf2.TextMatrix(0, 3) = "Num.Mov."
msf2.TextMatrix(0, 4) = "Detalle"

msf2.ColAlignment(0) = 1 'izq
msf2.ColAlignment(1) = 1
msf2.ColAlignment(2) = 9 'der
msf2.ColAlignment(3) = 9 'der
msf2.ColAlignment(4) = 1
End Sub


Private Sub Form_Load()
Load cyb_movcaja
Call barraesag(Me)
Check1 = 0
Check2 = 1
Check3 = 1
Call carga_cuentas_cont(C_subrubro, "C", "D")
C_subrubro.AddItem "<Todas>", 0
C_subrubro.ListIndex = 0

Call carga_formas_pago(C_rubro, "T")
C_rubro.AddItem "<Todas>", 0
C_rubro.ListIndex = 0

Call carga_cuentas_cont(c_n1, "T", "D")
c_n1.AddItem "<Todos>", 0
c_n1.ListIndex = 0

Call carga_usuarios(c_usuario)
c_usuario.AddItem "<General (Todas)>", 0
c_usuario.ListIndex = 0

Call barra


End Sub

Sub barra()
Me.StatusBar1.Panels.item(2) = "[F1] AGREGA INGRESOS - [F5] AGREGA EGRESOS - [F7] Imprime"

End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload cyb_movcaja
End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1] Ingresos - [F5] Egresos  - [F7] Imprime - [F8] Borra Mov. - [ENTER] Modifica - [F9] Abre/Cierra "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 8 Then
  n = Val(msf1.TextMatrix(msf1.Row, 3))
  If Val(n) > 0 Then
    If Label7 = "A" Then
      If verificaperiodog(t_fecha) = "A" Then
        Call borramovcaja(n)
        Call limpia
      Else
        MsgBox ("Periodo cerrado. Imposible grabar operacion")
      End If
    Else
      MsgBox ("La caja esta cerrada. Imposible realizar opoeracion")
    End If
  End If
 Else
  Call sinpermisos
 End If

End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 7 Then
 
  ni = Val(msf1.TextMatrix(msf1.Row, 3))
  If ni > 0 Then
      Call carga(ni)
  End If
 Else
    Call sinpermisos
 End If
 End If
End Sub

'FIXIT: Declare 'n' con un tipo de datos de enlace en tiempo de compilación                FixIT90210ae-R1672-R1B8ZE
Sub carga(ByVal n)
      cyb_movcaja.t_numint = n
      q = "SELECT * FROM cyb_05 WHERE [num_mov_caja] = " & n
      Set rs = New ADODB.Recordset
      rs.Open q, cn1
      If Not rs.EOF And Not rs.BOF Then
         cyb_movcaja.limpia
         cyb_movcaja.t_numint = rs("num_mov_caja")
         cyb_movcaja.t_fecha = rs("FECHA")
         cyb_movcaja.t_destino = rs("descripcion")
         cyb_movcaja.t_importe = rs("IMPORTE")
         cyb_movcaja.t_op = rs("operacion")
         cyb_movcaja.c_caja.ListIndex = buscaindice(cyb_movcaja.c_caja, rs("ID_forma_pago"))
         If rs("UBICACION") = "D" Then
            cyb_movcaja.c_tipo.ListIndex = 0
         Else
            cyb_movcaja.c_tipo.ListIndex = 1
         End If
         Set rs = Nothing
         cyb_movcaja.Show
      End If
End Sub
Private Sub msf1_LostFocus()
Call barra
End Sub

Private Sub msf2_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F1] Ingresos - [F5] Egresos  - [F7] Imprime - [F8] Borra Mov. - [ENTER] Modifica - [F9] Abre-Cierra"

End Sub

Private Sub msf2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF8 Then
 Call nivel_acceso(1)
 If para.id_grupo_modulo_actual > 8 Then
  n = Val(msf2.TextMatrix(msf2.Row, 3))
  If Val(n) > 0 Then
    If Label7 = "A" Then
       Call borramovcaja(n)
       Call limpia
    Else
      MsgBox ("La caja esta cerrada. Imposible realizar opoeracion")
    End If
  End If
 Else
  Call sinpermisos
 End If

End If

End Sub

Private Sub msf2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 Call nivel_acceso(3)
 If para.id_grupo_modulo_actual >= 7 Then
 
  ni = Val(msf2.TextMatrix(msf2.Row, 3))
  If ni > 0 Then
      Call carga(ni)
      
    End If
  Else
    Call sinpermisos
  End If
 End If
End Sub


Private Sub msf2_LostFocus()
Call barra
End Sub

Private Sub mv_1_DateDblClick(ByVal DateDblClicked As Date)
t_fecha = mv_1.Value
Call limpia
End Sub


Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If t_fecha <> "" Then
  If IsDate(t_fecha) Then
     Call limpia
  End If
 End If
End If
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

Private Sub UpDown1_DownClick()
t_fecha = DateValue(t_fecha) - 1
Call limpia
End Sub

Private Sub UpDown1_UpClick()
t_fecha = DateValue(t_fecha) + 1
Call limpia
End Sub
