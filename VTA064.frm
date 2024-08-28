VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form vta_perc 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERCEPCIONES REALIZADAS por VENTAS"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   18315
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   7080
      TabIndex        =   11
      Top             =   120
      Width           =   8175
      Begin VB.ComboBox c_concepto 
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
         Left            =   2040
         TabIndex        =   14
         Top             =   840
         Width           =   5655
      End
      Begin VB.ComboBox c_imp 
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
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   5655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Concepto:"
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
         TabIndex        =   15
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Impuesto:"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   4320
      TabIndex        =   9
      Top             =   360
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   38666241
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3975
      Begin VB.TextBox t_fecha2 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox t_fecha 
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
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
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
         TabIndex        =   8
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
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
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   16320
      TabIndex        =   3
      Top             =   8280
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "VTA064.frx":0000
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
         Picture         =   "VTA064.frx":0882
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
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   9330
      Width           =   18315
      _ExtentX        =   32306
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   14111
            MinWidth        =   14111
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
      Height          =   6375
      Left            =   0
      TabIndex        =   10
      Top             =   1680
      Width           =   18135
      _ExtentX        =   31988
      _ExtentY        =   11245
      _Version        =   393216
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
End
Attribute VB_Name = "vta_perc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
Dim totperc As Double
Dim totret As Double
Const l = "-------------------------------------------"

Sub carga()
Call armagrid

q = "select * from vta_02, vta_016, i_01 where vta_02.num_int = vta_016.num_int and id_percepcion = id_impuesto and tipo_i1 = 'P'"
    
    If IsDate(t_fecha) Then
     q = q & " and datevalue([fecha]) >= datevalue('" & t_fecha & "')"
    End If
 
    If IsDate(t_fecha2) Then
      q = q & " and  datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
    End If
    
    
   If c_concepto.ListIndex > 0 Then
      q = q & " and  id_impuesto = " & c_concepto.ItemData(c_concepto.ListIndex)
   End If
      
      
   If c_imp.ListIndex > 0 Then
      q = q & " and  impuesto_i1 = '" & c_imp & "'"
   End If
   
   q = q & " order by fecha, id_impuesto"
    
    Set rs2 = New ADODB.Recordset
    rs2.Open q, cn1
    tcp = 0
    t = "0"
    bi = 0
    ali = 0
    While Not rs2.EOF
        F = Format$(rs2("fecha"), "dd/mm/yyyy")
        nc = rs2("letra") & " " & Format$(rs2("sucursal"), "0000") & "-" & Format$(rs2("num_comp"), "00000000")
        If rs2("moneda") = "P" Then
           c5 = 1
        Else
           c5 = rs2("cotiz_dolar")
        End If
        bi = rs2("base_imponible")
        ali = rs2("alicuota")
        If rs2("grabado") = "S" Then
          t = Format$(rs2("importe") * c5, "######0.00")
        Else
           t = Format$(-rs2("importe") * c5, "######0.00")
        End If
        tcodperc = tcodperc + Val(t)
        ti = ti + Val(t)
        tt = tt + Val(t)
        msf1.AddItem F & Chr(9) & rs2("detalle") & Chr$(9) & rs2("cliente02") & Chr(9) & rs2("cuit02") & " " & Chr(9) & " " & nc & Chr(9) & Format$(bi, para.formato_numerico) & Chr(9) & Format$(ali, para.formato_numerico) & Chr(9) & Format$(t, para.formato_numerico) & Chr(9) & Format$(rs2("vta_02.num_int"), "00000")
        rs2.MoveNext

    Wend
    Set rs2 = Nothing
    
    msf1.AddItem "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "TOTAL PERCEPCIONES " & Chr$(9) & "" & Chr(9) & Chr(9) & Format$(tt, para.formato_numerico)
         
       




  
  
   
   
   
End Sub





Private Sub btnacepta_Click()
espere.Show
espere.Refresh
Call carga

Unload espere

End Sub

Private Sub btnsale_Click()
Unload Me
End Sub







Private Sub cal1_DblClick()
If cal1.Tag = "1" Then
  t_fecha = cal1.Value
Else
  t_fecha2 = cal1.Value
End If
cal1.Visible = False
End Sub

Private Sub cal1_LostFocus()
cal1.Visible = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
   Case Is = vbKeyF12
     Unload Me
End Select
End Sub

Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 9
msf1.ColWidth(0) = 1500
msf1.ColWidth(1) = 2200
msf1.ColWidth(2) = 3100
msf1.ColWidth(3) = 2000
msf1.ColWidth(4) = 3300
msf1.ColWidth(5) = 2100
msf1.ColWidth(6) = 1500
msf1.ColWidth(7) = 2100
msf1.ColWidth(8) = 2300



msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Tipo Impuesto"
msf1.TextMatrix(0, 2) = "Cliente"
msf1.TextMatrix(0, 3) = "Cuit"
msf1.TextMatrix(0, 4) = "Tipo y Nro.Comprob."
msf1.TextMatrix(0, 5) = "BaseImp."
msf1.TextMatrix(0, 6) = "Alicuota"
msf1.TextMatrix(0, 7) = "Impuesto"
msf1.TextMatrix(0, 8) = "Num.Int."

For i = 0 To 4
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 5 To 7
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call cargaret
Call carga_percepciones_venta(c_concepto)
c_concepto.AddItem "<Todos>", 0
c_concepto.ListIndex = 0


End Sub
Sub cargaret()
'impuestos
c_imp.clear
c_imp.AddItem "<Todos>", 0
c_imp.AddItem "IVA", 1
c_imp.AddItem "IBBA", 2
c_imp.AddItem "GAN", 3
c_imp.AddItem "SEGSO", 4
c_imp.ListIndex = 0

End Sub


Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - "

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
    For i = 8 To 14
      c(i) = -1
    Next i
    Call imprimegrid(msf1, c(), "LISTADO DE PERCEPCIONES REALIZADAS por VENTAS", "", "Periodo: " & t_fecha & " : " & t_fecha2, " Impuesto:" & c_concepto & " ** " & c_imp & "**", 95, 6, True, False)
  End If

End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
     Load vta_cc_detalle
     vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 7)
     vta_cc_detalle.Show
  
  End If
End If

End Sub

Private Sub t_fecha_DblClick()
cal1.Visible = True
cal1.Tag = "1"


End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = Format$(Now, "dd/mm/yyyy")
  End If
End If
End Sub

Private Sub t_fecha2_DblClick()
cal1.Visible = True
cal1.Tag = "2"

End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = Format$(Now, "dd/mm/yyyy")
  End If
End If

End Sub
