VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form vta_movprodcli 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "INFORME DE MOVIMIENTOS DE UN PRODUCTOS (UNIDADES)"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   12240
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registros"
      Height          =   735
      Left            =   6480
      TabIndex        =   23
      Top             =   1080
      Width           =   1335
      Begin VB.TextBox t_encontrados 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   13
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   7920
      TabIndex        =   20
      Top             =   1080
      Width           =   3735
      Begin VB.ComboBox c_op 
         Height          =   315
         ItemData        =   "VTA022.frx":0000
         Left            =   1200
         List            =   "VTA022.frx":0013
         TabIndex        =   21
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Operacion:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   7575
      Begin VB.ComboBox c_prod 
         Height          =   315
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Caption         =   "Producto:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   6255
      Begin VB.ComboBox c_vend 
         Height          =   315
         Left            =   1440
         TabIndex        =   15
         Top             =   600
         Width           =   4575
      End
      Begin VB.ComboBox c_prov 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Vendedor:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cliente:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComCtl2.MonthView cal1 
      Height          =   2370
      Left            =   9120
      TabIndex        =   9
      Top             =   1200
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   14737632
      Appearance      =   1
      StartOfWeek     =   172949505
      CurrentDate     =   38750
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   8040
      TabIndex        =   6
      Top             =   0
      Width           =   3615
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   3
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "VTA022.frx":0053
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
         Picture         =   "VTA022.frx":08D5
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
      Top             =   8415
      Width           =   12240
      _ExtentX        =   21590
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
            TextSave        =   "28/02/2024"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:14 p.m."
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5055
      Left            =   0
      TabIndex        =   10
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   8916
      _Version        =   393216
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4200
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "vta_movprodcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984
'FIXIT: Declare 'ti' con un tipo de datos de enlace en tiempo de compilación               FixIT90210ae-R1672-R1B8ZE
Dim ti, t As Double
'FIXIT: Declare 'reg' con un tipo de datos de enlace en tiempo de compilación              FixIT90210ae-R1672-R1B8ZE
Dim reg, regi As Integer


Sub carga()
  Call armagrid
 'busco el producto en las ventas
  q = "select * from vta_02, vta_03, vta_01, vta_06 where [id_producto] = " & c_prod.ItemData(c_prod.ListIndex) & " and vta_03.[num_int] = vta_02.[num_int] and vta_02.[id_cliente] = vta_01.[id_cliente]  and vta_02.[id_tipocomp] = vta_06.[id_tipocomp] and vta_02.[sucursal] = vta_06.[sucursal]"
  c = " and "
  
  If c_prov.ListIndex > 0 Then
       q = q & c & " vta_02.[id_cliente] = " & c_prov.ItemData(c_prov.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
       q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
       
  If IsDate(t_fecha2) Then
       q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_vend.ListIndex > 0 Then
       q = q & c & " vta_02.[Id_vendedor] = " & c_vend.ItemData(c_vend.ListIndex)
  End If
  
  If c_op.ListIndex > 0 Then
     Select Case c_op.ListIndex
      Case Is = 1
        q = q & c & " vta_02.[id_tipocomp] = 1"
      Case Is = 2
        q = q & c & " vta_02.[id_tipocomp] = 45"
      Case Is = 3
        q = q & c & " (vta_02.[id_tipocomp] = 1 or vta_02.[id_tipocomp] = 3) "
      Case Is = 4
         q = q & c & " vta_02.[id_tipocomp] >= 45 and vta_02.[id_tipocomp] <= 46 "
     End Select
     c = " and "
   End If
               
   
  q = q & " order by [fecha], vta_02.[id_cliente], vta_02.[id_tipocomp], [num_comp]"
        
  Set rs2 = New ADODB.Recordset
  rs2.Open q, cn1
      
   ttr = 0
   ttf = 0
   ttp = 0
   ttd = 0
   reg = 0
   ttnc = 0
   te = 0
   While Not rs2.EOF
       tr = 0
       tf = 0
       tp = 0
       td = 0
       tnc = 0
      
        reg = reg + 1
        Label8 = reg
        Label8.Refresh
        
        Select Case rs2("vta_02.id_tipocomp")
          Case Is = 1, Is = 2 'fact y nd
            tf = tf + rs2("cantidad")
            ttf = ttf + rs2("cantidad")
          Case Is = 3 'nc
            tnc = tnc + rs2("cantidad")
            ttnc = ttnc + rs2("cantidad")
          Case Is = 45 'rtos
            tr = tr + rs2("cantidad_original")
            ttr = ttr + rs2("cantidad_original")
            tp = tp + (rs2("cantidad"))
            ttp = ttp + (rs2("cantidad"))
          Case Is = 46 'dev
            td = td + rs2("cantidad")
            ttd = ttd + rs2("cantidad")
        End Select
        
        F = rs2("fecha")
        cli = rs2("denominacion")
        comp = rs2("abreviatura") & " " & rs2("letra") & " " & Format$(rs2("vta_02.sucursal"), "0000") & "-" & Format$(rs2("num_comp"), "00000000")
        ni = rs2("vta_02.num_int")
        
        
        If rs2("vta_02.id_tipocomp") = 1 Then
         'busco remitos aplicados
          Set rs3 = New ADODB.Recordset
          q = "select * from vta_08, vta_02 where [id_factura] = " & rs2("vta_02.num_int") & " and [id_remito] = [num_int]"
          rs3.Open q, cn1
          trr = ""
          While Not rs3.EOF
            trr = trr & Format$(rs3("sucursal"), "0000") & "-" & Format$(rs3("num_comp"), "00000000") & "*"
            rs3.MoveNext
          Wend
          Set rs3 = Nothing
        Else
          trr = ""
        End If
        
        If tr > 0 Or tf > 0 Or tnc > 0 Or tp > 0 Or td > 0 Then
          te = te + 1
          t_encontrados = te
          t_encontrados.Refresh
          msf1.AddItem F & Chr(9) & cli & Chr(9) & comp & Chr(9) & Format$(tr, "#####0.00") & Chr(9) & Format$(td, "#####0.00") & Chr(9) & Format$(tf, "#####0.00") & Chr(9) & Format$(tnc, "######0.00") & Chr$(9) & Format$(tp, "#####0.00") & Chr(9) & ni & Chr(9) & trr
        End If
        rs2.MoveNext
      Wend
      Set rs2 = Nothing
      
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________" & Chr(9) & "_____________________"
   msf1.AddItem "" & Chr(9) & "" & Chr(9) & " " & Chr(9) & Format$(ttr, "#####0.00") & Chr(9) & Format$(ttd, "#####0.00") & Chr(9) & Format$(ttf, "#####0.00") & Chr(9) & Format$(ttnc, "#####0.00") & Chr(9) & Format$(ttp, "#####0.00")
   
  
   
   
   
End Sub
Private Sub btnacepta_Click()
   Call carga
  
End Sub

Private Sub btnsale_Click()
Unload Me
End Sub





Private Sub c_op_LostFocus()
If c_op.ListIndex < 0 Then
  c_op.ListIndex = 0
End If
End Sub

Private Sub c_prod_LostFocus()
If c_prod.ListIndex < 0 Then
  c_prod.ListIndex = 0
End If
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



Private Sub Command5_Click()

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
msf1.Cols = 11
msf1.ColWidth(0) = 1000
msf1.ColWidth(1) = 3000
msf1.ColWidth(2) = 2200
msf1.ColWidth(3) = 1100
msf1.ColWidth(4) = 1100
msf1.ColWidth(5) = 1100
msf1.ColWidth(6) = 1100
msf1.ColWidth(7) = 1100
msf1.ColWidth(8) = 1000
msf1.ColWidth(9) = 3500
msf1.ColWidth(10) = 700
msf1.TextMatrix(0, 0) = "Fecha"
msf1.TextMatrix(0, 1) = "Cliente"
msf1.TextMatrix(0, 2) = "Comprobante"
msf1.TextMatrix(0, 3) = "Remitido"
msf1.TextMatrix(0, 4) = "Devoluc."
msf1.TextMatrix(0, 5) = "Facturado"
msf1.TextMatrix(0, 6) = "N.C."
msf1.TextMatrix(0, 7) = "Pendiente"
msf1.TextMatrix(0, 8) = "Num.Int."
msf1.TextMatrix(0, 9) = "Aplicados"
msf1.TextMatrix(0, 10) = ""

For i = 0 To 2
  msf1.ColAlignment(i) = 1 'izq
Next i
For i = 3 To 6
  msf1.ColAlignment(i) = 9 'der
Next i

End Sub

Private Sub Form_Load()
Call barraesag(Me)
cal1.Visible = False
Call armagrid
Call carga_clientes(c_prov)
c_prov.AddItem "<Todos>", 0
c_prov.ListIndex = 0

Call carga_vendedores(c_vend)
c_vend.AddItem "<Todos>", 0
c_vend.ListIndex = 0

Call carga_productos(c_prod)
c_prod.ListIndex = 0

c_op.ListIndex = 0

Option1 = True

End Sub




Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(2) = "[F7] Imprime - [F11] Excel "

End Sub

Private Sub msf1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF7 Then
  Dim c(15) As Double
  J = MsgBox("Prepare Impresora y confirme", 4)
  If J = 6 Then
    c(0) = 10
    c(1) = 0
    c(2) = 1
    c(3) = 2
    c(4) = 3
    c(5) = 4
    c(6) = 5
    c(7) = 6
    c(8) = 7
    For i = 9 To 14
      c(i) = -1
    Next i
    tt = "Cliente: " & c_prov & "     Vendedor: " & c_vend
    Call imprimegrid(msf1, c(), "INFORME DE MOVIMIENTOS DE UN PRODUCTO", "Producto: " & c_prod, "Periodo: " & t_fecha & " : " & t_fecha2, tt, 85, 7, True, False)
  End If

End If


If KeyCode = vbKeyF11 Then
  Call exportaexcel(msf1)
End If

End Sub

Private Sub msf1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If msf1.Row > 0 Then
    Load cc_detalle
    vta_cc_detalle.t_numint = msf1.TextMatrix(msf1.Row, 8)
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
