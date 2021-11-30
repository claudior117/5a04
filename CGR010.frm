VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form cgr_admasientos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ADMINISTRADOR GENERAL DE ASIENTOS AUTOMATICOS"
   ClientHeight    =   8850
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordenado por"
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   4575
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Numero Interno"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   1695
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   11535
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Opciones"
         Height          =   615
         Left            =   6960
         TabIndex        =   16
         Top             =   960
         Width           =   3135
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Detallado"
            Height          =   255
            Left            =   1680
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Resumido"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.ComboBox c_usuarios 
         Height          =   315
         Left            =   8640
         TabIndex        =   13
         Text            =   "c_usuarios"
         Top             =   600
         Width           =   2775
      End
      Begin VB.ComboBox c_modulo 
         Height          =   315
         ItemData        =   "CGR010.frx":0000
         Left            =   8640
         List            =   "CGR010.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox t_fecha2 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox t_fecha 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox c_cuenta 
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Text            =   "c_cuenta"
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6960
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Modulo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6960
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Hasta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Fecha Desde:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800080&
         Caption         =   "Cuenta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Funciones"
      Height          =   975
      Left            =   9720
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
      Begin VB.CommandButton btnsale 
         Cancel          =   -1  'True
         Height          =   615
         Left            =   840
         Picture         =   "CGR010.frx":0070
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
         Picture         =   "CGR010.frx":08F2
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
      Top             =   8595
      Width           =   12090
      _ExtentX        =   21325
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   21167
            MinWidth        =   21167
            Text            =   "Sistema"
            TextSave        =   "Sistema"
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin MSFlexGridLib.MSFlexGrid msf1 
      Height          =   5175
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9128
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   1
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
Attribute VB_Name = "cgr_admasientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FIXIT: Utilice Option Explicit para evitar la creación implícita de variables de tipo Variant.     FixIT90210ae-R383-H1984

Sub cabecera()
List1.clear
'Call cabeceralist(List1)

End Sub

Sub carga()
  Call armagrid
  q = "select * from C_02, C_03 where C_02.[num_interno] = c_03.[num_interno] "
  c = " and "
  If c_cuenta.ListIndex > 0 Then
     q = q & c & " [id_cuenta] = " & c_cuenta.ItemData(c_cuenta.ListIndex)
  End If
  
  If IsDate(t_fecha) Then
     q = q & c & " datevalue([fecha]) >= datevalue('" & t_fecha & "')"
  End If
  
  If IsDate(t_fecha2) Then
     q = q & c & " datevalue([fecha]) <= datevalue('" & t_fecha2 & "')"
  End If
  
  If c_usuarios.ListIndex > 0 Then
     q = q & c & " [id_usuario] = " & c_usuarios.ItemData(c_usuarios.ListIndex)
  End If
  
  If c_modulo.ListIndex > 0 Then
     q = q & c & " [modulo] = '" & Mid$(c_modulo, 2, 1) & "'"
  End If
  
  If Option4 = True Then
    q = q & " order by c_02.[num_interno]"
  Else
    q = q & " order by [fecha], c_02.[num_interno]"
  End If
  
  Set rs = New ADODB.Recordset
  rs.Open q, cn1
  t = 0
  numint = 0
  While Not rs.EOF
   If numint <> rs("C_02.NUM_INTERNO") Then
      numint = rs("C_02.NUM_INTERNO")
      F = Format$(rs("fecha"), "dd/mm/yyyy")
      ope = rs("c_02.descripcion")
      ni = Format$(rs("c_02.num_interno"), "00000000")
      obs = rs("observaciones")
      di = Format$(rs("debe"), "#####0.00")
      hi = Format$(rs("haber"), "#####0.00")
      msf1.AddItem " " & Chr$(9) & ni & Chr$(9) & F & Chr$(9) & ope & Chr$(9) & di & Chr$(9) & hi & Chr$(9) & obs
      If Option2 = True Then
        Call armaasiento(numint)
      End If
   End If
   rs.MoveNext
  Wend
  
   
End Sub

Sub armaasiento(ByVal i As Long)
q = "select * from c_03, c_01 where [num_interno] = " & i & " and c_03.[id_cuenta] = c_01.[id_cuenta]"
q = q & " order by [ubicacion], [renglon]"
Set rs2 = New ADODB.Recordset
rs2.Open q, cn1
While Not rs2.EOF
   ic = rs2("c_03.id_cuenta")
   dc = Format$(Left$(rs2("c_01.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   d = Format$(rs2("c_03.importe"), "######0.00")
   If rs2("ubicacion") = "D" Then
     id = d
     ih = " "
     dc = Format$(Left$("----" & rs2("c_01.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   Else
     id = " "
     ih = d
     dc = Format$(Left$("--------" & rs2("c_01.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   End If
   o = Format$(Left$(rs2("c_03.descripcion"), 35), "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@!")
   msf1.AddItem "" & Chr$(9) & "" & Chr$(9) & "[" & ic & "]" & Chr$(9) & dc & Chr$(9) & id & Chr$(9) & ih & Chr$(9) & o
 rs2.MoveNext
Wend
Set rs2 = Nothing
End Sub
Private Sub btnacepta_Click()
Call carga
End Sub

Private Sub btnsale_Click()
Unload Me
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

Private Sub c_modulo_LostFocus()
If c_modulo.ListIndex < 0 Then
  c_modulo.ListIndex = 0
End If

End Sub

Private Sub c_usuarios_LostFocus()
If c_usuarios.ListIndex < 0 Then
  c_usuarios.ListIndex = 0
End If

End Sub

Private Sub Form_Load()


Call carga_cuentas_cont(c_cuenta, "C", "D")
c_cuenta.AddItem "<Todos>", 0
c_cuenta.ListIndex = 0

Call carga_usuarios(c_usuarios)
c_usuarios.AddItem "<Todos>", 0
c_usuarios.ListIndex = 0

c_modulo.ListIndex = 0

Call barraesag(Me)
Call armagrid
Option1 = True
Option3 = True
End Sub
Sub armagrid()
'armar grilla
msf1.clear
msf1.Rows = 1
msf1.Cols = 7
msf1.AllowUserResizing = flexResizeNone
msf1.FixedCols = 0
msf1.SelectionMode = flexSelectionByRow
msf1.FocusRect = flexFocusNone
msf1.ColWidth(0) = 300
msf1.ColWidth(1) = 1200
msf1.ColWidth(2) = 1200
msf1.ColWidth(3) = 4000
msf1.ColWidth(4) = 1200
msf1.ColWidth(5) = 1200
msf1.ColWidth(6) = 2000

msf1.TextMatrix(0, 0) = " "
msf1.TextMatrix(0, 1) = "Id.asiento"
msf1.TextMatrix(0, 2) = "Fecha"
msf1.TextMatrix(0, 3) = "Descricion / Cuenta"
msf1.TextMatrix(0, 4) = "Debe"
msf1.TextMatrix(0, 5) = "Haber"
msf1.TextMatrix(0, 6) = "Obs"

'msf1.FocusRect = flexFocusNone
For i = 0 To 5
  msf1.ColAlignment(i) = 9 'der
Next i
msf1.ColAlignment(3) = 1
msf1.ColAlignment(6) = 1

End Sub



Private Sub msf1_GotFocus()
Me.StatusBar1.Panels.item(1) = "[F3] Asiento Resumen.  - [F4]Saca As. - [F5]Marca Todos -  [F6]Marca As. -  [F8]Borra As. -[F7] Imprime   "
msf1.FocusRect = flexFocusHeavy
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
    
    For i = 7 To 14
      c(i) = -1
    Next i
    
    If Option2 = True Then
      t = "Detallado"
    Else
      t = "Resumido"
    End If
    
    Call imprimegrid(msf1, c(), "Lista de Asientos Provisorios", "", "Periodo......: " & t_fecha & "  " & t_fecha2, "Tipo.........:" & t, 85, 7, True, False)
  
  End If
    
    
End If

If KeyCode = vbKeyF8 Then
 Call nivel_acceso(2)
 If para.id_grupo_modulo_actual >= 8 Then
   If msf1.Row > 0 Then
     J = MsgBox("Confirma Eliminar Asientos Marcados ", 4)
     If J = 6 Then
       For i = 1 To msf1.Rows - 1
          If Val(msf1.TextMatrix(i, 1)) > 0 Then
              If Mid$(msf1.TextMatrix(i, 0), 1, 2) = "**" Then
                  nicgr = Val(msf1.TextMatrix(i, 1))
                  cn1.BeginTrans
                   QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & nicgr
                   cn1.Execute QUERY
      
                   QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & nicgr
                   cn1.Execute QUERY
                   cn1.CommitTrans
               End If
           End If
       Next i
     End If
  End If
 Else
   Call sinpermisos
 End If
End If

If KeyCode = vbKeyF3 Then 'asiento resumen
     
   espere.Show
   espere.Label1 = "Armando Asiento...."
   espere.Refresh
   
   If msf1.Rows > 1 Then
         Load abm_asientos
         abm_asientos.limpia
         If t_fecha <> "" Then
            abm_asientos.t_f1 = t_fecha
        Else
            abm_asientos.t_f1 = ""
        End If
        abm_asientos.t_descripciong = "Asiento Resumen " & abm_asientos.t_f1
        abm_asientos.t_funcion = "A"
        For i = 1 To msf1.Rows - 1
          If Val(msf1.TextMatrix(i, 1)) > 0 Then
             If Mid$(msf1.TextMatrix(i, 0), 1, 2) = "**" Then
                Call agrega(Val(msf1.TextMatrix(i, 1)))
             End If
          End If
        Next i
       
        abm_asientos.calcula_totales
        abm_asientos.Show
        Unload espere
   End If
End If

If KeyCode = vbKeyF5 Then
  J = MsgBox("Confirma seleccionar todos los asientos", 4)
  If J = 6 Then
   If msf1.Rows > 1 Then
        For i = 1 To msf1.Rows - 1
          If Val(msf1.TextMatrix(i, 1)) > 0 Then
             If Mid$(msf1.TextMatrix(i, 0), 1, 2) = "**" Then
                msf1.TextMatrix(i, 0) = " "
             Else
                msf1.TextMatrix(i, 0) = "**"
             End If
          End If
        Next i
   End If
  End If
End If
 

If KeyCode = vbKeyF6 Then
   If msf1.Row > 0 Then
     If Val(msf1.TextMatrix(msf1.Row, 1)) > 0 Then
       If Mid$(msf1.TextMatrix(msf1.Row, 0), 1, 2) = "**" Then
                msf1.TextMatrix(msf1.Row, 0) = " "
        Else
                msf1.TextMatrix(msf1.Row, 0) = "**"
       End If
     End If
   End If
End If

If KeyCode = vbKeyF4 Then
   If msf1.Row > 0 Then
     If Val(msf1.TextMatrix(msf1.Row, 1)) > 0 Then
       msf1.RemoveItem msf1.Row
       e = 0
       While e = 0
         If Val(msf1.TextMatrix(msf1.Row, 1)) = 0 Then
           'remuevo
            msf1.RemoveItem msf1.Row
         Else
            e = 1
         End If
       Wend
     End If
   End If
End If





End Sub
Sub pasaasientos()
 On Error GoTo ERRORGRABA
 espere.Show
 espere.ProgressBar1.Min = 0
 espere.ProgressBar1.Max = msf1.Rows
 espere.Refresh
 For i = 1 To msf1.Rows - 1
     espere.ProgressBar1.Value = i
    If Mid$(msf1.TextMatrix(i, 0), 1, 2) = "**" Then
      'el asiento sera pasado
      
      'saco numero
        a = Format$(Val(Mid$(msf1.TextMatrix(i, 2), 7, 4)), "0000")
        m = Format$(Val(Mid$(msf1.TextMatrix(i, 2), 4, 2)), "00")
       a1 = Val(a & m & "000")
       a2 = Val(a & m & "999")
       Set rs = New ADODB.Recordset
       q = "select * from c_11 where [año] = " & Val(a) & " and [mes] = " & Val(m)
       rs.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs.EOF And Not rs.BOF Then
         rs.MoveLast
         na = rs("num_asiento") + 1
       Else
         na = Val(a & m & "001")
       End If
       Set rs = Nothing
      
       'busco asiento temporal
       numint2 = Val(msf1.TextMatrix(i, 1))
       q = "select * from c_02, c_03 where c_02.[num_interno] = " & numint2 & " and c_02.[num_interno] = c_03.[num_interno]"
       Set rs1 = New ADODB.Recordset
       rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
       If Not rs1.EOF And Not rs1.BOF Then
          'grabo aisento
           cn1.BeginTrans
           QUERY = "INSERT INTO c_11([num_asiento], [fecha], [descripcion], [id_periodo], [importe], [año], [mes])"
           QUERY = QUERY & " VALUES (" & na & ", '" & rs1("fecha") & "', '" & rs1("c_02.descripcion") & "', " & para.id_periodo_contable & ", " & rs1("debe") & ", " & Val(a) & ", " & Val(m) & ")"
           cn1.Execute QUERY
      
           qr = "SELECT @@IDENTITY AS NewID"
           Set rs = cn1.Execute(qr)
           nic = rs.Fields("NewID").Value
       
           'grabo cuentas
           s = 1
           While Not rs1.EOF
             QUERY = "INSERT INTO c_12([id_asiento], [secuencia], [id_cuenta], [importe], [descripcion], [ubicacion])"
             QUERY = QUERY & " VALUES (" & nic & ", " & s & ", " & rs1("id_cuenta") & ", " & rs1("importe") & ", '" & rs1("c_03.descripcion") & "', '" & rs1("ubicacion") & "')"
             cn1.Execute QUERY
             s = s + 1
             rs1.MoveNext
           Wend
          Set rs1 = Nothing
          
          'borro asiento temporal
          ' query = "delete"
           QUERY = "DELETE FROM c_03 WHERE [num_interno] = " & numint2
           cn1.Execute QUERY
          
           QUERY = "DELETE FROM c_02 WHERE [num_interno] = " & numint2
           cn1.Execute QUERY
          
       End If
       
       cn1.CommitTrans
    End If
 Next i
 Unload espere
Exit Sub

ERRORGRABA:
  cn1.RollbackTrans
  Unload espere
  MsgBox ("Error en el Proceso de Traspaso. Toda la operacion ha sido cancelada")
  
  Exit Sub
End Sub
Private Sub t_fecha_GotFocus()
t_fecha = ""
End Sub

Private Sub t_fecha_LostFocus()
If t_fecha <> "" Then
  If Not IsDate(t_fecha) Then
    t_fecha = ""
  End If
End If
End Sub

Private Sub t_fecha2_GotFocus()
t_fecha2 = ""
End Sub

Private Sub t_fecha2_LostFocus()
If t_fecha2 <> "" Then
  If Not IsDate(t_fecha2) Then
    t_fecha2 = ""
  End If
End If

End Sub
Sub agrega(ByVal idas)
'busco asiento
Set rs = New ADODB.Recordset
q = "SELECT * FROM c_01, C_02, C_03 WHERE C_02.[num_interno] = " & idas & " and c_02.[num_interno] = c_03.[num_interno] and c_01.[id_cuenta] = c_03.[id_cuenta]"
'MsgBox (q)
rs.Open q, cn1
While Not rs.EOF
  cod = rs("c_03.id_cuenta")
  Detalle = rs("c_02.descripcion")
  importe = rs("c_03.importe")
  cuenta = rs("c_01.descripcion")
  o = rs("observaciones")
  If rs("ubicacion") = "D" Then
       'debe
       abm_asientos.msf1.AddItem abm_asientos.msf1.Rows & Chr$(9) & cod & Chr$(9) & Detalle & Chr$(9) & Format$(importe, "######0.00") & Chr$(9) & cuenta & Chr$(9) & o
  Else
       'haber
       abm_asientos.msf2.AddItem abm_asientos.msf2.Rows & Chr$(9) & cod & Chr$(9) & Detalle & Chr$(9) & Format$(importe, "######0.00") & Chr$(9) & cuenta & Chr$(9) & o
  End If
  rs.MoveNext
Wend
Set rs = Nothing
End Sub
