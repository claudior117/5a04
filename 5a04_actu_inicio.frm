VERSION 5.00
Begin VB.Form actu_inicio 
   Caption         =   "ACTUALIZADOR DEL SISTEMA  GestionE"
   ClientHeight    =   2655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "Opcion 2"
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Opcion 1"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ACTUALIZAR"
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ultima Version sistema anterior 181.GestionE11 Factura Electronica empieza en 201"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   7455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Ultima Actualizacion instalada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4440
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "163"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "Ultima Actualizacion disponible"
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
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "230"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "actu_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
J = InputBox$("Ingrese Password de Administrador General")
prueba = "N"
If J = "1975" Then
 If Option1 = True Then
   o = 1
 Else
   o = 2
 End If
 
 x = InputBox$("Ingrese Numero de Axctualizacion a realizar")
 
 
 If abrirconexion(o) Then
  
  Select Case Val(x)
  Case Is = 90
   Call actu90
  Case Is = 91
   Call actu91
  Case Is = 92
   Call actu92
  Case Is = 93
   Call actu93
  Case Is = 94
   Call actu94
  Case Is = 95
   Call actu95
  Case Is = 98
   Call actu98
  Case Is = 101
   Call actu101
  Case Is = 102
   Call actu102
  Case Is = 103
   Call actu103
  Case Is = 104
   Call actu104
  Case Is = 106
   Call actu106
  Case Is = 108
   Call actu108
  Case Is = 109
   Call actu109
  Case Is = 111
   Call actu111
  Case Is = 114
   Call actu114
  Case Is = 116
   Call actu116
  
  Case Is = 119
   Call actu119
  Case Is = 123
   Call actu123
  Case Is = 124
   Call actu124
 Case Is = 125
   Call actu125
 Case Is = 128
   Call actu128
 Case Is = 129
   Call actu129
 Case Is = 130
   Call actu130
 Case Is = 132
   Call actu132
 Case Is = 133
   Call actu133
 Case Is = 134
   Call actu134
 Case Is = 136
   Call actu136
 Case Is = 138
   Call actu138
 Case Is = 140
   Call actu140
 Case Is = 141
   Call actu141
 Case Is = 142
   Call actu142
 Case Is = 145
   Call actu145
Case Is = 148
   Call actu148
 Case Is = 1482
   Call actu1482
 Case Is = 149
   Call actu149
 Case Is = 156
   Call actu156
 Case Is = 158
   Call actu158
 Case Is = 159
   Call actu159
 Case Is = 160
   Call actu160
 Case Is = 162
   Call actu162
  Case Is = 163
   Call actu163
  Case Is = 164
   Call actu164
  Case Is = 165
   Call actu165
   Case Is = 166
   Call actu166
  Case Is = 167
   Call actu167
  Case Is = 168
   Call actu168
 Case Is = 169
   Call actu169
 Case Is = 170
   Call actu170
 Case Is = 171
   Call actu171
 Case Is = 172
   Call actu172
 Case Is = 173
   Call actu173
 Case Is = 174
   Call actu174
 Case Is = 176
   Call actu176
 Case Is = 177
   MsgBox ("Incluir archivo firma1.jpg de 175x175 a la carpeta del sistema -->tools. Si no usa firma incluir archivo con imagen en blanco ")
   cn1.BeginTrans
   q = "update g0 set  [actualizacion]=177"
   q = q & " where [sucursal]=0 "
   cn1.Execute q
   cn1.CommitTrans
Case Is = 178
   Call actu178
 Case Is = 179
    Call actu179
 Case Is = 180
    Call actu180
 Case Is = 181
    Call actu181
 Case Is = 201
    Call actu201
 Case Is = 202
    Call actu202
 Case Is = 203
    Call actu203
 Case Is = 204
    Call actu204
 Case Is = 205
    Call actu205
 Case Is = 206
    Call actu206
 Case Is = 207
    Call actu207
Case Is = 208
    Call actu208
 Case Is = 210
    Call actu210
Case Is = 211
    Call actu211
 Case Is = 212
    Call actu212
 Case Is = 213
    Call actu213
Case Is = 215
    Call actu215
 Case Is = 216
    Call actu216
 Case Is = 217
    Call actu217
 Case Is = 218
    Call actu218
 Case Is = 219
    Call actu219
 Case Is = 220
    Call actu220
 Case Is = 221
    Call actu221
 Case Is = 222
    Call actu222
 Case Is = 223
    Call actu223
 Case Is = 224
    Call actu224
 Case Is = 225
    Call actu225 'percepcion iva
 Case Is = 226
    Call actu226 'resolucion minima pantalla 1024x768
Case Is = 227
    Call actu227
Case Is = 228
    Call actu228
Case Is = 229
    Call actu229
Case Is = 230
    Call actu230
 Case Is = 999
   Call actu999
   
  End Select
   MsgBox ("Proceso Terminado")
   
  
  

End If

Call validaactu
End If

Exit Sub
err1:
MsgBox ("Error en la actualizacion, salga de todas las terminales y vuela intentarlo con la opcion1 y luego con la opcion2")
End
End Sub

Sub actu1()
h = MsgBox("Cambia codigo de articulo 1 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
 cp = InputBox$("Ingrese nuevo codigo(debe estar creado en a2) para poner a los articulos cod. 1")
 If Val(cp) > 1 Then
  q = "select * from vta_03 where [num_int] < 33525 [id_producto] = 1"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  a = 0
  While Not rs.EOF
    espere.Label1 = "Espere... 1 " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_producto") = Val(cp)
    rs.Update
    rs.MoveNext
  Wend
  Set rs = Nothing

  q = "select * from a6 where [num_int] < 33525 [id_producto] = 1"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  a = 0
  While Not rs.EOF
    espere.Label1 = "Espere... 2 " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_producto") = Val(cp)
    rs.Update
    rs.MoveNext
  Wend
  Set rs = Nothing


  q = "select * from stk_01 where  [id_producto] = 1"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  a = 0
  While Not rs.EOF
    espere.Label1 = "Espere... 3 " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_producto") = Val(cp)
    rs.Update
    rs.MoveNext
  Wend
  Set rs = Nothing


 End If
Unload espere
End If

End Sub

Sub actu90()
espere.Show
espere.Refresh
c = 0
Set rs = New ADODB.Recordset
q = "select * from a5 "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    c = c + 1
    espere.Label1 = "Espere... Procesando registro " & c
    espere.Label1.Refresh
    rs("fecha_vto") = rs("fecha")
    rs.Update
    rs.MoveNext
  Wend
  Set rs = Nothing
  Unload espere
End Sub
Sub actu91()
espere.Show
espere.Refresh
c = 0
Set rs = New ADODB.Recordset
q = "select * from a2 "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    c = c + 1
    espere.Label1 = "Espere... Procesando registro " & c
    espere.Label1.Refresh
    rs("emite_etiqueta") = "N"
    rs("texto_central") = " "
    
    rs.Update
    rs.MoveNext
  Wend
  Set rs = Nothing
  Unload espere

End Sub

Sub actu92()
espere.Show
espere.Refresh
c = 0
Set rs = New ADODB.Recordset
q = "select * from vta_02, vta_01 where vta_02.[id_cliente] = vta_01.[id_cliente] "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
   If rs("vta_02.id_cliente") > 1 Then
      cli = rs("denominacion")
      d = rs("direccion")
      CUIT = rs("cuit")
      l = rs("localidad")
      ti = rs("id_tipoiva")
    Else
      cli = "-"
      d = "-"
      CUIT = "0"
      l = "-"
      ti = 1
    End If
    c = c + 1
    espere.Label1 = "Espere... Procesando registro " & c
    espere.Label1.Refresh
    rs("cliente02") = Left$(cli, 50)
    rs("direccion02") = Left$(d, 50)
    rs("cuit02") = CUIT
    rs("localidad02") = Left$(l, 50)
    rs("id_tipo_iva02") = ti
    rs.Update
    rs.MoveNext
  Wend
  Set rs = Nothing
  Unload espere

End Sub
Sub actu93()
espere.Show
espere.Refresh
c = 0
Set rs = New ADODB.Recordset
q = "select * from A2 "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    rs("vigente") = 1
    rs("texto_central") = " "
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere

End Sub

Sub actu94()
espere.Show
espere.Refresh
c = 0
Set rs = New ADODB.Recordset
q = "select * from c_01 "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
    rs("tipo_cuentacaja") = "A"
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere

End Sub

Sub actu95()
espere.Show
espere.Refresh
c = 0
Set rs = New ADODB.Recordset
q = "select * from a6 "
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
  If rs("descuento") > 0 Then
      d1 = 100 - rs("descuento")
      psd = (rs("pu") * 100) / d1
      rs("pusindto") = psd
  Else
      rs("pusindto") = rs("pu")
  End If
  rs.Update
  rs.MoveNext
Wend
Set rs = Nothing
Unload espere

End Sub



Sub actu98()
h = MsgBox("Actualizacion 98 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a1"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_provincia") = 2
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Unload espere
End If
End Sub
Sub actu101()
h = MsgBox("Actualizacion 101 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from g1"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("imprime_pie_reportes") = True
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from a5"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("estado_pago") = "N" Then
      rs("saldo_impago") = rs("total")
    Else
      rs("saldo_impago") = 0
    End If
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing


Unload espere
End If

End Sub

Sub actu102()
h = MsgBox("Actualizacion 101 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a5"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("estado_pago") = "N" Then
      If rs("moneda") = "P" Then
        rs("saldo_impago") = rs("total")
      Else
        rs("saldo_impago") = rs("total_d")
      End If
    Else
      rs("saldo_impago") = 0
    End If
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If


End Sub


Sub actu103()
h = MsgBox("Actualizacion 103 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("chofer02") = " "
    rs("dominio02") = " "
    rs("dominio_acoplado02") = " "
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If
End Sub
Sub actu108()
h = MsgBox("Actualizacion 108 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_01"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("id_tipoiva") <> 3 And rs("id_tipoiva") <> 8 Then
       'lleva cuit
       If Len(rs("cuit")) = 13 Then
          c = Mid$(rs("cuit"), 1, 2) & Mid$(rs("cuit"), 4, 8) & Mid$(rs("cuit"), 13, 1)
          rs("cuit") = c
          rs.Update
       End If
     End If
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu134()
h = MsgBox("Actualizacion 134 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a1"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("cod_tipoiva") <> 3 And rs("cod_tipoiva") <> 8 Then
       'lleva cuit
       If Len(rs("cuit")) = 13 Then
          c = Mid$(rs("cuit"), 1, 2) & Mid$(rs("cuit"), 4, 8) & Mid$(rs("cuit"), 13, 1)
          rs("cuit") = c
          rs.Update
       End If
     End If
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu136()
h = MsgBox("Actualizacion 136 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from cyb_02 where [num_int_op] > 0"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    Set rs2 = New ADODB.Recordset
    q = "select * from cyb_04 where [modulo] = 'C' and [num_mov_int] = " & rs("num_int_op")
    rs2.Open q, cn1
    If Not rs2.EOF And Not rs2.BOF Then
      nib = rs2("num_mov_banco")
    Else
      nib = 0
    End If
    rs("num_mov_banco") = nib
    rs.Update
    rs.MoveNext
    Set rs2 = Nothing
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu138()
h = MsgBox("Actualizacion 138 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a5"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("id_tipocomp") = 40 Then
       rs("zona") = 2
    Else
       rs("zona") = 1
    End If
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu140()
h = MsgBox("Actualizacion 140 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a5"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    If IsNull(rs("fecha_vto")) Then
       rs("fecha_vto") = rs("fecha")
       rs.Update
       a = a + 1
    End If
    
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu141()
h = MsgBox("Actualizacion 141 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a5, a1 where a5.[id_proveedor] = a1.[id_proveedor] and a5.[id_proveedor] > 1"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("cuit05") = Val(rs("cuit"))
    rs("proveedor05") = Left$(rs("denominacion"), 50)
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub


Sub actu142()
h = MsgBox("Actualizacion 142 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a21"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("cant_lineas") = 1
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub


Sub actu145()
h = MsgBox("Actualizacion 145 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_06"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("imprime_desc_extra") = "N"
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from g2"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("imprime_desc_extra") = "N"
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing




Unload espere
End If

End Sub


Sub actu148()
h = MsgBox("Actualizacion 148 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_03"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("tasaib") = 3
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from a2"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("id_tasaib") = 1
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing




Unload espere
End If

End Sub


Sub actu149()
h = MsgBox("Actualizacion 149 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from cyb_03 where [num_mov_banco_e] > 0"
rs.Open q, cn1
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    Set rs1 = New ADODB.Recordset
    q = "select * from cyb_05 where [modulo] = 'B' and [num_mov_int] = " & rs("num_mov_banco_e") & "  and [importe] = " & rs("importe") & "  And [num_int_ch_terc] = 0 "
    rs1.Open q, cn1, adOpenDynamic, adLockOptimistic
    If Not rs1.EOF And Not rs1.BOF Then
       rs1("num_int_ch_terc") = rs("num_interno")
       rs1.Update
    End If
    Set rs1 = Nothing
    rs.MoveNext
Wend
Set rs = Nothing


Unload espere
End If

End Sub

Sub actu156()

h = MsgBox("Actualizacion 156 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
q = "alter table pro_04 add column tipo04 integer "
cn1.Execute q

q = "alter table pro_05 add column obs string(25) "
cn1.Execute q


Set rs = New ADODB.Recordset
q = "select * from pro_05 where [tipo_comprobante] <> 1"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    Select Case rs("tipo_comprobante")
    Case Is = 2
    rs("tipo_comprobante") = 65
    Case Is = 3
    rs("tipo_comprobante") = 1
    End Select
    rs("obs") = " "
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing


Set rs = New ADODB.Recordset
q = "select * from pro_04"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    Set rs2 = New ADODB.Recordset
    q = "select * from pro_05 where [num_referencia] = " & rs("num_referencia") & " and [secuencia] = 1 "
    rs2.Open q, cn1
    On Error GoTo err1
    If rs2("tipo_comprobante") = 1 Then
      rs("tipo04") = 1 'pedido produccion
    Else
      rs("tipo04") = 2 'oc directa
    End If
    rs.Update
    rs.MoveNext
    Set rs2 = Nothing
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from pro_03"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
rs.AddNew
rs("id_tipocomp") = 2
rs("descripcion") = "Minuta Interna"
rs("abreviatura") = "minuta"
rs("copias") = 1
rs("ult_numero") = 0
rs.Update
Set rs = Nothing


Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub

Sub actu158()

h = MsgBox("Actualizacion 158 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
q = "alter table g1 add column estilo_rc integer "
cn1.Execute q

Set rs = New ADODB.Recordset
q = "select * from g1"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("estilo_rc") = 0
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing



Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub

Sub actu159()

h = MsgBox("Actualizacion 159 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
q = "alter table vta_01 add column [direccion_local] string(80) "
cn1.Execute q

Set rs = New ADODB.Recordset
q = "select * from vta_01"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("direccion_local") = rs("direccion")
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing


Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub


Sub actu160()

h = MsgBox("Actualizacion 160 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
cn1.BeginTrans
q = "alter table a2 add column [id_prod_prov] string(10)  "
cn1.Execute q
cn1.CommitTrans

cn1.BeginTrans
q = "alter table a2 add column [dto_compra2] single  "
cn1.Execute q
cn1.CommitTrans


Set rs = New ADODB.Recordset
q = "select * from a2"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    rs("id_prod_prov") = 0
    rs("dto_compra2") = 0
    
    rs.Update
    rs.MoveNext
    a = a + 1
Wend
Set rs = Nothing


Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub


Sub actu162()

h = MsgBox("Actualizacion 162 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
cn1.BeginTrans
q = "alter table g0 add column [texto_resumen1] string(100) NOT NULL DEFAULT '*', [texto_resumen2] string(100)NOT NULL DEFAULT '*', [imprime_texto_resumen] int NOT NULL DEFAULT 0,  [actualizacion] int NOT NULL DEFAULT 162 "
cn1.Execute q
cn1.CommitTrans

q = "update g0 set  [texto_resumen1]='*' , [texto_resumen2]='*', [imprime_texto_resumen]=0, [actualizacion]=162"
      q = q & " where [sucursal]=0 "
      cn1.BeginTrans
      cn1.Execute q
      cn1.CommitTrans

Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub


Sub actu163()

h = MsgBox("Actualizacion 163 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
cn1.BeginTrans
q = "alter table g0 add column [precio_remito_factura] int NOT NULL DEFAULT 0 "
cn1.Execute q
cn1.CommitTrans

q = "update g0 set  [actualizacion]=163, [precio_remito_factura]=0"
q = q & " where [sucursal]=0 "
cn1.BeginTrans
  cn1.Execute q
cn1.CommitTrans

Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub

Sub actu164()

h = MsgBox("Actualizacion 164 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a5, a1 where [id_tipocomp] = 65 and a5.[id_proveedor] = a1.[id_proveedor]"
rs.Open q, cn1, adOpenDynamic, adLockOptimistic
While Not rs.EOF
  rs("proveedor05") = Left$(rs("denominacion"), 50)
  rs("cuit05") = Val(rs("cuit"))
  rs.Update
  rs.MoveNext
Wend
Set rs = Nothing

q = "update g0 set  [actualizacion]=164"
q = q & " where [sucursal]=0 "
cn1.BeginTrans
  cn1.Execute q
cn1.CommitTrans


Unload espere
End If
Exit Sub
err1:
Resume Next
End Sub

Sub actu165()

h = MsgBox("Actualizacion 165 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
f = InputBox$("Fecha corte para pasar a facturadas O.C", , "31/08/2013")
If IsDate(f) Then
  espere.Show
  espere.Refresh
  Set rs = New ADODB.Recordset
  q = "select * from a5 where [id_tipocomp] = 65 and datevalue([fecha]) <=  DateValue('" & f & "')"
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
   rs("estado") = "F"
   rs.Update
   rs.MoveNext
  Wend
  Set rs = Nothing

  q = "update g0 set  [actualizacion]=165"
  q = q & " where [sucursal]=0 "
  cn1.BeginTrans
  cn1.Execute q
  cn1.CommitTrans
  
  Unload espere
Else
  If f <> "" Then
    MsgBox ("Formatro incorrecto de fecha")
  End If
  
End If

End If
Exit Sub
err1:
Resume Next
End Sub


Sub actu166()

h = MsgBox("Actualizacion 166 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
  cn1.BeginTrans
  q = "alter table g0 add column [tipo_redondeo] int NOT NULL DEFAULT 0 "
  cn1.Execute q
  cn1.CommitTrans

  q = "update g0 set  [actualizacion]=166, [tipo_redondeo]=0"
  q = q & " where [sucursal]=0 "
  cn1.BeginTrans
  cn1.Execute q
  cn1.CommitTrans
  
  Unload espere
Else
  If f <> "" Then
    MsgBox ("Formatro incorrecto de fecha")
  End If
  
End If


Exit Sub
err1:
Resume Next
End Sub


Sub actu171()

h = MsgBox("Actualizacion 171 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
  cn1.BeginTrans
  q = "alter table a5 add column [id_cliente] long NOT NULL DEFAULT 0 "
  cn1.Execute q
 
  q = "alter table stk_01 add column [id_cliente] long NOT NULL DEFAULT 0 "
  cn1.Execute q
  

  q = "update g0 set  [actualizacion]=171"
  q = q & " where [sucursal]=0 "
 
  cn1.Execute q
  cn1.CommitTrans
  
  MsgBox ("El siguiente proceso va demorar. Espere por favor")
  
  espere.Show
  espere.Label1 = "Actualizando Comprobantes..."
  espere.Refresh
  q = "select * from a5 "
 
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  On Error GoTo err1
  While Not rs.EOF
     rs("id_cliente") = 0
     rs.MoveNext
  Wend
  Set rs = Nothing
  
  espere.Label1 = "Actualizando Stock..."
  espere.Refresh
  q = "select * from stk_01 "
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  On Error GoTo err1
  While Not rs.EOF
     rs("id_cliente") = 0
     rs.MoveNext
  Wend
  Set rs = Nothing
 Unload espere
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu172()

h = MsgBox("Actualizacion 172 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
  cn1.BeginTrans
  q = "alter table g0 add column [id_cuenta_retibba_ventas] long NOT NULL DEFAULT 110101"
  cn1.Execute q
  q = "alter table g0 add column [id_cuenta_retiva_ventas] long NOT NULL DEFAULT 110101"
  cn1.Execute q
  q = "alter table g0 add column [id_cuenta_retgan_ventas] long NOT NULL DEFAULT 110101"
  cn1.Execute q
  q = "alter table g0 add column [id_cuenta_retsuss_ventas] long NOT NULL DEFAULT 110101"
  cn1.Execute q

 
  cn1.CommitTrans
  
  
  cn1.BeginTrans
  
  q = "update g0 set  [actualizacion]=172, [id_cuenta_retibba_ventas]=110101, [id_cuenta_retiva_ventas]=110101, [id_cuenta_retgan_ventas]=110101, [id_cuenta_retsuss_ventas]=110101  "
  q = q & " where [sucursal]=0 "
 
  cn1.Execute q
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu173()

h = MsgBox("Actualizacion 173 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
  cn1.BeginTrans
  q = "alter table g2 add column [cod_afip_A] long NOT NULL DEFAULT 1"
  cn1.Execute q
  q = "alter table g2 add column [cod_afip_B] long NOT NULL DEFAULT 6"
  cn1.Execute q
  q = "alter table g2 add column [cod_afip_C] long NOT NULL DEFAULT 11"
  cn1.Execute q
  
  cn1.CommitTrans
  
  
  cn1.BeginTrans
  
  q = "update g0 set  [actualizacion]=173  "
  q = q & " where [sucursal]=0 "
  cn1.Execute q
  
  q = "update g2 set  [cod_afip_A]=1, [cod_afip_B]=6, [cod_afip_C]=11  "
  q = q & " where [id_tipo_comp]=1 "
  cn1.Execute q
  
  q = "update g2 set  [cod_afip_A]=2, [cod_afip_B]=7, [cod_afip_C]=12  "
  q = q & " where [id_tipo_comp]=20 "
  cn1.Execute q
  
  q = "update g2 set  [cod_afip_A]=3, [cod_afip_B]=8, [cod_afip_C]=13  "
  q = q & " where [id_tipo_comp]=30 "
  
  cn1.Execute q
  
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu174()

h = MsgBox("Actualizacion 174 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  
  cn1.BeginTrans
  
  q = "update g0 set  [actualizacion]=174  "
  q = q & " where [sucursal]=0 "
  cn1.Execute q
  
  q = "update i_01 set  [retencion-minima]=90  "
  q = q & " where [id_impuesto]=217 "
  cn1.Execute q
  
  q = "update i_02 set  [importe_noretenido]=7500  "
  q = q & " where [id_impuesto]=217 and [id_concepto]=25 "
  cn1.Execute q
  
  q = "update i_02 set  [importe_noretenido]=5000  "
  q = q & " where [id_impuesto]=217 and [id_concepto]=27 "
  cn1.Execute q
  
  q = "update i_02 set  [importe_noretenido]=100000  "
  q = q & " where [id_impuesto]=217 and [id_concepto]=78 "
  cn1.Execute q
  
  q = "update i_02 set  [importe_noretenido]=30000  "
  q = q & " where [id_impuesto]=217 and [id_concepto]=94 "
  cn1.Execute q
  
  q = "update i_02 set  [importe_noretenido]=30000  "
  q = q & " where [id_impuesto]=217 and [id_concepto]=95 "
  cn1.Execute q
  
  q = "update i_02 set  [importe_noretenido]=30000  "
  q = q & " where [id_impuesto]=217 and [id_concepto]=116 "
  cn1.Execute q
  
  
  cn1.Execute q
  
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu176()

h = MsgBox("Actualizacion 176 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  
  cn1.BeginTrans
  q = "alter table g0 add column [nc_en_recibo] string(1) "
  cn1.Execute q
  
  
  q = "update g0 set  [actualizacion]=176 , [nc_en_recibo]= 'N' "
  q = q & " where [sucursal]=0 "
  cn1.Execute q
  
  
  
  
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu178()

h = MsgBox("Actualizacion 178 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
  q = "alter table g0 add column [muestra_saldo_fact_venta] string(1) "
  cn1.Execute q
  cn1.CommitTrans
 
  
   cn1.BeginTrans
  
  q = "update g0 set  [actualizacion]=178, [muestra_saldo_fact_venta]='N'"
  q = q & " where [sucursal]=0 "
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub



Sub actu179()

h = MsgBox("Actualizacion 179 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
  q = "alter table g1 add column [imprime_cabecera_reportes] string(1) "
  cn1.Execute q
  cn1.CommitTrans
 
  
   cn1.BeginTrans
  
  q = "update g1 set  [imprime_cabecera_reportes]='S'"
  'q = q & " where [id_usuario]>0 "
  
   cn1.Execute q
   
  q = "update g0 set  [actualizacion]=179"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu180()

h = MsgBox("Actualizacion 180 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  MsgBox ("Esta actualizacion requiere acciones manuales. Importe tabla A23")
  
  
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=180"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu181()

h = MsgBox("Actualizacion 181 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  MsgBox ("Esta actualizacion requiere acciones manuales. Borrar y volver a importar tablas I_01, I_02, I_03")
  
  
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=181"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu201()

h = MsgBox("Actualizacion 201 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  MsgBox ("Gestione 11. Factura Electonica. Ejecutar Actu e Importar Manualmente fe_01. Poner Nombre fantasia, fecha inicio actividades y numib")
  
  
  
  cn1.BeginTrans
  q = "alter table vta_02 add column [cae] string(255) DEFAULT '0' "
  cn1.Execute q
  q = "alter table vta_02 add column [cae_vence] date DEFAULT '1-1-2018' "
  cn1.Execute q
    
    
  q = "alter table g0 add column [nombre_fantasia] string(255) DEFAULT 'Nombre fantasia' "
  cn1.Execute q
    
    
  q = "alter table g0 add column [fecha_inicio_actividades] string(20) DEFAULT '01/01/2000' "
  cn1.Execute q
    
    
  q = "alter table g0 add column [numero_ingresos_brutos] string(50) DEFAULT '20202020202' "
  cn1.Execute q
  
  cn1.CommitTrans
 
  
   cn1.BeginTrans
  
  q = "update vta_02 set  [cae]='0', [cae_vence]='01-01-2018'"
  q = q & " where [num_int]>=0 "
  
  
  
  q = "update g0 set  [nombre_fantasia]='Nombre Fantasia', [fecha_inicio_actividades]='01-01-2018', [numero_ingresos_brutos]='2020202020202'"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
  
  cn1.CommitTrans
  
  
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=201"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu202()

h = MsgBox("Actualizacion 202 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
    
  cn1.BeginTrans
  q = "alter table vta_02 add column [tipo_op] integer DEFAULT 1 "
  cn1.Execute q
    
  
  cn1.CommitTrans
 
  
   cn1.BeginTrans
  
  q = "update vta_02 set  [tipo_op]=2"
  q = q & " where [num_int]>=0 "
  
   cn1.Execute q
  
  
  
  
  cn1.CommitTrans
  
  
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=202"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu203()

h = MsgBox("Actualizacion 203 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
    
  
  
   cn1.BeginTrans
  
  q = "update vta_02 set  [cae]='u2', [cae_vence]='01/01/2018'"
  q = q & " where [num_int]>=0 "
  
   cn1.Execute q
  
  
  
  
  cn1.CommitTrans
  
  
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=203"
  q = q & " where [sucursal]=0 "
  
  
  cn1.Execute q
    
  cn1.CommitTrans
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu204()

h = MsgBox("Actualizacion 204 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table g0 add column [id_sistema] integer DEFAULT 0 "
  cn1.Execute q
  cn1.CommitTrans
    
   
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=204"
    q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  MsgBox ("Colocar en la tabla g0 el id_sistema unico para cada cliente. Si no tiene acceso web dejar 0")
 
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu205()

h = MsgBox("Actualizacion 205 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table g0 add column [cbu] string(50) DEFAULT ' ' "
  cn1.Execute q
  cn1.CommitTrans
    
   
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=205"
    q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  MsgBox ("Colocar en la tabla g0 el CBU para cobros por transferencias del cliente(saldrá en la factura electrónica)")
 
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu206()

h = MsgBox("Actualizacion 206 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table g0 add column [ult_sinc_nube] string(30) DEFAULT '01/01/2001' "
  cn1.Execute q
  cn1.CommitTrans
    
   
   cn1.BeginTrans
  
    q = "update g0 set  [actualizacion]=206, [ult_sinc_nube]='01/01/2001'"
    q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu207()

h = MsgBox("Actualizacion 207 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  MsgBox ("IMPORTANTE: Reemplazar tabla G4 - Despues de la actualizacion complete los campos path_facturas y dias_pago_cc en tabla fe_01")

  MsgBox ("IMPORTANTE: Agreagar los archivos logo.png afip.png y factura.csv en la carpeta c:\5a04")

  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table fe_01 add column [path_facturas] string(100) DEFAULT 'c:\pyafipws\facturas', [dias_pago_cc] integer DEFAULT 10 "
  cn1.Execute q
  
   
       
  q = "update g0 set  [actualizacion]=207"
  q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu208()

h = MsgBox("Actualizacion 208 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  MsgBox ("Despues de actualizar ingresar datos de servidor de correo envio")
  
  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table fe_01 add column [servidor_email] string(150), [usuario_email] string(150), [pass_email] string(50), [email_remite] string(150) "
  cn1.Execute q
  
   
       
  q = "update g0 set  [actualizacion]=208"
  q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu210()

h = MsgBox("Actualizacion 210 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  
  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table emp_02 add column [id] autoincrement constraint primarykey PRIMARY KEY "
  cn1.Execute q
  'ALTER TABLE nombredetutabla ADD COLUMN nombredetucampo AUTOINCREMENT CONSTRAINT PrimaryKey PRIMARY KEY   AGREGAR CAMPO AUTONUMERCO COMO CLAVE PRIMARIA
       
  
  q = "alter table emp_02 add index I_NUMMOV (num_mov_int)"
  q = "CREATE INDEX i_mov ON emp_02(num_mov_int)"
  cn1.Execute q
  'alter table libros  add index i_editorial (editorial);     AGREGAR INDICE
  
  q = "update g0 set  [actualizacion]=210"
  q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu211()

h = MsgBox("Actualizacion 211 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  
  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
  q = "alter table g1 add column [punto_venta_inicio] integer "
  cn1.Execute q
       
  
  
  q = "update g0 set  [actualizacion]=211"
  q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
 
 cn1.BeginTrans
  q = "update g1 set [punto_venta_inicio] = 1 "
  cn1.Execute q
       
    
  cn1.CommitTrans
 
 
 
 
 
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu212()

h = MsgBox("Actualizacion 212 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  
  
    
  MsgBox ("ACTU 212. Agregue en gen.ini de cada terminal una linea con el punto de venta manual de inicio de cada equipo")
 
 
  cn1.BeginTrans
 
  q = "update g0 set  [actualizacion]=212"
  q = q & " where [sucursal]=0 "
  
  cn1.Execute q
    
  cn1.CommitTrans
 
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu213()
h = MsgBox("Actualizacion 213 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  MsgBox ("Agregar manualmente los comprobantes bancarios 700 y 701 a la tabla G2 comprobantes de compra")
  espere.Show
  espere.Refresh
    
  cn1.BeginTrans
     q = "alter table c_12 alter column [descripcion] string(80)  "
  cn1.Execute q
    
            
                 
   
   
 
       
  q = "update g0 set  [actualizacion]=213"
  q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
  
  
 
   MsgBox ("Operación cerrada")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu215()
h = MsgBox("Actualizacion 215 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  MsgBox ("Reemplazar tabla Fsc_001. Si el cliente tiene controlador fiscal verificar prametros impresora")
  espere.Show
  espere.Refresh
    
    
   cn1.BeginTrans
  q = "update g0 set  [actualizacion]=215"
  q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
  
  
 
   MsgBox ("Operación cerrada")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu216()
'mofidica el ancho del campo descripcion a 80 en tabla c_12
'modifica descripcion en asientos
h = MsgBox("Actualizacion 216 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  espere.Show
  espere.Refresh
    
    
   
   cn1.BeginTrans
     q = "alter table c_12 alter column [descripcion] string(80)  "
     cn1.Execute q
   
  q = "update g0 set  [actualizacion]=216"
  q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
  
  
 
   MsgBox ("Operación cerrada")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu217()
'agrega alias (bancaria) a G0 para emision de FCE
h = MsgBox("Actualizacion 217 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  
  espere.Show
  espere.Refresh
    
    
   
   cn1.BeginTrans
     q = "alter table g0 add column [alias] string(150)  "
     cn1.Execute q
   
  q = "update g0 set  [actualizacion]=217, alias='" & "Alias Banco" & "'"
  q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
  
   
    
   MsgBox ("Operación Completa. Si tiene sistema de Facturacion electronica es obligatorio informar CBU y Alias cuetnta bancaria")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu218()
'agrega cod-fiscal2 en g3 y codigo_driver_fiscal en cyb_01
h = MsgBox("Actualizacion 218 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  h = MsgBox("Esta actualizacion corrige una actualizacion para el uso de controladores fiscales nuevo protocolo. Puede ser que algun sistema ya lo tenga.")
  espere.Show
  espere.Refresh
    
    
   
   cn1.BeginTrans
     q = "alter table g3 add column [cod_fiscal2] int  "
     cn1.Execute q
   
     q = "alter table cyb_01 add column [codigo_driver_fiscal] int  "
     cn1.Execute q
   
   
    q = "update g0 set  [actualizacion]=218"
    q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
  
   
    
   MsgBox ("Operación Completa. Si usa CF verificar valores de campos cod_fiscal2 en g3 y codigo_driver_fiscal en cyb_01")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub

Sub actu219()
'agrega tabla vta_016 percepciones de venta
h = MsgBox("Actualizacion 219 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  h = MsgBox("Importar tabla vta_016 (percepciones de venta")
  espere.Show
  espere.Refresh
    
    
   cn1.BeginTrans
   
   
    q = "update g0 set  [actualizacion]=219"
    q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
  
   
    
   MsgBox ("Operación Completa. Borrar datos de tabla vta_016 importada")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu220()
'agrega datos a productos para calcular costos de estructura de productos
h = MsgBox("Actualizacion 220 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
    
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     q = "alter table a2 add column [num_int_ult_compra] double, [dolar_ult_compra] double  "
     cn1.Execute q
   
    q = "update g0 set  [actualizacion]=220"
    q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
    
    
   MsgBox ("Se va a proceder a realizar una actualizacion de datos")
   
  q = "select * from a2"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
      If IsNull(rs("ultima_compra")) Then
        lc = "X"
      Else
           lc = Mid$(rs("ultima_compra"), 1, 1)
      End If
      If lc = "A" Then
          tc = 1
          sc = Val(Mid$(rs("ultima_compra"), 2, 4))
          nc = Val(Mid$(rs("ultima_compra"), 7, 8))
          pc = rs("id_proveedor_ult_compra")
          q = "select * from a5 where sucursal=" & sc & " and num_comprobante=" & nc & " and letra= '" & lc & "' and id_tipocomp=" & tc & " and id_proveedor=" & pc
          Set rs2 = New ADODB.Recordset
          rs2.Open q, cn1
          If Not rs2.EOF And Not rs2.BOF Then
            niuc = rs2("num_int")
            cduc = rs2("cotiz_dolar")
          Else
            niuc = 0
            cduc = 0
          End If
          Set rs2 = Nothing
      Else
         niuc = 0
         cduc = 1
      End If
  
      rs("num_int_ult_compra") = niuc
      rs("dolar_ult_compra") = cduc
      rs.Update
   
      rs.MoveNext
  Wend
  Set rs = Nothing
  MsgBox ("proceso terminado")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub


Sub actu222()
'agrega datos a lista de precio talle, color, etc
h = MsgBox("Actualizacion 222 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
 
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     q = "alter table a2 add column [talle] text(10), [color] text(25), [medida] text(35)  "
     cn1.Execute q
   
     q = "alter table g0 add column [cuotas_sininteres] int, [interes_cuota] double  "
     cn1.Execute q
     
     q = "alter table vta_02 add column [numint_asociado] long "
     cn1.Execute q
     
   
   
    q = "update g0 set  [actualizacion]=222, [cuotas_sininteres]=2, [interes_cuota] = 7"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    q = "update vta_02 set [numint_asociado]=0"
    cn1.Execute q
    
  cn1.CommitTrans
    
    
   MsgBox ("Se va a proceder a realizar una actualizacion de datos")
   
  q = "select * from a2"
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
      rs("talle") = "*"
      rs("color") = "*"
      rs("medida") = "*"
      rs.Update
   
      rs.MoveNext
  Wend
  Set rs = Nothing
  
  
  
  q = "select * from vta_02, vta_010 where [estado_pago] = 'P' and [num_int] = [num_int_comp] and [saldo_comprobante]=0 "
  Set rs = New ADODB.Recordset
  
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  While Not rs.EOF
      Set rs2 = New ADODB.Recordset
      q = "select fecha from vta_02 where [num_int] = " & rs("num_int_rbo")
      rs2.Open q, cn1
      If Not rs2.EOF And Not rs2.BOF Then
         f = rs2("fecha")
      Else
         f = "01-01-2000"
      End If
      Set rs2 = Nothing
      
      rs("fecha_pago") = f
      rs.Update
   
      rs.MoveNext
  Wend
  Set rs = Nothing
  
  
  
  MsgBox ("proceso terminado")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu223()
'agrega datos a lista de precio talle, color, etc
h = MsgBox("Actualizacion 223 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
 
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table g0 add column [minimo_informar_cons_final] double  "
     cn1.Execute q
     
    q = "update g0 set  [actualizacion]=223, [minimo_informar_cons_final]=43010"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("proceso terminado")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu224()
'corrige ajustes de stock
h = MsgBox("Actualizacion 224 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
 
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table g0 add column [ult_num_ajuste_stock] long  "
     cn1.Execute q
     
    q = "update g0 set  [actualizacion]=224, [ult_num_ajuste_stock]=100"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado. Verificar que el campo numint en stk02 esté como autonumerico. Agregar xml.txt en crapeta c:\5a04\log de cada terminal de factura electronica")
   
  MsgBox ("Agregar comprobantes de venta 36 (Nota de Venta en cuotas), 251 (Cuotas), y 401(Pagare)")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub



Sub actu225()
'agrega percepciones de venta articulos limpiaeza
h = MsgBox("Actualizacion 225 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
     
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table a2 add column [percibe_5329] text(1)  "
     cn1.Execute q
     
   
     q = "alter table I_01 add column [id_otrostributos] int, [tasa_i1] double"
     cn1.Execute q
     
     
   cn1.CommitTrans
   
   cn1.BeginTrans
     
    q = "update a2 set [percibe_5329]='N'"
    cn1.Execute q
    
    q = "update i_01 set [id_otrostributos]=99, tasa_i1=0, [id_cuenta_i1] = 110302"
    cn1.Execute q
    
    
    q = "update g0 set  [actualizacion]=225"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado. Verificar articulos para aquellos que deban percibir Iva RG5329 ")
   
  MsgBox ("Agregar en I_01 registro 5329 para Percepcion 5329(3%) y  5328 Precepcion 5329(1.5%) ")
  
  MsgBox ("Modifique I_01 id_otrostirbutos segun excel otrostributos(6 percep iva - 7 perc ibba")
   
  MsgBox ("Revisar cuenta contable en I1")
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu227()
'agrega separador de miles
h = MsgBox("Actualizacion 227(Separador MIles) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
     
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table g1 add column [usa_separador_miles] text(1)  "
     cn1.Execute q
     
   cn1.CommitTrans
   
   
   cn1.BeginTrans
     
    q = "update g1 set [usa_separador_miles]='S'"
    cn1.Execute q
    
    
    
    q = "update g0 set  [actualizacion]=227"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado. Configure separador de miles por usuario ")
   
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub




Sub actu226()
'resolucion pantalla
h = MsgBox("Actualizacion 226 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
    
  
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
    
    q = "update g0 set  [actualizacion]=226"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado. Todas las terminales tendrán que tener resolución mínima de pantalla HD 1280x720 ")
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub



Sub actu228()
'percepciones rg5329 parte 2
h = MsgBox("Actualizacion 228(Percepciones RG5329 parte 2) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
     
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table vta_016 add column [base_imponible] Double, [alicuota] Double"
     cn1.Execute q
     
   cn1.CommitTrans
   
   
   cn1.BeginTrans
     
    q = "update vta_016 set [base_imponible]=0, [alicuota] = 1"
    cn1.Execute q
    
    
    
    q = "update g0 set  [actualizacion]=228"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado.  Redimencionar campo ALIAS en G0 a Text (20)")
   
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu229()
'Soluciona problema con percepciones de venta. Pasa todos los codigos de percepcion a la tabla I_01
h = MsgBox("Actualizacion 229(Percepciones Venta) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  espere.Show
  espere.Refresh
   cn1.BeginTrans
    q = "update g0 set  [actualizacion]=229"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado.  Agregar en tabla I_01 los codigos de percepcion que aparecen en la tabla A12(con el mismo codigo)")
   
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub



Sub actu230()
'Soluciona problema con percepciones de venta. Pasa todos los codigos de percepcion a la tabla I_01
h = MsgBox("Actualizacion 230(Agrega codigos Percepciones Venta) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  espere.Show
  espere.Refresh
   cn1.BeginTrans
   
   
    q = "alter table i_01 add column [tipo_i1] Text(1), [impuesto_i1] Text(5) "
     cn1.Execute q
   
   
    q = "update g0 set  [actualizacion]=230"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  MsgBox ("Proceso terminado.  Modifique tabla I_01 Tipo: P Percepciones, R Retenciones - Impuesto: IVA-IBBA-GAN")
   
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu221()
'agrega posibilidad de factura C
h = MsgBox("Actualizacion 221 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
    
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     q = "alter table g0 add column [id_tipo_iva] int  "
     cn1.Execute q
   
     q = "alter table vta_06 add column [cod_afip_c] int  "
     cn1.Execute q
   
   
    q = "update vta_06 set  [cod_afip_c]=0"
    
    
    q = "update g0 set  [actualizacion]=221, [id_tipo_iva] = 1"
    q = q & " where [sucursal]=0 "
  
   cn1.Execute q
    
  cn1.CommitTrans
    
  
  MsgBox ("Agregar comp venta 35(nota venta), 36(nota venta con cuotas) 401(pagare), 251(cuota)")
   
  MsgBox ("Poner codigos afip para comprobantes tipo C ")
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next
End Sub





Sub actu167()

h = MsgBox("Actualizacion 167 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
  q = "select * from vta_06 order by [sucursal], [id_tipocomp]"
  b = 0
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
  On Error GoTo err1
  While Not rs.EOF
     If b <> rs("sucursal") Then
       b = rs("sucursal")
       rs.AddNew
       rs("sucursal") = b
       rs("id_tipocomp") = 150
       rs("Descripcion") = "Orden de Empaque"
       rs("abreviatura") = "Empq"
       rs("ult_num_A") = 0
       rs("ult_num_B") = 0
       rs("ult_num_C") = 0
       rs("stock") = "N"
       rs("ctacte") = "N"
       rs("Iva") = "N"
       rs("tipo_impresora") = "G"
       rs("cant_lineas") = 25
       rs("cant_copias_A") = 1
       rs("cant_copias_B") = 1
       rs("cant_copias_C") = 1
       rs("moneda") = "A"
       rs("venta") = "N"
       rs("contabilidad") = "N"
       rs("propio") = "S"
       rs("cant_copias_e") = 1
       rs("ult_num_e") = 0
       rs("ib") = "N"
       rs("formato") = "1"
       rs("cod_afip_a") = 0
       rs("cod_afip_b") = 0
       rs("cod_afip_e") = 0
       rs("imprime_desc_extra") = "S"
       
     End If
     rs.MoveNext
 Wend

 Call actuversion(167)
 Unload espere

  

End If


Exit Sub
err1:
MsgBox ("Error")
Call actuversion(167)
 Unload espere
End Sub


Sub actu168()

h = MsgBox("Actualizacion 168 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then

  espere.Show
  espere.Refresh
  q = "select * from vta_06 order by [sucursal], [id_tipocomp]"
  b = 0
  Set rs = New ADODB.Recordset
  rs.Open q, cn1, adOpenDynamic, adLockOptimistic
 
  While Not rs.EOF
     If b <> rs("sucursal") Then
       b = rs("sucursal")
       rs.AddNew
       rs("sucursal") = b
       rs("id_tipocomp") = 25
       rs("Descripcion") = "Pro Forma Venta"
       rs("abreviatura") = "ProF"
       rs("ult_num_A") = 0
       rs("ult_num_B") = 0
       rs("ult_num_C") = 0
       rs("stock") = "N"
       rs("ctacte") = "D"
       rs("Iva") = "N"
       rs("tipo_impresora") = "G"
       rs("cant_lineas") = 25
       rs("cant_copias_A") = 1
       rs("cant_copias_B") = 1
       rs("cant_copias_C") = 1
       rs("moneda") = "A"
       rs("venta") = "N"
       rs("contabilidad") = "N"
       rs("propio") = "S"
       rs("cant_copias_e") = 1
       rs("ult_num_e") = 0
       rs("ib") = "N"
       rs("formato") = "1"
       rs("cod_afip_a") = 0
       rs("cod_afip_b") = 0
       rs("cod_afip_e") = 0
       rs("imprime_desc_extra") = "S"
       
       
       
     End If
     rs.MoveNext
 Wend

 Call actuversion(168)
 Unload espere

  

End If

End Sub
Sub actu169()
h = MsgBox("Actualizacion 169 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  MsgBox ("ATENCION!!! La version 169  tiene modificaciones manuales")
  MsgBox ("MODIFICACION: Copiar nueva Base de Ingresos Brutos PIB")
  
 Call actuversion(169)

  

End If


End Sub

Sub actu170()
h = MsgBox("Actualizacion 170 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  MsgBox ("ATENCION!!! La version 170  tiene modificaciones manuales")
  MsgBox ("MODIFICACION: Cambiar campo id_prod_prov en A2 (texto 20) ")
  
 Call actuversion(170)

  

End If


End Sub
Sub actuversion(v As Integer)
  q = "update g0 set  [actualizacion]=" & v
  q = q & " where [sucursal]=0 "
  cn1.BeginTrans
  cn1.Execute q
  cn1.CommitTrans
End Sub

Sub actu1482()
h = MsgBox("Actualizacion 1482. ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_03"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    If rs("id_producto") > 1 Then
      Set rs1 = New ADODB.Recordset
      q = "select * from a2, g12 where [id_producto] = " & rs("id_producto") & " and a2.[id_tasaib] = g12.[id_tasaib]"
      rs1.Open q, cn1
      If Not rs1.EOF And Not rs1.BOF Then
        rs("tasaib") = rs1("tasaib")
      Else
        rs("tasaib") = 3
      End If
      Set rs1 = Nothing
    
    Else
        rs("tasaib") = 3
    End If
    rs.Update
    a = a + 1
    
    rs.MoveNext
Wend
Set rs = Nothing


Unload espere
End If

End Sub


Sub actu109()
h = MsgBox("Actualizacion 109 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a6"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("exportacion") = 0
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If


End Sub
Sub actu111()
h = MsgBox("Actualizacion 111 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_camion02") = 1
    rs("dni_chofer02") = 0
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub
Sub actu114()
h = MsgBox("Actualizacion 114 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a2"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("reg_faltante") = 0
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If
End Sub

Sub actu116()
h = MsgBox("Actualizacion 116 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a2"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("tipo_carga_tique") = "M"
    rs("abreviatura") = Mid$(rs("descripcion"), 1, 6)
    
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from vta_02"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("num_z") = 1
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing




Unload espere
End If

End Sub
Sub actu119()
h = MsgBox("Actualizacion Critica 119 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02, vta_06 where vta_02.[id_tipocomp] = vta_06.[id_tipocomp]  and vta_06.[sucursal] = [sucursal_ingreso]"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("contado") = "S" Then
      rs("cta_cte") = rs("ctacte")
      rs.Update
    End If
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If
End Sub

Sub actu123()
h = MsgBox("Actualizacion Critica 123 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from cyb_04"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("num_mov_int_compras") = 0
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub
Sub actu124()
h = MsgBox("Actualizacion 124 (solo para los sistemas fiscales) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02 where [id_tipocomp] = 300"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("grabado") = "N"
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu125()
h = MsgBox("Actualizacion 125  . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a2 where [id_producto] > 1"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    Set rs1 = New ADODB.Recordset
    q = "select * from A5, A6 where [id_producto] = " & rs("id_producto") & " and a5.[num_int] = a6.[num_int] order by a5.[fecha] desc"
    rs1.MaxRecords = 1
    rs1.Open q, cn1
    If Not rs1.EOF And Not rs1.BOF Then
       ip = rs1("id_proveedor")
    Else
       ip = 1
    End If
    Set rs1 = Nothing
    
    rs("id_proveedor_ult_compra") = ip
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu128()
h = MsgBox("Actualizacion 128  . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a1 "
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_cuenta_a1") = 110101
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from vta_02, vta_09 where vta_02.[num_int] = vta_09.[num_int] "
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1000
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("id_cuenta09") = rs("id_cuenta")
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing





Unload espere
End If

End Sub

Sub actu129()
h = MsgBox("Actualizacion 129  . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_06 "
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    Select Case rs("id_tipocomp")
    Case Is = 1
       coda = 1
       codb = 6
       code = 19
    Case Is = 2
       coda = 2
       codb = 7
       code = 20
    Case Is = 3
       coda = 3
       codb = 8
       code = 21
    Case Is = 6
       coda = 1
       codb = 6
       code = 19
    Case Is = 101
       coda = 39
       codb = 40
       code = 99
    Case Is = 300
       coda = 99
       codb = 99
       code = 99
    Case Is = 310
       coda = 99
       codb = 83
       code = 99
    Case Else
       coda = 39
       codb = 40
       code = 99
    End Select
    rs("cod_afip_a") = coda
    rs("cod_afip_b") = codb
    rs("cod_afip_e") = code
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Set rs = New ADODB.Recordset
q = "select * from vta_02"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("id_tipo_iva02") <> 3 And rs("id_tipo_iva02") <> 8 Then
       'lleva cuit
       If Len(rs("cuit02")) = 13 Then
          c = Mid$(rs("cuit02"), 1, 2) & Mid$(rs("cuit02"), 4, 8) & Mid$(rs("cuit02"), 13, 1)
          rs("cuit02") = c
          rs.Update
       Else
        If Len(rs("cuit02")) <> 11 Then
             rs("cuit02") = "0"
             rs.Update
        End If
       End If
     Else
       If rs("cuit02") <> "0" Then
         If Len(rs("cuit02")) < 7 Or Len(rs("cuit02")) > 8 Or Val(rs("cuit02")) <= 0 Then
           rs("cuit02") = "0"
           rs.Update
         End If
       End If
    End If
    rs.MoveNext
Wend
Set rs = Nothing


Set rs = New ADODB.Recordset
q = "select * from vta_01"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("id_tipoiva") <> 3 And rs("id_tipoiva") <> 8 Then
       'lleva cuit
       If Len(rs("cuit")) = 13 Then
          c = Mid$(rs("cuit"), 1, 2) & Mid$(rs("cuit"), 4, 8) & Mid$(rs("cuit"), 13, 1)
          rs("cuit") = c
          rs.Update
       Else
         If Len(rs("cuit")) <> 11 Then
             rs("cuit") = "11111111111"
         End If
       End If
     Else
       If rs("cuit") <> "0" Then
         If Len(rs("cuit")) < 7 Or Len(rs("cuit")) > 8 Or Val(rs("cuit")) <= 0 Then
           rs("cuit") = "0"
         End If
       End If
    End If
    rs.MoveNext
Wend
Set rs = Nothing

Unload espere
End If

End Sub

Sub actu130()
h = MsgBox("Actualizacion 130  . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh

Set rs = New ADODB.Recordset
q = "select * from pro_04"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("total_facturado") = rs("total_recibido")
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Unload espere
End If

End Sub

Sub actu132()
h = MsgBox("Actualizacion 132  . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh

Set rs = New ADODB.Recordset
q = "select * from a12"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("tipo12") = "P"
    Select Case rs("id_percepcion")
    Case Is = 1
      rs("impuesto12") = "I"
    Case Is = 2
      rs("impuesto12") = "B"
    Case Else
      rs("impuesto12") = "O"
    End Select
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Unload espere
End If

End Sub
Sub actu133()
h = MsgBox("Actualizacion 133  . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh

Set rs = New ADODB.Recordset
q = "select * from a13, a12 where a13.[id_percepcion] = a12.[id_percepcion]"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    Select Case rs("impuesto12")
    Case Is = "I"
      rs("cod_regimen") = 493
    Case Is = "G"
      rs("cod_regimen") = 78
    Case Is = "B"
      rs("cod_regimen") = 1
    Case Is = "S"
      rs("cod_regimen") = 353
    Case Else
      rs("cod_regimen") = 1
    End Select
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing

Unload espere
End If

End Sub

Sub actu999()
'modifica listad de iva vents que puso todos 111111
h = MsgBox("Actualizacion general . ¿Esta seguro que quiere actualizar? ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02, vta_01 where vta_02.[id_cliente] = vta_01.[id_cliente] and [cuit02] = '11111111111'"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("cuit02") = rs("cuit")
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If
End Sub


Sub actu106()
h = MsgBox("Actualizacion 106 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from a6"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    rs("Unidad06") = " "
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub

Sub actu104()
h = MsgBox("Actualizacion 104 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
espere.Show
espere.Refresh
Set rs = New ADODB.Recordset
q = "select * from vta_02"
rs.Open q, cn1, adOpenStatic, adLockOptimistic
a = 1
While Not rs.EOF
    espere.Label1 = "Espere... " & a
    espere.Label1.Refresh
    a = a + 1
    If rs("Estado_pago") = "N" Then
       If rs("moneda") = "P" Then
        rs("saldo_impago02") = rs("total")
       Else
        rs("saldo_impago02") = rs("total_otra_moneda")
       End If
    Else
        rs("saldo_impago02") = 0
        If rs("Moneda") = "P" Then
           i = rs("total")
        Else
           i = rs("Total_otra_moneda")
        End If
        
        sr = Val(Mid$(rs("recibo_pago"), 1, 4))
        nr = Val(Mid$(rs("recibo_pago"), 6, 8))
        Set rs2 = New ADODB.Recordset
        q = "select * from vta_02 where [sucursal] = " & sr & " and [num_comp] = " & nr & " and [id_tipocomp] = 50"
        rs2.Open q, cn1
        If Not rs2.EOF And Not rs2.BOF Then
          nir = rs2("num_int")
        Else
          nir = 0
        End If
        Set rs2 = Nothing
        
        If nir > 0 Then
         Set rs2 = New ADODB.Recordset
         q = "select * from vta_010"
         rs2.Open q, cn1, adOpenDynamic, adLockOptimistic
         rs2.AddNew
          rs2("num_int_comp") = rs("num_int")
          rs2("num_int_rbo") = nir
          rs2("importe_pagado") = i
          rs2("saldo_comprobante") = 0
         rs2.Update
        Set rs2 = Nothing
       End If
    End If
    rs.Update
    rs.MoveNext
Wend
Set rs = Nothing
Unload espere
End If

End Sub
Sub validaactu()
  On Error GoTo errv
 If Option1 = True Then
   o = 1
 Else
   o = 2
 End If
 If abrirconexion(o) Then
  Set rs = New ADODB.Recordset
  q = "select * from g0 where [sucursal] = 0"
  rs.Open q, cn1
  Label2 = rs("actualizacion")
  Set rs = Nothing
  cn1.Close
  Exit Sub
 
 Else
  Label2 = "N/C"
 End If
 
 
errv:
  Label2 = "N/D"
  cn1.Close
  Exit Sub
  
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Option1 = True
Call validaactu
End Sub
