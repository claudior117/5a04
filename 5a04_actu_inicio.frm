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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "12"
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
      TabIndex        =   11
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Version"
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
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ultima Version 11 234 "
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
      Top             =   1080
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
      Top             =   1080
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
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Caption         =   "005"
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
      Top             =   600
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
  Case Is = 1
    Call actu001
  Case Is = 2
    Call actu002
  Case Is = 3
    Call actu003
  Case Is = 4
    Call actu004
  Case Is = 5
    Call actu005
  
  
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

Sub actu001()
'corrige ajustes de stock
h = MsgBox("Actualizacion 001 Version12 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table g0 add column [version_webservice] Int  "
     cn1.Execute q
     
  cn1.CommitTrans
  
 cn1.BeginTrans
    
    q = "update g0 set  [actualizacion]=001, [version_webservice]=3"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub

Sub actu003()
'corrige REPORTES
h = MsgBox("Actualizacion 003 Version12. ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  espere.Show
  espere.Refresh
  
   MsgBox ("Actualizacion tabla reportes. Modifique campo cod_barra como Text(20) en rep/dat/rep.mdb")
 
 cn1.BeginTrans
    q = "update g0 set  [actualizacion]=003"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
  cn1.CommitTrans
    
   
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu004()
'Agrega logo arca factura electronica
h = MsgBox("Actualizacion 004 Version12. Logo Arca¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  espere.Show
  espere.Refresh
  
   MsgBox ("Actualizacion logo arca. Copiar archivo arca.png en carpeta tools")
 
 cn1.BeginTrans
    q = "update g0 set  [actualizacion]=004"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
  cn1.CommitTrans
    
   
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu005()
'Corrige Venta directa (205/206/207)
h = MsgBox("Actualizacion 005 Version12. Corrige cantidades en Venta Directa ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  espere.Show
  espere.Refresh
  
   
   Set rs = New ADODB.Recordset
   q = "select * from vta_02, vta_03 where id_tipocomp >=205 and id_tipocomp <=207 and vta_02.num_int = vta_03.num_int"
   MsgBox (q)
   
   rs.Open q, cn1, adOpenDynamic, adLockOptimistic
   
   While Not rs.EOF And Not rs.BOF
       rs("cantidad_original") = rs("cantidad")
       rs("tunidad") = " "
       rs("bultos") = 1
       rs("pu_final") = rs("pu") + rs("pu") * rs("tasaiva") / 100
       rs("tasaib") = 0
       rs.Update
       rs.MoveNext
   Wend
 
  Set rs = Nothing
 
 
 cn1.BeginTrans
    q = "update g0 set  [actualizacion]=005"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
  cn1.CommitTrans
    
   
  
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub



Sub actu002()
'corrige ajustes de stock
h = MsgBox("Actualizacion 002 Version12 . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  
  espere.Show
  espere.Refresh
    
   cn1.BeginTrans
     
     q = "alter table g3 add column [cod_fe] Int  "
     cn1.Execute q
     
  cn1.CommitTrans
  
 cn1.BeginTrans
    
    q = "update g3 set  [cod_fe]=1 where [cod_tipoiva]=1"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=15 where [cod_tipoiva]=2"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=5 where [cod_tipoiva]=3"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=6 where [cod_tipoiva]=4"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=4 where [cod_tipoiva]=5"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=1 where [cod_tipoiva]=6"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=7 where [cod_tipoiva]=7"
    cn1.Execute q
    
    q = "update g3 set  [cod_fe]=9 where [cod_tipoiva]=8"
    cn1.Execute q
    
    q = "update g0 set  [actualizacion]=002"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
    
   
  
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
    
    q = "update i_01 set [id_otrostributos]=99, [tasa_i1]=0, [id_cuenta_i1] = 110302"
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

Sub actu231()
'remitos predefinidos
h = MsgBox("Actualizacion 231(Remitos Predefinidos) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  espere.Show
  espere.Refresh
   
   cn1.BeginTrans
   q = "update g0 set  [actualizacion]=231"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
  MsgBox ("Proceso terminado.  Importar tablas vta_017 y vta_018")
   
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub

Sub actu232()
'remitos predefinidos
h = MsgBox("Actualizacion 232(Habilita Sistema de Logs) . ¿Esta seguro que quiere actualizar?  ", 4)
If h = 6 Then
  espere.Show
  espere.Refresh
   
   cn1.BeginTrans
   q = "update g0 set  [actualizacion]=232"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
  MsgBox ("Proceso terminado.  Importar tabla g15")
   
   
   
 Unload espere
  
End If

Exit Sub


err1:
Resume Next

End Sub


Sub actu233()
'corrige oc
h = MsgBox("Actualizacion 233. Verificar si existe el campo tipo04 en PRO_04 ¿Existe?  ", 4)

  espere.Show
  espere.Refresh
   
   cn1.BeginTrans
   q = "update g0 set  [actualizacion]=233"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
    
  cn1.CommitTrans
  
  If h <> 6 Then
    cn1.BeginTrans
    q = "alter table pro_04 add column [tipo04] int "
     cn1.Execute q
    cn1.CommitTrans
  
      cn1.BeginTrans
     q = "update pro_04 set  [tipo04]=1"
   
     cn1.Execute q
    cn1.CommitTrans
  
  
  
  
  End If
  
  
  MsgBox ("Proceso terminado. Renovar tabla G15")
   
   
   
 Unload espere
  

Exit Sub


err1:
Resume Next

End Sub


Sub actu234()
'corrige oc
h = MsgBox("Actualizacion 234. Agrega PLU", 4)

  espere.Show
  espere.Refresh
   
   cn1.BeginTrans
   
     q = "alter table A2 add column [plu] Int default 0 "
     cn1.Execute q
      
 
   
   
   q = "update g0 set  [actualizacion]=234"
    q = q & " where [sucursal]=0 "
    cn1.Execute q
   
   
   
    q = "update a2 set plu = 0"
    cn1.Execute q
  cn1.CommitTrans
  
  MsgBox ("Proceso terminado.")
   
   
   
 Unload espere
  

Exit Sub


err1:
Resume Next

End Sub





Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
Option1 = True
Call validaactu
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
