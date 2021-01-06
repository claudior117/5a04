Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms

Public Partial Class Form1
	Inherits Form
	Private Shared MODELO As IFUniversal.ModeloPrn = IFUniversal.ModeloPrn.modEpsonTMU220AF
	Private Shared PUERTO As Integer = 4

	Public Sub New()
		InitializeComponent()
	End Sub

	Private Sub ImprimirComprobante(Tipo As IFUniversal.TipoDeComprobante, EnviaDatosCliente As [Boolean])

		Try

			Dim Fiscal As IFUniversal.IDriver = New IFUniversal.Driver()
			Fiscal.Modelo = MODELO

			If Fiscal.[Error] <> 0 Then
				MessageBox.Show(Fiscal.ErrorDesc)
			End If

			Fiscal.Puerto = PUERTO
			Fiscal.Baudios = IFUniversal.Baudio.bd9600

			If Not Fiscal.Inicializar() Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			Fiscal.CancelarComprobante()

			'Si el ticket supera 1000 pesos debe enviarse
			If EnviaDatosCliente Then
				If Not Fiscal.DatosCliente("Abel Miranda", IFUniversal.TipoDeDocumento.tdCUIT, "20939802593", IFUniversal.ResponsabilidadIVA.riResponsableInscripto, "Blanco Encalada 1204 5to A") Then
					Throw New Exception(Fiscal.ErrorDesc)
				End If
			End If

            '****** USAR ESTE METODO PARA INFORMAR COMPROBANTES RELACIONADOS EN CASOS DE FACTURAS,NC, ND ********
            '  If Not Fiscal.DocumentoDeReferencia2g(tcRemito, "0001-00000001") Then
            '     Err.Raise Fiscal.Error, "", Fiscal.ErrorDesc
            '  End If

            If Not Fiscal.AbrirComprobante(Tipo) Then
                Throw New Exception(Fiscal.ErrorDesc)
            End If
			If Not Fiscal.ImprimirTextoFiscal("Devolucion Item") Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If
			If Not Fiscal.ImprimirItem("Item 1", 2, 100, 21, 0) Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			If Not Fiscal.ImprimirItem("Item 2", 2, 100, 21, 0) Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			If Not Fiscal.ImprimirDescuentoUltimoItem("Descuento Item", 10) Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			If Not Fiscal.ImprimirDescuentoGeneral("Item 1", 10) Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			If Not Fiscal.ImprimirPago("Visa", 100) Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			If Not Fiscal.ImprimirPago("Efectivo", 100) Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			Fiscal.CerrarComprobante()


			MessageBox.Show("Comprobante impreso exitosamente")
		Catch E As Exception
			MessageBox.Show(E.Message)
		End Try
	End Sub

	Private Sub Button3_Click(sender As Object, e As EventArgs)
		ImprimirComprobante(IFUniversal.TipoDeComprobante.tcFactura_A, True)
	End Sub

	Private Sub button1_Click(sender As Object, e As EventArgs)
		ImprimirComprobante(IFUniversal.TipoDeComprobante.tcFactura_B, True)
	End Sub

	Private Sub button2_Click(sender As Object, e As EventArgs)
		ImprimirComprobante(IFUniversal.TipoDeComprobante.tcTique, False)
	End Sub

	Private Sub button4_Click(sender As Object, e__1 As EventArgs)
		Try

			Dim Fiscal As IFUniversal.IDriver = New IFUniversal.Driver()
			Fiscal.Modelo = MODELO

			If Fiscal.[Error] <> 0 Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			Fiscal.Puerto = PUERTO
			Fiscal.Baudios = IFUniversal.Baudio.bd9600

			If Not Fiscal.Inicializar() Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If


			Fiscal.CancelarComprobante()

			If Not Fiscal.CierreZ() Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			MessageBox.Show("Cierre realizado exitosamente")
		Catch E As Exception
            MessageBox.Show(E.Message)
		End Try
	End Sub

	Private Sub button5_Click(sender As Object, e__1 As EventArgs)
		Try

			Dim Fiscal As IFUniversal.IDriver = New IFUniversal.Driver()
			Fiscal.Modelo = MODELO

			If Fiscal.[Error] <> 0 Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			Fiscal.Puerto = PUERTO
			Fiscal.Baudios = IFUniversal.Baudio.bd9600

			If Not Fiscal.Inicializar() Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			Fiscal.CancelarComprobante()

			If Not Fiscal.CierreX() Then
				Throw New Exception(Fiscal.ErrorDesc)
			End If

			MessageBox.Show("Cierre realizado exitosamente")
		Catch E As Exception
            MessageBox.Show(E.Message)
		End Try

	End Sub
End Class
