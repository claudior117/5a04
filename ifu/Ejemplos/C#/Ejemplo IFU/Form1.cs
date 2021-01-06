using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Ejemplo_IFU
{
    public partial class Form1 : Form
    {
        private static IFUniversal.ModeloPrn MODELO = IFUniversal.ModeloPrn.modEpsonTMU220AF;
        private static int PUERTO = 4;

        public Form1()
        {
            InitializeComponent();
        }

        private void ImprimirComprobante(IFUniversal.TipoDeComprobante Tipo, Boolean EnviaDatosCliente)
        {

            try {

                IFUniversal.IDriver Fiscal = new IFUniversal.Driver(); 
                Fiscal.Modelo = MODELO;

                if (Fiscal.Error != 0) {
                  MessageBox.Show(Fiscal.ErrorDesc);
                }

                Fiscal.Puerto = PUERTO;
                Fiscal.Baudios = IFUniversal.Baudio.bd9600;

                if (!Fiscal.Inicializar())
                    throw new Exception(Fiscal.ErrorDesc); 

                Fiscal.CancelarComprobante();

                //Si el ticket supera 1000 pesos debe enviarse
                if (EnviaDatosCliente) {
                  if (!Fiscal.DatosCliente("Abel Miranda", IFUniversal.TipoDeDocumento.tdCUIT, "20939802593", IFUniversal.ResponsabilidadIVA.riResponsableInscripto, "Blanco Encalada 1204 5to A")){
                    throw new Exception(Fiscal.ErrorDesc);
                  }
                }

                if (!Fiscal.AbrirComprobante(Tipo)) {
                    throw new Exception(Fiscal.ErrorDesc);
                }
                if (!Fiscal.ImprimirTextoFiscal("Devolucion Item")) {
                    throw new Exception(Fiscal.ErrorDesc);
                }
                if (!Fiscal.ImprimirItem("Item 1", 2, 100, 21, 0)) {
                    throw new Exception(Fiscal.ErrorDesc);
                }

                if (!Fiscal.ImprimirItem("Item 2", 2, 100, 21, 0)) {
                    throw new Exception(Fiscal.ErrorDesc);
                }

                if (!Fiscal.ImprimirDescuentoUltimoItem("Descuento Item", 10)) {
                    throw new Exception(Fiscal.ErrorDesc);
                }

                if (!Fiscal.ImprimirDescuentoGeneral("Item 1", 10)) {
                    throw new Exception(Fiscal.ErrorDesc);
                }

                if (!Fiscal.ImprimirPago("Visa", 100)) {
                    throw new Exception(Fiscal.ErrorDesc);
                }

                if (!Fiscal.ImprimirPago("Efectivo", 100)){
                    throw new Exception(Fiscal.ErrorDesc);
                }

                Fiscal.CerrarComprobante();

                MessageBox.Show("Comprobante impreso exitosamente");

            } catch (Exception E){
                  MessageBox.Show(E.Message);
            }
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            ImprimirComprobante(IFUniversal.TipoDeComprobante.tcFactura_A, true);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ImprimirComprobante(IFUniversal.TipoDeComprobante.tcFactura_B, true);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ImprimirComprobante(IFUniversal.TipoDeComprobante.tcTique, false);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try {

                IFUniversal.IDriver Fiscal = new IFUniversal.Driver(); 
                Fiscal.Modelo = MODELO;

                if (Fiscal.Error != 0)
                    throw new Exception(Fiscal.ErrorDesc);
                
                Fiscal.Puerto = PUERTO;
                Fiscal.Baudios = IFUniversal.Baudio.bd9600;

                if (!Fiscal.Inicializar()) 
                    throw new Exception(Fiscal.ErrorDesc); 
                

                Fiscal.CancelarComprobante();

                if (!Fiscal.CierreZ())
                    throw new Exception(Fiscal.ErrorDesc);

                MessageBox.Show("Cierre realizado exitosamente");
            }
            catch (Exception E)
            {
                  MessageBox.Show(E.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {

                IFUniversal.IDriver Fiscal = new IFUniversal.Driver();
                Fiscal.Modelo = MODELO;

                if (Fiscal.Error != 0)
                    throw new Exception(Fiscal.ErrorDesc); 

                Fiscal.Puerto = PUERTO;
                Fiscal.Baudios = IFUniversal.Baudio.bd9600;

                if (!Fiscal.Inicializar())
                    throw new Exception(Fiscal.ErrorDesc); 

                Fiscal.CancelarComprobante();

                if (!Fiscal.CierreX())
                    throw new Exception(Fiscal.ErrorDesc);

                MessageBox.Show("Cierre realizado exitosamente");
            }
            catch (Exception E)
            {
                MessageBox.Show(E.Message);
            }

        }
    }
}
