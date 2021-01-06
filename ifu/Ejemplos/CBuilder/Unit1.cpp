//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "Unit1.h"
#include "IFUniversal_TLB.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TForm1 *Form1;
//---------------------------------------------------------------------------
__fastcall TForm1::TForm1(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TForm1::ImprimirComprobante(TipoDeComprobante Tipo, bool EnviaDatosCliente)
{
  CoInitialize(NULL); //Init COM library DLLs

  IDriver* Fiscal;

  HRESULT hr = CoCreateInstance ( CLSID_Driver,
                                  NULL,
                                  CLSCTX_INPROC_SERVER,
                                  IID_IDriver,
                                  (void**) &Fiscal );
  if (SUCCEEDED (hr)) {

    try{
      Fiscal->Modelo = Ifuniversal_tlb::modHasarPT1000F;

      if (Fiscal->Error != 0) {
        ShowMessage(Fiscal->ErrorDesc);
      }

      Fiscal->Puerto = 31;
      Fiscal->Baudios = Ifuniversal_tlb::bd9600;

      if (!Fiscal->Inicializar()) {
         throw new Exception(Fiscal->ErrorDesc);
      }

      Fiscal->CancelarComprobante();

      //Si el ticket supera 1000 pesos debe enviarse
      if (EnviaDatosCliente) {
        if (!Fiscal->DatosCliente(WideString("Abel Miranda"), Ifuniversal_tlb::tdCUIT, WideString("20939802593"),
          Ifuniversal_tlb::riResponsableInscripto, WideString("Blanco Encalada 1204 5to A"))) {
           throw new Exception(Fiscal->ErrorDesc);
        }
      }

      if (!Fiscal->AbrirComprobante(Tipo))
           throw new Exception(Fiscal->ErrorDesc);
      if (!Fiscal->ImprimirItem2g(WideString("Item 1"), 1.0, 0.1, 21.0, 0.0, Ifuniversal_tlb::Gravado,
        Ifuniversal_tlb::tiFijo, 1, WideString("7790001001054"), WideString(""), Ifuniversal_tlb::Unidad))
           throw new Exception(Fiscal->ErrorDesc);
      if (!Fiscal->ImprimirPago2g(WideString("Efectivo"), 5, WideString(""), Ifuniversal_tlb::Efectivo, 1, WideString(""), WideString("")))
           throw new Exception(Fiscal->ErrorDesc);

      Fiscal->CerrarComprobante();

      ShowMessage("Comprobante impreso exitosamente.");

    } catch(Exception* e) {
      ShowMessage(e->Message);
    }

       Fiscal->Release();
  }

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button1Click(TObject *Sender)
{
  ImprimirComprobante(tcFactura_B, false);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button2Click(TObject *Sender)
{
  ImprimirComprobante(tcFactura_A, true);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button3Click(TObject *Sender)
{
  ImprimirComprobante(tcTique, false);
}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button9Click(TObject *Sender)
{
  ImprimirComprobante(tcNota_Debito_B, false);
}
//---------------------------------------------------------------------------


void __fastcall TForm1::Button5Click(TObject *Sender)
{
  CoInitialize(NULL); //Init COM library DLLs

  IDriver* Fiscal;

  HRESULT hr = CoCreateInstance ( CLSID_Driver,
                                  NULL,
                                  CLSCTX_INPROC_SERVER,
                                  IID_IDriver,
                                  (void**) &Fiscal );
  if (SUCCEEDED (hr)) {

    try{
      Fiscal->Modelo = Ifuniversal_tlb::modHasarPT1000F;

      if (Fiscal->Error != 0) {
        ShowMessage(Fiscal->ErrorDesc);
      }

      Fiscal->Puerto = 31;
      Fiscal->Baudios = Ifuniversal_tlb::bd9600;

      if (!Fiscal->Inicializar()) {
         throw new Exception(Fiscal->ErrorDesc);
      }

      Fiscal->CancelarComprobante();

      Fiscal->CierreX();

      ShowMessage("Cierre realizado exitosamente.");

    } catch(Exception* e) {
      ShowMessage(e->Message);
    }

       Fiscal->Release();
  }

}
//---------------------------------------------------------------------------

void __fastcall TForm1::Button4Click(TObject *Sender)
{
  CoInitialize(NULL); //Init COM library DLLs

  IDriver* Fiscal;

  HRESULT hr = CoCreateInstance ( CLSID_Driver,
                                  NULL,
                                  CLSCTX_INPROC_SERVER,
                                  IID_IDriver,
                                  (void**) &Fiscal );
  if (SUCCEEDED (hr)) {

    try{
      Fiscal->Modelo = Ifuniversal_tlb::modHasarPT1000F;

      if (Fiscal->Error != 0) {
        ShowMessage(Fiscal->ErrorDesc);
      }

      Fiscal->Puerto = 31;
      Fiscal->Baudios = Ifuniversal_tlb::bd9600;

      if (!Fiscal->Inicializar()) {
         throw new Exception(Fiscal->ErrorDesc);
      }

      Fiscal->CancelarComprobante();

      Fiscal->CierreZ();

      ShowMessage("Cierre realizado exitosamente.");

    } catch(Exception* e) {
      ShowMessage(e->Message);
    }

       Fiscal->Release();
  }


}
//---------------------------------------------------------------------------

