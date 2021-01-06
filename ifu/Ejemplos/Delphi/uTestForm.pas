unit uTestForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics,
  Controls, Forms, Dialogs, StdCtrls,
  IFUniversal_TLB;

const

  MODELO = IFUniversal_TLB.modHasarPT1000F;
  PUERTO = 31;

type

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    Button6: TButton;
    Button8: TButton;
    Button9: TButton;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
  private
    { Private declarations }
    procedure ImprimirComprobante(Tipo: TipoDeComprobante; EnviaDatosCliente: boolean);
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation


{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
begin
  ImprimirComprobante(tcFactura_B, false);
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  ImprimirComprobante(tcFactura_A, true);
end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  ImprimirComprobante(tcTique, false);
end;

procedure TForm1.Button4Click(Sender: TObject);
var
  Fiscal: IDriver;
begin
  Fiscal := CoDriver.Create;
  try

    Fiscal.Modelo := MODELO;

    if Fiscal.Error <> 0 then
      ShowMessage(Fiscal.ErrorDesc);

    Fiscal.Puerto := PUERTO;
    Fiscal.Baudios := bd9600;

    if not Fiscal.Inicializar then
      raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CancelarComprobante;

    if not Fiscal.CierreZ then
      raise Exception.Create(Fiscal.ErrorDesc);

    ShowMessage('Cierre realizado exitosamente.');

  except
    on E:Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TForm1.Button5Click(Sender: TObject);
var
  Fiscal: IDriver;
begin
  Fiscal := CoDriver.Create;
  try

    Fiscal.Modelo := MODELO;

    if Fiscal.Error <> 0 then
      ShowMessage(Fiscal.ErrorDesc);


    Fiscal.Puerto := PUERTO;
    Fiscal.Baudios := bd9600;

    if not Fiscal.Inicializar then
      raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CancelarComprobante;

    if not Fiscal.CierreX then
      raise Exception.Create(Fiscal.ErrorDesc);

    ShowMessage('Cierre realizado exitosamente.');

  except
    on E:Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TForm1.Button6Click(Sender: TObject);
var
  Fiscal: IDriver;
begin
  Fiscal := CoDriver.Create;
  try

    Fiscal.Modelo := MODELO;

    if Fiscal.Error <> 0 then
      ShowMessage(Fiscal.ErrorDesc);


    Fiscal.Puerto := PUERTO;
    Fiscal.Baudios := bd9600;

    if not Fiscal.Inicializar then
      raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CancelarComprobante;

    if not Fiscal.AbrirComprobante(tcNo_Fiscal) then
      raise Exception.Create(Fiscal.ErrorDesc);

    if not Fiscal.ImprimirTextoNoFiscal('Texto no fiscal') then
      raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CerrarComprobante;

    ShowMessage('Comprobante emitido exitosamente.');

  except
    on E:Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TForm1.Button7Click(Sender: TObject);
var
  Fiscal: IDriver;
begin
  Fiscal := CoDriver.Create;
  try

    Fiscal.Modelo := modHasarPR5F;

    if Fiscal.Error <> 0 then
      ShowMessage(Fiscal.ErrorDesc);


    Fiscal.Puerto := pcCOM2;
    Fiscal.Baudios := bd9600;

    Fiscal.Inicializar;

    Fiscal.CancelarComprobante;

    Fiscal.AbrirComprobante(tcFactura_B);

    Fiscal.ImprimirItem('prueba', 1, 10, 21, 0);

    Fiscal.ImprimirPago('Efectivo', 12.1);

    Fiscal.CerrarComprobante;

    Fiscal.DNFHFarmacias('Obra Social', 'Coseguro1', 'Coseguro2', 'Coseguro3',
      '123456', 'Abel Miranda', '20160101', 'Domicilio', 'Domicilio2', 'Establecimiento',
      '12345', 'Nota1', 'Nota2');

  except
    on E:Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TForm1.Button8Click(Sender: TObject);
var
  Fiscal: IDriver;
begin
  Fiscal := CoDriver.Create;
  try

    Fiscal.Modelo := MODELO;

    if Fiscal.Error <> 0 then
      ShowMessage(Fiscal.ErrorDesc);


    Fiscal.Puerto := PUERTO;
    Fiscal.Baudios := bd9600;

    if not Fiscal.Inicializar then
      raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CancelarComprobante;

    if not Fiscal.ReporteZFechas('160101', '160130', true) then
      raise Exception.Create(Fiscal.ErrorDesc);

    ShowMessage('Reporte realizado exitosamente.');
  except
    on E:Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
  ImprimirComprobante(tcNota_Debito_B, false);
end;

procedure TForm1.ImprimirComprobante(Tipo: TipoDeComprobante; EnviaDatosCliente: boolean);
var
  Fiscal: IDriver;
begin
  Fiscal := CoDriver.Create;

  try

    Fiscal.Modelo := MODELO;

    if Fiscal.Error <> 0 then
      ShowMessage(Fiscal.ErrorDesc);


    Fiscal.Puerto := PUERTO;
    Fiscal.Baudios := bd9600;

    if not Fiscal.Inicializar then
       Raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CancelarComprobante;

    //Si el ticket supera 1000 pesos debe enviarse
    if EnviaDatosCliente then begin
      If Not Fiscal.DatosCliente('Abel Miranda', tdCUIT, '20939802593', riMonotributo, 'Blanco Encalada 1204 5to A') Then
         Raise Exception.Create(Fiscal.ErrorDesc);
    end;

    If Not Fiscal.AbrirComprobante(Tipo) Then
       Raise Exception.Create(Fiscal.ErrorDesc);
    If Not Fiscal.ImprimirItem2g('Item 1', 1, 0.1, 21, 0, IFUniversal_TLB.Gravado, IFUniversal_TLB.tiFijo, 1, '7790001001054', '', IFUniversal_TLB.Unidad) Then
       Raise Exception.Create(Fiscal.ErrorDesc);
    If Not Fiscal.ImprimirPago2g('Efectivo', 5, '', IFUniversal_TLB.Efectivo, 1, '', '') Then
       Raise Exception.Create(Fiscal.ErrorDesc);

    Fiscal.CerrarComprobante;

    ShowMessage('Comprobante impreso exitosamente.');

  except
    on E:Exception do
      ShowMessage(E.Message);
  end;
end;

end.
