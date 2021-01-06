package ifu.wsh  ;

import com4j.*;

/**
 */
public enum TiposPago implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  Cambio(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  CartaDeCreditoDocumentario(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  CartaDeCreditoSimple(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  Cheque(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  ChequeCancelatorios(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  CreditoDocumentario(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  CuentaCorriente(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  Deposito(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  Efectivo(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  EndosoDeCheque(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  FacturaDeCredito(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  GarantiaBancaria(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  Giro(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  LetraDeCambio(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  MedioDePagoDeComercioExterior(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  OrdenDePagoDocumentaria(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  OrdenDePagoSimple(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  PagoContraReembolso(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  RemesaDocumentaria(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  RemesaSimple(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  TarjetaDeCredito(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  TarjetaDeDebito(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  Ticket(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  TransferenciaBancaria(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  TransferenciaNoBancaria(24),
  /**
   * <p>
   * The value of this constant is 99
   * </p>
   */
  OtrosMediosPago(99),
  ;

  private final int value;
  TiposPago(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
