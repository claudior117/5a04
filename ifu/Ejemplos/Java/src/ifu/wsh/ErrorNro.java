package ifu.wsh  ;

import com4j.*;

/**
 */
public enum ErrorNro implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  errNoError(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  errControladorNoDisponible(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  errComandoInvalido(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  errParametroInvalido(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  errExcepcion(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  errMemoriaFiscal(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  errMemoriaTrabajo(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  errBateriaBaja(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  errComandoDesconocido(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  errDesbordamientoTotales(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  errMemoriaFiscalLlena(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  errMemoriaFiscalCasiLlena(11),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  errFallaImpresora(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  errImpresoraFueraLinea(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  errFaltaPapelDiario(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  errFaltaPapelTicket(16),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  errTapaImpresoraAbierta(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  errCajonCerradoOAusente(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  errCampoDatosInvalido(20),
  ;

  private final int value;
  ErrorNro(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
