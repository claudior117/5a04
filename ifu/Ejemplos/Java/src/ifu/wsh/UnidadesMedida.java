package ifu.wsh  ;

import com4j.*;

/**
 */
public enum UnidadesMedida implements ComEnum {
  /**
   * <p>
   * The value of this constant is 0
   * </p>
   */
  SinDescripcion(0),
  /**
   * <p>
   * The value of this constant is 1
   * </p>
   */
  Kilo(1),
  /**
   * <p>
   * The value of this constant is 2
   * </p>
   */
  Metro(2),
  /**
   * <p>
   * The value of this constant is 3
   * </p>
   */
  MetroCuadrado(3),
  /**
   * <p>
   * The value of this constant is 4
   * </p>
   */
  MetroCubico(4),
  /**
   * <p>
   * The value of this constant is 5
   * </p>
   */
  Litro(5),
  /**
   * <p>
   * The value of this constant is 6
   * </p>
   */
  KWH(6),
  /**
   * <p>
   * The value of this constant is 7
   * </p>
   */
  Unidad(7),
  /**
   * <p>
   * The value of this constant is 8
   * </p>
   */
  Par(8),
  /**
   * <p>
   * The value of this constant is 9
   * </p>
   */
  Docena(9),
  /**
   * <p>
   * The value of this constant is 10
   * </p>
   */
  Quilate(10),
  /**
   * <p>
   * The value of this constant is 11
   * </p>
   */
  Millar(11),
  /**
   * <p>
   * The value of this constant is 12
   * </p>
   */
  MegaUInterActAntib(12),
  /**
   * <p>
   * The value of this constant is 13
   * </p>
   */
  UnidadInternaActInmung(13),
  /**
   * <p>
   * The value of this constant is 14
   * </p>
   */
  Gramo(14),
  /**
   * <p>
   * The value of this constant is 15
   * </p>
   */
  Milimetro(15),
  /**
   * <p>
   * The value of this constant is 16
   * </p>
   */
  MilimetroCubico(16),
  /**
   * <p>
   * The value of this constant is 17
   * </p>
   */
  Kilometro(17),
  /**
   * <p>
   * The value of this constant is 18
   * </p>
   */
  Hectolitro(18),
  /**
   * <p>
   * The value of this constant is 19
   * </p>
   */
  MegaUnidadIntActInmung(19),
  /**
   * <p>
   * The value of this constant is 20
   * </p>
   */
  Centimetro(20),
  /**
   * <p>
   * The value of this constant is 21
   * </p>
   */
  KilogramoActivo(21),
  /**
   * <p>
   * The value of this constant is 22
   * </p>
   */
  GramoActivo(22),
  /**
   * <p>
   * The value of this constant is 23
   * </p>
   */
  GramoBase(23),
  /**
   * <p>
   * The value of this constant is 24
   * </p>
   */
  UIACTHOR(24),
  /**
   * <p>
   * The value of this constant is 25
   * </p>
   */
  JuegoPaqueteMazoNaipes(25),
  /**
   * <p>
   * The value of this constant is 26
   * </p>
   */
  MUIACTHOR(26),
  /**
   * <p>
   * The value of this constant is 27
   * </p>
   */
  CentimetroCubico(27),
  /**
   * <p>
   * The value of this constant is 28
   * </p>
   */
  UIACTANT(28),
  /**
   * <p>
   * The value of this constant is 29
   * </p>
   */
  Tonelada(29),
  /**
   * <p>
   * The value of this constant is 30
   * </p>
   */
  DecametroCubico(30),
  /**
   * <p>
   * The value of this constant is 31
   * </p>
   */
  HectometroCubico(31),
  /**
   * <p>
   * The value of this constant is 32
   * </p>
   */
  KilometroCubico(32),
  /**
   * <p>
   * The value of this constant is 33
   * </p>
   */
  Microgramo(33),
  /**
   * <p>
   * The value of this constant is 34
   * </p>
   */
  Nanogramo(34),
  /**
   * <p>
   * The value of this constant is 35
   * </p>
   */
  Picogramo(35),
  /**
   * <p>
   * The value of this constant is 36
   * </p>
   */
  MUIACTANT(36),
  /**
   * <p>
   * The value of this constant is 37
   * </p>
   */
  UIACTIG(37),
  /**
   * <p>
   * The value of this constant is 41
   * </p>
   */
  Miligramo(41),
  /**
   * <p>
   * The value of this constant is 47
   * </p>
   */
  Mililitro(47),
  /**
   * <p>
   * The value of this constant is 48
   * </p>
   */
  Curie(48),
  /**
   * <p>
   * The value of this constant is 49
   * </p>
   */
  Milicurie(49),
  /**
   * <p>
   * The value of this constant is 50
   * </p>
   */
  Microcurie(50),
  /**
   * <p>
   * The value of this constant is 51
   * </p>
   */
  UInterActHormonal(51),
  /**
   * <p>
   * The value of this constant is 52
   * </p>
   */
  MegaUInterActHor(52),
  /**
   * <p>
   * The value of this constant is 53
   * </p>
   */
  KilogramoBase(53),
  /**
   * <p>
   * The value of this constant is 54
   * </p>
   */
  Gruesa(54),
  /**
   * <p>
   * The value of this constant is 55
   * </p>
   */
  MUIACTIG(55),
  /**
   * <p>
   * The value of this constant is 61
   * </p>
   */
  KilogramoBruto(61),
  /**
   * <p>
   * The value of this constant is 62
   * </p>
   */
  Pack(62),
  /**
   * <p>
   * The value of this constant is 63
   * </p>
   */
  Horma(63),
  /**
   * <p>
   * The value of this constant is 90
   * </p>
   */
  Donacion(90),
  /**
   * <p>
   * The value of this constant is 91
   * </p>
   */
  Ajustes(91),
  /**
   * <p>
   * The value of this constant is 96
   * </p>
   */
  Anulacion(96),
  /**
   * <p>
   * The value of this constant is 97
   * </p>
   */
  SenasAnticipos(97),
  /**
   * <p>
   * The value of this constant is 98
   * </p>
   */
  OtrasUnidades(98),
  /**
   * <p>
   * The value of this constant is 99
   * </p>
   */
  Bonificacion(99),
  ;

  private final int value;
  UnidadesMedida(int value) { this.value=value; }
  public int comEnumValue() { return value; }
}
