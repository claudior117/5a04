#FILE "C:\xampp\bitingenieria.com.ar\ifu\Ejemplos\Fujitsu Cobol\MainForm\Debug\PrintForm.PRC"
000001 IDENTIFICATION  DIVISION.
000002* PrintForm.
000003 PROGRAM-ID.     PrintForm.
000004 ENVIRONMENT     DIVISION.
000005 CONFIGURATION   SECTION.
#LINE 11
000011 SPECIAL-NAMES.
000012 REPOSITORY.
000013*<SCRIPT DIVISION="ENVIRONMENT", SECTION="CONFIGURATION", PARAGRAPH="REPOSITORY">
000014     CLASS COM AS "*OLE".
000015*</SCRIPT>
000016 .
000017 INPUT-OUTPUT    SECTION.
000018 FILE-CONTROL.
000019 DATA            DIVISION.
#LINE 24
000024 LINKAGE         SECTION.
000025 01  POW-FORM IS GLOBAL.
000026   02  POW-SELF PIC S9(9) COMP-5.
000027   02  POW-SUPER  PIC X(4).
000028   02  POW-THIS PIC S9(9) COMP-5.
000029   02  CmCommand1 PIC S9(9) COMP-5.
000030 01  PrintForm REDEFINES POW-FORM GLOBAL PIC S9(9) COMP-5.
000031 01  POW-CONTROL-ID PIC S9(9) COMP-5.
000032 01  POW-EVENT-ID   PIC S9(9) COMP-5.
000033 01  POW-OLE-PARAM  PIC X(4).
000034 01  POW-OLE-RETURN PIC X(4).
000035 PROCEDURE       DIVISION USING POW-FORM POW-CONTROL-ID POW-EVENT-ID POW-OLE-PARAM POW-OLE-RETURN.
000036     EVALUATE POW-CONTROL-ID
000037     WHEN 117440517
000038     EVALUATE POW-EVENT-ID
000039     WHEN -600
000040       CALL "POW-SCRIPTLET1"
000041     END-EVALUATE
000042     END-EVALUATE
000043     EXIT PROGRAM.
000044 IDENTIFICATION  DIVISION.
000045* CmCommand1-Click.
000046 PROGRAM-ID.     POW-SCRIPTLET1.
000047*<SCRIPT DIVISION="PROCEDURE", CONTROL="CmCommand1", EVENT="Click", POW-NAME="SCRIPTLET1", TYPE="ETC">
000048 ENVIRONMENT     DIVISION.
000049 DATA            DIVISION.
000050 WORKING-STORAGE SECTION.
000051 01 OBJ-DRIVER    OBJECT REFERENCE COM.
000052 01 PROGID-DRIVER PIC X(8192) VALUE "IFUniversal.Driver".
000053 01 IS-OK PIC S9(4) COMP-5.
000054 01 Modelo PIC S9(9) COMP-5 VALUE 23.
000055 01 Puerto PIC S9(9) COMP-5 VALUE 31.
000056 01 MSG_SUCESS PIC X(8192) VALUE "CIERRE REALIZADO CON EXITO".
#LINE 57,#START,#OTHER
000057 01 POW-0000 PIC S9(18) COMP-5.
000057 01 POW-0001 PIC S9(9) COMP-5.
000057 01 POW-0002 PIC S9(9) COMP-5.
000057 01 POW-0003 PIC S9(9) COMP-5.
000057 01 POW-0004 PIC S9(9) COMP-5.
000057 01 POW-0005 PIC S9(9) COMP-5.
000057 01 POW-0006 PIC X(8192).
000057 01 POW-0007 PIC S9(9) COMP-5.
#LINE 56,#END
000057 PROCEDURE       DIVISION.
000058     invoke COM "CREATE-OBJECT" using PROGID-DRIVER
000059                                returning OBJ-DRIVER.
000060     invoke OBJ-DRIVER "SET-MODELO" using Modelo.
000061     invoke OBJ-DRIVER "SET-PUERTO" using Puerto.
000062     invoke OBJ-DRIVER "Inicializar"
000063     invoke OBJ-DRIVER "GET-Error" returning IS-OK.
000064     IF IS-OK = 0 THEN
000065       invoke OBJ-DRIVER "cierreZ".
000066       invoke OBJ-DRIVER "GET-Error" returning IS-OK.
000067       IF IS-OK = 0 THEN
#LINE 68,#START,INVOKE(68,16)
000068     MOVE 117441026 TO POW-0000 
000068     MOVE 1 TO POW-0001 
000068     MOVE 16387 TO POW-0003 
000068     MOVE 0 TO POW-0004 
000068     MOVE 1 TO POW-0005 
000068     MOVE MSG_SUCESS TO POW-0006 
000068     MOVE 33636360 TO POW-0007 
000068     CALL "XPOW_INVOKE_BY_ID_2" USING VALUE POW-SELF REFERENCE POW-0000 
000068     VALUE POW-0001 POW-0003 REFERENCE POW-0002 VALUE POW-0004 POW-0005 
000068     POW-0007 REFERENCE POW-0006 END-CALL 
000068                                                          .
#LINE 68,#END
000069*</SCRIPT>
000070 END PROGRAM     POW-SCRIPTLET1.
000071 END PROGRAM     PrintForm.
#FILE