echo "genera clave privada, claveprivada.key va el nombre del archivo que contendra la clave "
openssl genrsa -out claveprivada.key 2048
echo .
echo .
echo "generar CRS, cambiar nombre del cliente(O), el niombre de certificado(CN) y el cuit del cliente"
openssl req -new -key claveprivada.key -subj "/C=AR/O=Claudio Ravagnan/CN=fact_e/serialNumber=CUIT 20202956034" -out claudiocsr.csr
pause 