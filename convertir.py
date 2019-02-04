#!/usr/bin/env python
# encoding=utf8
import sys
reload(sys)
sys.setdefaultencoding('utf8')

#leer xml
import xml.etree.ElementTree as ET
import csv
from os import listdir
from os.path import isfile, join
    

#Abrir CSV
Convertidos = open('../../../convertidos.csv', 'w')
#Escribir
csvwriter = csv.writer(Convertidos)

namespaces = {'cfdi': 'http://www.sat.gob.mx/cfd/3', 'tfd' : 'http://www.sat.gob.mx/TimbreFiscalDigital', 'pago10': 'http://www.sat.gob.mx/Pagos', 'nomina12': 'http://www.sat.gob.mx/nomina12'} 


#encabezado del csv
cabeza = ['Tipo', 'Serie', 'Folio', 'Fecha Emision', 'Rfc Emisor', 'Razon Social Emisor', 'Rfc Receptor', 'Razon Social Receptor', 'Forma Pago', 'Metodo de Pago', 'Tipo Cambio', 'Moneda', 'Uso Cfdi', 'SubTotal', 'Descuento', 'I.V.A.', 'I.E.P.S.', 'I.V.A. RET', 'I.S.R.', 'Otros Impuestos', 'Total', 'Uuid', 'Fecha Pago', 'Forma de Pago', 'Moneda Pago', 'Monto Pago', 'Uuid Cfdi']


#obtener archivos en la ruta del py
ruta = "../../../"
archivos = [f for f in listdir(ruta) if isfile(join(ruta, f))]

print "Iniciando"

import os
print "Ruta", os.path.dirname(os.path.abspath(__file__))
#print 
cant = 0
for archivo in archivos:
    if archivo[-4:] == ".xml":
        cant = cant + 1
print "Hay", len(archivos) - cant, "archivos en la carpeta que no son xml"
print "Hay", cant, "archivos xml a convertir ..."


csvwriter.writerow(cabeza)
print "Escribiendo encabezado"
cant  = 0

for archivo in archivos:
    #si no es xml
    if archivo[-4:] != ".xml":
        print "No convertido", archivo
        continue

    cant = cant + 1
    print "Convirtiendo archivo", cant
    renglon = []
    
    #explorar el archivo
    arbol = ET.parse(ruta + archivo)
    raiz = arbol.getroot()    
    
    try:
        x = raiz.attrib['TipoDeComprobante']
        if x == "I":
            x = "Factura"
        elif x == "E":
            x = "Nota de crédito"
        elif x == "T":
            x = "Traslado"
        elif x == "N":
            x = "Nómina"
        elif x == "P":
            x = "Pago"
        else:
            x = "--"
    except:
        x = ""
    renglon.append(x)

    x = ""
    try:
        x = raiz.attrib['Serie']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.attrib['Folio']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.attrib['Fecha']
    except:
        x = ""
    renglon.append(x)


    try:
        x = raiz.find('cfdi:Emisor', namespaces).attrib['Rfc']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.find('cfdi:Emisor', namespaces).attrib['Nombre']
    except:
        x = ""
    renglon.append(x)
    
    try:
        x = raiz.find('cfdi:Receptor', namespaces).attrib['Rfc']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.find('cfdi:Receptor', namespaces).attrib['Nombre']
    except:
        x = ""
    renglon.append(x)


    try:
        x = raiz.attrib['FormaPago']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.attrib['MetodoPago']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.attrib['TipoCambio']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.attrib['Moneda']
    except:
        x = ""
    renglon.append(x)

    try:
        x = raiz.find('cfdi:Receptor', namespaces).attrib['UsoCFDI']
    except:
        x = ""
    renglon.append(x)

    try:
        x = raiz.attrib['SubTotal']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.attrib['Descuento']
    except:
        x = ""
    renglon.append(x)

    ivatras, iepstras = "", ""
    try:
        Imp = raiz.find('cfdi:Impuestos', namespaces)
        Tra = Imp.find('cfdi:Traslados', namespaces)
        for tras in Tra:
            if tras.attrib['Impuesto'] == '002':
                ivatras = tras.attrib['Importe']
            elif tras.attrib['Impuesto'] == '003':
                iepstras = tras.attrib['Importe']
    except:
        pass
    
    ivaret, iepsret, isrret = "", "", ""
    try:
        Imp = raiz.find('cfdi:Impuestos', namespaces)
        Ret = Imp.find('cfdi:Retenciones', namespaces)
        for rets in Ret:
            if rets.attrib['Impuesto'] == '002':
                ivaret = rets.attrib['Importe']
            elif rets.attrib['Impuesto'] == '003':
                iepsret = rets.attrib['Importe']
            elif rets.attrib['Impuesto'] == '001':
                isrret = rets.attrib['Importe']
    except:
        try:
            Imp = raiz.find('cfdi:Complemento', namespaces)
            Ret = Imp.find('nomina12:Nomina', namespaces)
            Z = Ret.find('nomina12:Deducciones', namespaces)
            isrret = Z.attrib['TotalImpuestosRetenidos']
            iepsret = Z.attrib['TotalOtrasDeducciones']
        except:
            pass    
    renglon.append(ivatras)
    renglon.append(iepstras)
    renglon.append(ivaret)
    renglon.append(isrret)
    renglon.append(iepsret)

    try:
        x = raiz.attrib['Total']
    except:
        x = ""
    renglon.append(x)

    try:
        x = raiz.find('cfdi:Complemento', namespaces).find('tfd:TimbreFiscalDigital', namespaces).attrib['UUID']
    except:
        x = ""
    renglon.append(x)

    try:
        x = raiz.find('cfdi:Complemento', namespaces).find('pago10:Pagos', namespaces)[0].attrib['FechaPago']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.find('cfdi:Complemento', namespaces).find('pago10:Pagos', namespaces)[0].attrib['FormaDePagoP']
    except:
        x = ""
    renglon.append(x)
    try:
        x = raiz.find('cfdi:Complemento', namespaces).find('pago10:Pagos', namespaces)[0].attrib['MonedaP']
    except:
        x = ""
    renglon.append(x)

    try:
        P = raiz.find('cfdi:Complemento', namespaces).find('pago10:Pagos', namespaces)
        suma = 0
        for pago in P:
            a, b = pago.attrib['Monto'].split('.')
            a = int(a)
            b = int (b)
            suma = suma + a * 100 + b
        x = str(suma // 100) + '.' + str(suma%100/10) + str(suma%10)
    except:
        x = ""
    renglon.append(x)

    try:
        x = raiz.find('cfdi:Complemento', namespaces).find('tfd:TimbreFiscalDigital', namespaces).attrib['UUID']
    except:
        x = ""
    renglon.append(x)

    csvwriter.writerow(renglon)
    print "listo", archivo[0:5] + "..." + archivo[-7:-4] + ".xml"
Convertidos.close()
print "Finalizado"
print "Debe aparecer un archivo llamado convertidos.csv en la misma carpeta "