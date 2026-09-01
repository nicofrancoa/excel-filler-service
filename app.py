from flask import Flask, request, jsonify, send_file
import openpyxl
import json
import os
import tempfile

app = Flask(__name__)

PLANTILLA = '/app/plantilla_copropiedades.xlsx'

COLUMNAS_P = {
    'MAPFRE': 'B', 'SBS': 'C', 'PREVISORA': 'D',
    'SEGUROS DEL ESTADO': 'E', 'AXA COLPATRIA': 'F',
    'ZURICH': 'G', 'ALLIANZ': 'H', 'EQUIDAD': 'I',
    'HDI': 'J', 'ASEGURADORA SOLIDARIA': 'K'
}

COLUMNAS_C = {
    'MAPFRE': 'C', 'SBS': 'D', 'PREVISORA': 'E',
    'SEGUROS DEL ESTADO': 'F', 'ZURICH': 'G',
    'AXA COLPATRIA': 'H', 'ALLIANZ': 'I',
    'SEGUROS BOLIVAR': 'J', 'EQUIDAD': 'K',
    'HDI': 'L', 'ASEGURADORA SOLIDARIA': 'M'
}

COLUMNAS_R = {
    'MAPFRE': 'B', 'SBS': 'C', 'PREVISORA': 'D',
    'SEGUROS DEL ESTADO': 'E', 'ZURICH': 'F',
    'AXA COLPATRIA': 'G', 'ALLIANZ': 'H',
    'SEGUROS BOLIVAR': 'I', 'EQUIDAD': 'J',
    'HDI': 'K', 'ASEGURADORA SOLIDARIA': 'L'
}

COLUMNAS_E = {
    'MAPFRE': ['B','C'], 'SBS': ['D','E'], 'PREVISORA': ['F','G'],
    'SEGUROS DEL ESTADO': ['H','I'], 'AXA COLPATRIA': ['J','K'],
    'ZURICH': ['L','M'], 'ALLIANZ': ['N','O'],
    'SEGUROS BOLIVAR': ['P','Q'], 'EQUIDAD': ['R','S'],
    'HDI': ['T','U'], 'ASEGURADORA SOLIDARIA': ['V','W']
}

@app.route('/llenar', methods=['POST'])
def llenar():
    datos = request.get_json()
    wb = openpyxl.load_workbook(PLANTILLA)
    wsP = wb['PRIMA1']
    wsC = wb['copropiedades']
    wsR = wb['RCE']
    wsE = wb['Ejemplos de siniestro']

    for aseg in datos.get('aseguradoras', []):
        nombre = aseg.get('nombre', '').upper()
        colP = COLUMNAS_P.get(nombre)
        colC = COLUMNAS_C.get(nombre)
        colR = COLUMNAS_R.get(nombre)
        colsE = COLUMNAS_E.get(nombre)

        if colP:
            wsP[f'{colP}3'] = aseg.get('prima', {}).get('danos_materiales') or ''
            wsP[f'{colP}4'] = aseg.get('prima', {}).get('rce') or 'INCLUYE'
            wsP[f'{colP}5'] = aseg.get('prima', {}).get('dao') or 'INCLUYE'
            wsP[f'{colP}6'] = aseg.get('prima', {}).get('asistencia') or 'INCLUYE'

        if colC:
            b = aseg.get('bienes', {})
            wsC[f'{colC}6'] = b.get('areas_comunes') or ''
            wsC[f'{colC}7'] = b.get('cimientos') or ''
            wsC[f'{colC}8'] = b.get('muebles_enseres') or ''
            wsC[f'{colC}9'] = b.get('equipo_electronico') or ''
            wsC[f'{colC}10'] = b.get('maquinaria') or ''
            wsC[f'{colC}11'] = b.get('dineros') or ''
            wsC[f'{colC}12'] = b.get('total') or ''
            a = aseg.get('amparos', {})
            wsC[f'{colC}14'] = a.get('todo_riesgo') or ''
            wsC[f'{colC}15'] = a.get('hurto_simple') or ''
            wsC[f'{colC}16'] = a.get('rotura_vidrios') or ''
            wsC[f'{colC}17'] = a.get('dao') or ''
            wsC[f'{colC}18'] = a.get('manejo_global') or ''
            d = aseg.get('deducibles', {})
            wsC[f'{colC}23'] = d.get('incendio') or ''
            wsC[f'{colC}24'] = d.get('agua_anegacion') or ''
            wsC[f'{colC}25'] = d.get('hmacc') or ''
            wsC[f'{colC}26'] = d.get('amit') or ''
            wsC[f'{colC}27'] = d.get('terremoto') or ''
            wsC[f'{colC}28'] = d.get('hurto_calificado') or ''
            wsC[f'{colC}29'] = d.get('hurto_simple_eee') or ''
            wsC[f'{colC}30'] = d.get('dano_maquinaria') or ''
            wsC[f'{colC}31'] = d.get('dano_electronico') or ''
            wsC[f'{colC}32'] = d.get('asistencia') or ''
            wsC[f'{colC}33'] = d.get('rotura_vidrios') or ''
            wsC[f'{colC}34'] = d.get('dao') or ''
            wsC[f'{colC}35'] = d.get('manejo_global') or ''
            dm = aseg.get('demerito', {})
            wsC[f'{colC}37'] = dm.get('equipo_electronico') or ''
            wsC[f'{colC}38'] = dm.get('maquinaria') or ''
            ac = aseg.get('asistencia_comun', {})
            wsC[f'{colC}43'] = ac.get('eventos') or ''
            wsC[f'{colC}44'] = ac.get('monto_evento') or ''
            wsC[f'{colC}45'] = ac.get('servicios_base') or ''
            wsC[f'{colC}46'] = ac.get('juridica') or ''
            ap = aseg.get('asistencia_privada', {})
            wsC[f'{colC}48'] = ap.get('eventos') or ''
            wsC[f'{colC}49'] = ap.get('monto_evento') or ''
            wsC[f'{colC}50'] = ap.get('servicios_base') or ''

        if colR:
            ra = aseg.get('rce', {}).get('amparos', {})
            wsR[f'{colR}4'] = ra.get('basico') or ''
            wsR[f'{colR}5'] = ra.get('patronal') or ''
            wsR[f'{colR}6'] = ra.get('contratistas') or ''
            wsR[f'{colR}7'] = ra.get('cruzada') or ''
            wsR[f'{colR}8'] = ra.get('bienes_cuidado') or ''
            wsR[f'{colR}9'] = ra.get('extrapatrimoniales') or ''
            wsR[f'{colR}10'] = ra.get('parqueaderos') or ''
            wsR[f'{colR}11'] = ra.get('gastos_medicos') or ''
            wsR[f'{colR}12'] = ra.get('mascotas') or ''
            wsR[f'{colR}13'] = ra.get('vehiculos') or ''
            rd = aseg.get('rce', {}).get('deducibles', {})
            wsR[f'{colR}17'] = rd.get('basico') or ''
            wsR[f'{colR}18'] = rd.get('patronal') or ''
            wsR[f'{colR}19'] = rd.get('contratistas') or ''
            wsR[f'{colR}20'] = rd.get('cruzada') or ''
            wsR[f'{colR}21'] = rd.get('bienes_cuidado') or ''
            wsR[f'{colR}22'] = rd.get('extrapatrimoniales') or ''
            wsR[f'{colR}23'] = rd.get('parqueaderos') or ''
            wsR[f'{colR}24'] = rd.get('gastos_medicos') or ''
            wsR[f'{colR}25'] = rd.get('mascotas') or ''
            wsR[f'{colR}26'] = rd.get('vehiculos') or ''

        if colsE:
            d = aseg.get('deducibles', {})
            wsE[f'{colsE[0]}4'] = d.get('terremoto_porcentaje') or 0.02
            wsE[f'{colsE[1]}4'] = d.get('terremoto') or '2% del valor asegurable'

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    return send_file(tmp.name, as_attachment=True, download_name='comparativo.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
