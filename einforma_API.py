from flask import Flask, request, jsonify
from flask_restful import Resource, Api
from marshmallow import Schema, fields, validate
import pandas as pd
import subprocess
import json, robot, os

#init app
app = Flask(__name__)
api = Api(app)

# Departamentos
Departamentos = {
    'AMAZONAS':'91',
    'ANTIOQUIA':'05',
    'ARAUCA': '81',
    'ATLANTICO':'08',
    'BOGOTA':'11',
    'BOLIVAR':'13',
    'BOYACA':'15',
    'CALDAS':'17',
    'CAQUETA':'18',
    'CASANARE':'85',
    'CAUCA':'19',
    'CESAR':'20',
    'CHOCO':'27',
    'CORDOBA':'23',
    'CUNDINAMARCA':'25',
    'GUAINIA':'94',
    'GUAVIARE':'95',
    'HUILA':'41',
    'LA GUAJIRA':'44',
    'MAGDALENA':'47',
    'META':'50',
    'NARIÑO':'52',
    'NORTE SANTANDER':'54',
    'PUTUMAYO':'86',
    'QUINDIO':'63',
    'RISARALDA':'66',
    'SAN ANDRES':'88',
    'SANTANDER':'68',
    'SUCRE':'70',
    'TOLIMA':'73',
    'VALLE':'76',
    'VAUPES':'97',
    'VICHADA':'99',
    '':'00'
}

# Valores para reemplazar
replacements = (
    ("á", "a"),
    ("é", "e"),
    ("í", "i"),
    ("ó", "o"),
    ("ú", "u"),
)

# Variables Gobales
Departamento_id = '00'

# Validación de la departamento, la cual debe de estar entre los datos de la lista Departamentps
def ValidaDepartamento(departamento: str):
    global Departamento_id

    # Convierte a Mayusculas
    departamento = departamento.upper()

    # Reemplaza tildes
    for a, b in replacements:
        departamento = departamento.replace(a, b).replace(a.upper(), b.upper())

    # Valida si es un departamento valido
    departamento_valido = departamento in Departamentos
    if departamento_valido:
        Departamento_id = Departamentos[departamento]

    return departamento_valido



# Clase del schema o modelo para realizar validaciones
class EmpresaSchema(Schema):
    busqueda = fields.String(required=True)
    departamento = fields.String(missing=None, validate=ValidaDepartamento)
    n = fields.Integer(required=True, validate=validate.Range(min=1))

# Leer archivo retornado por robot
def read_file():
    data = pd.read_excel('Info_Empresas.xlsx') 
    df = pd.DataFrame(data, columns= [
        'Razon_Social', 'Forma_Juridica', 'Departamento', 'Actividad_CIIU', 'Fecha_Constitucion', 'Fecha_Ultimo_Dato',
        'Fecha_Camara', 'ScreenPath', 'LinkToClic'])
    return df

# Clase para realizar las solicitudes
class Empresa(Resource):
    

    def get(self):
        # Define data de entrada por medio de los parametros pasados en el request
        input_data = {}
        input_data['busqueda'] = request.args.get('busqueda')
        input_data['departamento'] = request.args.get('departamento')
        input_data['n'] = request.args.get('n')

        # Crea el esquema y le carga la data a validar
        schema = EmpresaSchema()

        try:
            empresa = schema.load(input_data)
            #robot --variable busqueda:"Santa Fe" --variable departamento:11 --variable n:14 einforma.robot
            query = 'robot --variable busqueda:"{q}" --variable departamento:{d} --variable n:{n} einforma.robot'
            query = query.format(q = empresa.get('busqueda'), d = Departamento_id, n = empresa.get('n'))
            print('query', query)
            
            output = subprocess.getoutput(query)
            print(output)

            df_result = read_file()
            info_dict = {}

            for i in range(len(df_result)):
                info_dict[df_result['Razon_Social'][i]] = {
                    'Informacion_Básica': {
                        'Forma_Juridica': df_result['Forma_Juridica'][i],
                        'Departamento': df_result['Departamento'][i],
                        'Actividad_CIIU': df_result['Actividad_CIIU'][i],
                        'Fecha_Constitucion': df_result['Fecha_Constitucion'][i]
                    },
                    'Actualizaciones_importantes' : {
                        'Fecha_Ultimo_Dato': df_result['Fecha_Ultimo_Dato'][i],
                        'Fecha_Camara': df_result['Fecha_Camara'][i]
                    },
                    'Archivos': {
                        'Screenshot_Path': df_result['ScreenPath'][i],
                        'Link_Empresa': df_result['LinkToClic'][i]
                    }
                }
            
            if len(df_result) == 0:
                info_dict['Respuesta'] = {
                    'Codigo': 902,
                    'Mensaje': 'No existen datos de retorno'
                }

            data = {
                'Busqueda': {
                    'String Busqueda': empresa.get('busqueda'),
                    'Id Departamento': Departamento_id,
                    'Departamento': empresa.get('departamento'),
                    'Numero Empresas': empresa.get('n')
                },
                'Informacion': info_dict
            }

            return data

        except Exception as e:
            return{
                'Codigo': 901,
                'Mensaje': 'Error en los datos: ' + str(e)
            }



# Añade el recurso Empresa a la app
api.add_resource(Empresa, '/Empresa')

# Funcion Main
if __name__ == '__main__':
    app.run(debug=True)