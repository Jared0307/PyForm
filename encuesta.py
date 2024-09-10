from http.server import BaseHTTPRequestHandler, HTTPServer
import urllib.parse
import pandas as pd
import os

# Define el nombre del archivo Excel
EXCEL_FILE = 'respuestas_encuesta.xlsx'

# Función para guardar los datos en el archivo Excel
def guardar_respuestas(datos):
    # Verifica si el archivo ya existe
    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE)
    else:
        # Si no existe, crea un nuevo DataFrame con las columnas especificadas
        df = pd.DataFrame(columns=[
            'Nombre', 'Ubicación', 'Puesto', 'Departamento', 'Usuario', 'Contraseña',
            'Acrobat', 'Office', 'AutoCAD', 'Comentarios', 'Equipo',
            'Marca', 'Modelo', 'N° Serie', 'S.O.', 'Licencia Win 10', 'RAM',
            'Procesador', 'MAC Ethernet', 'MAC WIFI', 'DD', 'Nombre del dispositivo', 'Observaciones'
        ])
    
    # Agrega los nuevos datos al DataFrame
    df = pd.concat([df, pd.DataFrame([datos])], ignore_index=True)
    
    # Guarda el DataFrame en el archivo Excel
    df.to_excel(EXCEL_FILE, index=False)

class EncuestaHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/':
            self.enviar_formulario()
        elif self.path.endswith('.css'):
            self.enviar_archivo(self.path, 'text/css')
        elif self.path.endswith('.js'):
            self.enviar_archivo(self.path, 'application/javascript')
        else:
            self.enviar_error()

    def do_POST(self):
        if self.path == '/submit':
            self.process_form()
        else:
            self.enviar_error()

    def enviar_formulario(self):
        formulario = '''
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ENCUESTA TI</title>
    <link rel="stylesheet" href="styles.css">
    <script src="script.js" defer></script>
</head>
<body>
    <h1>DATOS</h1>
    <form method="post" action="/submit">
        <label for="nombre">Nombre:</label>
        <input type="text" id="nombre" name="nombre">
        
        <label for="ubicacion">Ubicación:</label>
        <input type="text" id="ubicacion" name="ubicacion">
        
        <label for="puesto">Puesto:</label>
        <input type="text" id="puesto" name="puesto">
        
        <label for="departamento">Departamento:</label>
        <input type="text" id="departamento" name="departamento">
        
        <label for="usuario">Usuario:</label>
        <input type="text" id="usuario" name="usuario">
        
        <label for="contraseña">Contraseña:</label>
        <input type="text" id="contraseña" name="contraseña">
        
        <label for="acrobat">Acrobat:</label>
        <input type="text" id="acrobat" name="acrobat">
        
        <label for="office">Office:</label>
        <input type="text" id="office" name="office">
        
        <label for="autocad">AutoCAD:</label>
        <input type="text" id="autocad" name="autocad">
        
        <label for="comentarios">Comentarios:</label>
        <input type="text" id="comentarios" name="comentarios">
        
        <label for="equipo">Equipo:</label>
        <input type="text" id="equipo" name="equipo">
        
        <label for="marca">Marca:</label>
        <a>wmic computersystem get manufacturer</a>
        <input type="text" id="marca" name="marca">
        
        <label for="modelo">Modelo:</label>
        <a>Get-WmiObject -Class Win32_ComputerSystemProduct</a>
        <input type="text" id="modelo" name="modelo">
        
        <label for="numero_serie">N° Serie:</label>
        <a>wmic bios get serialnumber</a>
        <input type="text" id="numero_serie" name="numero_serie">
        
        <label for="so">Sistema Operativo:</label>
        <a>wmic os get caption</a>
        <input type="text" id="so" name="so">
        
        <label for="licencia_win10">Licencia Win 10:</label>
        <a>wmic path softwarelicensingservice get OA3xOriginalProductKey</a>
        <input type="text" id="licencia_win10" name="licencia_win10">
        
        <label for="ram">RAM:</label>
        <a>wmic memorychip get capacity</a>
        <input type="text" id="ram" name="ram">
        
        <label for="procesador">Procesador:</label>
        <a>wmic cpu get name</a>
        <input type="text" id="procesador" name="procesador">
        
        <label for="mac_ethernet">MAC Ethernet:</label>
        <a>ipconfig /all</a>
        <input type="text" id="mac_ethernet" name="mac_ethernet">
        
        <label for="mac_wifi">MAC WIFI:</label>
        <a>ipconfig /all</a>
        <input type="text" id="mac_wifi" name="mac_wifi">
        
        <label for="dd">Disco Duro:</label>
        <a>Get-PhysicalDisk</a><br></br>
        <a>wmic logicaldisk get size,freespace,caption</a>
        <input type="text" id="dd" name="dd">
        
        <label for="nombre_dispositivo">Nombre del dispositivo:</label>
        <a>hostname</a>
        <input type="text" id="nombre_dispositivo" name="nombre_dispositivo">
        
        <label for="observaciones">Observaciones:</label>
        <textarea id="observaciones" name="observaciones"></textarea>
        
        <center><input type="submit" value="ENVIAR"></center>
    </form>
</body>
</html>
        '''
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(formulario.encode())

    def enviar_archivo(self, path, content_type):
        try:
            with open(path[1:], 'rb') as file:  # Remove leading '/' from path
                self.send_response(200)
                self.send_header('Content-type', content_type)
                self.end_headers()
                self.wfile.write(file.read())
        except FileNotFoundError:
            self.send_error(404, "Archivo no encontrado")

    def process_form(self):
        content_length = int(self.headers['Content-Length'])
        post_data = self.rfile.read(content_length)
        datos = urllib.parse.parse_qs(post_data.decode())

        # Extrae los datos del formulario
        datos_formulario = {
            'Nombre': datos.get('nombre', [''])[0],
            'Ubicación': datos.get('ubicacion', [''])[0],
            'Puesto': datos.get('puesto', [''])[0],
            'Departamento': datos.get('departamento', [''])[0],
            'Usuario': datos.get('usuario', [''])[0],
            'Contraseña': datos.get('contraseña', [''])[0],
            'Acrobat': datos.get('acrobat', [''])[0],
            'Office': datos.get('office', [''])[0],
            'AutoCAD': datos.get('autocad', [''])[0],
            'Comentarios': datos.get('comentarios', [''])[0],
            'Equipo': datos.get('equipo', [''])[0],
            'Marca': datos.get('marca', [''])[0],
            'Modelo': datos.get('modelo', [''])[0],
            'N° Serie': datos.get('numero_serie', [''])[0],
            'S.O.': datos.get('so', [''])[0],
            'Licencia Win 10': datos.get('licencia_win10', [''])[0],
            'RAM': datos.get('ram', [''])[0],
            'Procesador': datos.get('procesador', [''])[0],
            'MAC Ethernet': datos.get('mac_ethernet', [''])[0],
            'MAC WIFI': datos.get('mac_wifi', [''])[0],
            'DD': datos.get('dd', [''])[0],
            'Nombre del dispositivo': datos.get('nombre_dispositivo', [''])[0],
            'Observaciones': datos.get('observaciones', [''])[0],
        }

        # Guarda los datos en el archivo Excel
        guardar_respuestas(datos_formulario)

        # Envia la confirmación de recepción
        respuesta = '''
        <html>
        <head><title>Gracias</title></head>
        <body>
            <h1>Gracias por tu respuesta</h1>
            <a href="/">Volver a la encuesta</a>
        </body>
        </html>
        '''
        self.send_response(200)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(respuesta.encode())

    def enviar_error(self):
        self.send_response(404)
        self.send_header('Content-type', 'text/html')
        self.end_headers()
        self.wfile.write(b'404 Not Found')

def run(server_class=HTTPServer, handler_class=EncuestaHandler, port=8000):
    server_address = ('', port)
    httpd = server_class(server_address, handler_class)
    print(f'Servidor corriendo en el puerto {port}...')
    httpd.serve_forever()

if __name__ == "__main__":
    run()
