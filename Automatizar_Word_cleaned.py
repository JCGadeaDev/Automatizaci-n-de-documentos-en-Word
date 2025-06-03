import pandas as pd
from datetime import datetime
from docxtpl import DocxTemplate
import openpyxl


def main():
    # Cargar plantilla
    doc = DocxTemplate("plantilla.docx")

    # Datos constantes
    nombre = "Juan Carlos Gadea"
    telefono = "(39)344 852 9898"
    correo = "gadeanova@hotmail.com"
    fecha = datetime.today().strftime("%d/%m/%Y")
    constantes = {'nombre': nombre, 'telefono': telefono, 'correo': correo, 'fecha': fecha}

    print(f"Fecha actual: {fecha}")
    print("Constantes:", constantes)

    # Leer Excel
    df = pd.read_excel('Alumnos.xlsx')
    df.columns = df.columns.str.strip()  # quitar espacios extra en los nombres de columnas

    # Procesar cada fila
    for indice, fila in df.iterrows():
        contenido = {
            'nombre_alumno': fila['Nombre del Alumno'],
            'nota_mat': fila['Mat'],
            'nota_fis': fila['Fis'],
            'nota_qui': fila['Qui']
        }
        contenido.update(constantes)

        doc.render(contenido)
        output_filename = f"Notas_de_{fila['Nombre del Alumno']}.docx"
        doc.save(output_filename)

        print(f"Documento generado: {output_filename}")
        print("Contenido:", contenido)

    # Guardar documento de prueba (opcional)
    doc.render(constantes)
    doc.save("prueba.docx")
    print("Documento de prueba generado: prueba.docx")


if __name__ == "__main__":
    main()

