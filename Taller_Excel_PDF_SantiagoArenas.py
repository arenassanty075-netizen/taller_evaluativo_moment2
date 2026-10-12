import openpyxl

from tabulate import tabulate
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors


#1. Carga el libro del trabajo existente
wb = openpyxl.load_workbook('ventas_tecnologia_v2.xlsx', data_only=True)

#2. Selecciona la hoja activa (primera pestaña)
hoja = wb.active

print(f"Hoja cargada exitosamente: {hoja.title}") # muestra el nombre de la hoja (ventas)



#3. Crear un lista vacia para almacenar las filas
datos = []

#4.Recorre desde la fila 2 (omitimos encabezados de la fila 1)
for fila in hoja.iter_rows(min_row=2 , values_only=True):
    datos.append(list(fila))
print(f"Total de registros leídos en memoria: {len(datos)}") #ignora el encabezado da el total de datos

#5. Nombre de columnas
encabezados = ["ID" ,"Producto" ,"Categoría" ,"Precio ($)" ,"Ventas (Unid)"]

#6. Aliniacion por columna ("left", "center", "right")
aliniacion = ("center" ,"left" ,"center", "right" ,"right")
#7. Mostrar tabla formateada por consola
def ver_catalogo_completo():
    print("\n=== REPORTE DE VENTAS EN CONSOLA (22 PRODUCTOS) ===")
    print(tabulate(datos, headers=encabezados, tablefmt="fancy_grid", colalign=aliniacion))



# Reto 1: filtrar productos con precio superior a $100

def filtrar_por_precio():
    datos_filtrados = []

    for fila in datos:
        if fila[3] > 100:
            datos_filtrados.append(fila)

    print("\n=== PRODUCTOS CON PRECIO SUPERIOR A $100 ===")

    print(tabulate(
        datos_filtrados,
        headers=encabezados,
        tablefmt="fancy_grid",
        colalign=aliniacion
    ))




# Reto 2: calcular métricas
def calcular_metricas():
    total_unidades = 0
    suma_precios = 0

    for fila in datos:
        total_unidades += fila[4]
        suma_precios += fila[3]

    if len(datos) > 0:
        precio_promedio = suma_precios / len(datos)
    else:
        precio_promedio = 0

    print("\n=== RESUMEN DE VENTAS ===")
    print(f"Total de unidades vendidas: {total_unidades}")
    print(f"Precio promedio: ${precio_promedio:,.2f}")




# Reto 3: buscar producto por ID
def buscar_producto_Id():
    id_buscar = input("\nIngrese el ID del producto que desea buscar: ").strip()

    producto_encontrado = None

    for fila in datos:
        if str(fila[0]) == id_buscar:
            producto_encontrado = fila
            break

    if producto_encontrado is not None:
        print("\n=== PRODUCTO ENCONTRADO ===")
        print(tabulate(
            [producto_encontrado],
            headers=encabezados,
            tablefmt="fancy_grid",
            colalign=aliniacion
        ))
    else:
        print("Advertencia: no existe un producto con ese ID.")


#8. Crear documento base
def generar_pdf():
    

    # Filtrar productos con precio superior a $100
    datos_filtrados = []

    for fila in datos:
        if fila[3] > 100:
            datos_filtrados.append(fila)

    # Calcular métricas generales
    total_unidades = 0
    suma_precios = 0

    for fila in datos:
        total_unidades += fila[4]
        suma_precios += fila[3]

    
    if len(datos) > 0:
        precio_promedio = suma_precios / len(datos)
    else:
        precio_promedio = 0

    # Crear documento PDF
    doc = SimpleDocTemplate(
        "reporte_ejecutivo_ventas.pdf",
        pagesize=letter
    )

    estilos = getSampleStyleSheet()

#9.Eestilos perzonalizados
    estilo_titulo = ParagraphStyle(
        'Estilotitulo', parent=estilos['Title'],
        fontName = 'Helvetica-Bold', fontSize=18,
        textColor=colors.HexColor("#1A365D"), spaceAfter=10
    )

    estilo_cuerpo = ParagraphStyle('EstiloCuerpo', parent=estilos['Normal'], fontSize=9)
    estilo_encabezado = ParagraphStyle(
    'EstiloEncabezado', parent=estilos["Normal"],
        fontName='Helvetica-Bold', fontSize=10,
        textColor=colors.whitesmoke, alignment=1
    )

#10. lista story(carrito de compras)
    story=[
        Paragraph("Informe Ejecutivo de Ventas Tecno", estilo_titulo),
        Paragraph(f"<b>Total de Productos Procesados:</b> {len(datos)}", estilo_cuerpo),
        Paragraph(f"<b>Total de unidades vendidas:</b> {total_unidades}", estilo_cuerpo),
        Paragraph(f"<b>Precio promedio:</b> ${precio_promedio:,.2f}", estilo_cuerpo),
        Spacer(1, 12)
    ]

#11. Convertir celdas a Paragraph
    encabezados_pdf = [Paragraph(f"<b>{h}</b>", estilo_encabezado) for h in encabezados]
    filas_pdf = [
        [
            Paragraph(str(r[0]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(str(r[1]), ParagraphStyle('L', parent=estilo_cuerpo, alignment=0)),
            Paragraph(str(r[2]), ParagraphStyle('C', parent=estilo_cuerpo, alignment=1)),
            Paragraph(f"${r[3]:,.2f}", ParagraphStyle('R', parent=estilo_cuerpo, alignment=2)),
            Paragraph(str(r[4]), ParagraphStyle('R', parent=estilo_cuerpo, alignment=2))
        ]
        for r in datos_filtrados
    ]

#12. Configurar anchos y tablas
    anchos_cols =[40, 160, 100, 80, 80]
    tabla_pdf = Table([encabezados_pdf] + filas_pdf, colWidths=anchos_cols, repeatRows=1)

#13. Aplicar estilos viisual con TableStyle
    tabla_pdf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1A365D")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E0")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7FAFC")]),
    ]))

    story.append(tabla_pdf)

#14. Compilar PDF
    doc.build(story)
    print("PDF generado correctamente: reporte_ejecutivo_ventas.pdf")


# RETO 4: MENU INTERACTIVO
def mostrar_menu():
    while True:
        print("\n================================")
        print("       TECNOVENTAS")
        print("================================")
        print("1. Ver catalogo completo")
        print("2. Filtrar productos por precio > $100")
        print("3. Buscar producto por ID")
        print("4. Exportar informe PDF")
        print("5. Salir")
        print("================================")

        opcion = input("Seleccione una opcion: ").strip()

        if opcion == "1":
            ver_catalogo_completo()
            calcular_metricas()

        elif opcion == "2":
            filtrar_por_precio()

        elif opcion == "3":
            buscar_producto_Id()
            

        elif opcion == "4":
            generar_pdf()

        elif opcion == "5":
            print("Gracias por utilizar TecnoVentas.")
            break

        else:
            print("Opcion no valida. Intente nuevamente.")


# INICIAR PROGRAMA
mostrar_menu()
