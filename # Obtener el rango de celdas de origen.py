   # Obtener el rango de celdas de origen
    rango_origen = hoja['AWA67:AWA73']  # Columna anterior

    # Obtener el rango de celdas de destino
    rango_destino = hoja['AWB67:AWB73']  # Columna actual

    # Recorrer el rango de celdas de origen y adaptar la fórmula a las celdas de destino
    for celda_origen, celda_destino in zip(rango_origen, rango_destino):
        print(celda_origen)
        print(celda_destino)
        formula = celda_origen[0].value
        print(formula)
        if formula and formula.startswith('='):
            offset = celda_origen[0].column - celda_destino[0].column
            print(offset)
            nueva_formula = formula.replace(celda_origen[0].column_letter, get_column_letter(celda_origen[0].column + 1))
            print(nueva_formula)
            celda_destino[0].value = nueva_formula

    # Recalcular las fórmulas en la hoja de trabajo
    hoja.calculate_dimension()

    # Guardar el archivo modificado
    libro.save(archivo_salida)