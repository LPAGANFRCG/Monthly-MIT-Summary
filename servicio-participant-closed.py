import smartsheet

# ⚠️ Usa un token seguro
access_token = 'TU_NUEVO_API_TOKEN'

# IDs
sheet_id_origen = 7263899889231748
sheet_id_destino = 8251269215899524
columna_estado = 'Warranty Claim Status'

# Inicializar cliente
smartsheet_client = smartsheet.Smartsheet(access_token)

# Obtener sheet origen
sheet = smartsheet_client.Sheets.get_sheet(sheet_id_origen)

# Obtener ID de columna 'Warranty Claim Status'
col_map = {col.title: col.id for col in sheet.columns}
estado_col_id = col_map[columna_estado]

# Buscar filas con estado 'Closed'
filas_para_mover = []
for row in sheet.rows:
    for cell in row.cells:
        if cell.column_id == estado_col_id and cell.value == 'Closed':
            filas_para_mover.append(row)

# Procesar filas
for row in filas_para_mover:
    # Copiar fila al sheet destino
    copy_result = smartsheet_client.Sheets.copy_row(
        sheet_id_origen,
        [row.id],
        smartsheet.models.CopyOrMoveRowDirective(
            destination_sheet_id=sheet_id_destino
        )
    )

    # Obtener attachments
    attachments = smartsheet_client.Attachments.list_row_attachments(sheet_id_origen, row.id).data

    # Obtener nueva fila ID
    nueva_fila_id = copy_result.copy_or_move_row_result.row_mappings[0].new_row_id

    # Copiar attachments
    for att in attachments:
        contenido = smartsheet_client.Attachments.download_attachment(sheet_id_origen, att.id)
        smartsheet_client.Attachments.attach_file_to_row(
            sheet_id_destino,
            nueva_fila_id,
            (att.name, contenido.raw, att.mime_type)
        )
