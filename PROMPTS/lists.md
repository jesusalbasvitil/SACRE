1. ID BIRDS
const birdNamesSpanishToLatinAndDistance = {
  // Diccionario de especies json
const speciesMap = {<!-- [BIRDS]
const especiesDiccionario = {
  // Aves
  "pardillo": "Linaria cannabina",
  "rabilarga": "Sylvia undata",
  "alondra": "Alauda arvensis",
  "triguero": "Emberiza calandra",
  "ruiseñor": "Luscinia megarhynchos",
  "ruiseñor común": "Luscinia megarhynchos",
  "verdecillo": "Serinus serinus",
  "torcaz": "Columba palumbus",
  "paloma torcaz": "Columba palumbus",
  "azor": "Accipiter gentilis",
  "verderón": "Chloris chloris",
  "petirrojo": "Erithacus rubecula",
  "peti": "Erithacus rubecula",
  "herrerillo": "Cyanistes caeruleus",
  "curruca cabecinegra": "Sylvia melanocephala",
  "cabecinegra": "Sylvia melanocephala",
  "carrasqueña": "Sylvia cantillans",
  "curruca carrasqueña": "Sylvia cantillans",
  "capirotada": "Sylvia atricapilla"
  "jilguero": "Carduelis carduelis",
  "abejaruco": "Merops apiaster",
  "cogujada común": "Galerida cristata",
  "golondrina": "Hirundo rustica",
  "chova piquirroja": "Pyrrhocorax pyrrhocorax",
  "piquirroja": "Pyrrhocorax pyrrhocorax"
  "gorrión": "Passer domesticus",
  "gorrión común": "Passer domesticus",
  "gorrión moruno": "Passer hispanolienis"
  "chillón": "Petronia petronia",
  "lavandera blanca": "Motacilla alba",
  "turca": "Streptopelia decaocto",
  "tórtola común": "Streptopelia turtur",
  "estornino": "Sturnus vulgaris",
  "escribano soteño": "Emberiza cirlus",
  "avión común": "Delichon urbicum",
  "mirlo": "Turdus merula",
  "pinzón": "Fringilla coelebs",
  "mito": "Aegithalos caudatus",
  "carbonero": "Parus major",
  "carbonero común": "Parus major",
  "totovía": "Lullula arborea",
  "picido": "Picidae sp.",
  "arrendajo": "Garrulus glandarius",
  "reyezuelo": "Regulus sp.",
  "mosquitero papialbo": "Phylloscopus bonelli",
  "papialbo": "Phylloscopus bonelli",
  "ruiseñor bastardo": "Cettia cetti",
  "escribano": "Emberiza sp.",
  "tarabilla norteña": "Saxixola rubetra",
  "norteña": "Saxixola rubetra",
  "piquituerto": "Loxia curvirostra"; 
  "zorzal": "Turdus philomelos",
  "zorzal común": "Turdus philomelos",
  "zorzal charlo": "Turdus viscivorus",
  "cuco": "Cuculus canorus",
  "acentor": "Prunella modularis",
  "chochín": "Troglodytes troglodytes",
  "cernícalo": "Falco tinnunculus",
  "picapinos": "Dendrocopos major",
  "avión roquero": "Ptyonoprogne rupestris",
  "pito ibérico":  "Picus sharpei",
  "pito": "Picus sharpei",
  "buitre": "Gyps fulvus",
  "garrapinos": "Periparus ater",
  "mosquitero": "Phylloscopus collybita",
  "Emberiza": "Emberiza cia"
  
};

// Función para reemplazar los nombres comunes con nombre común + nombre científico
function agregarNombreCientifico(texto) {
  // Lista de prefijos numéricos que pueden aparecer antes del nombre
  const prefijos = ["- ", "- 1 ", "- 2 ", "- 3 ", "- 4 ", "- 5 ", "- 6 ", "- 3-4 ", "- 4-5 "];
  
  // Para cada especie en nuestro diccionario
  for (const [nombreComun, nombreCientifico] of Object.entries(especiesDiccionario)) {
    // Crear regex para buscar el patrón: prefijo + nombreComun (con potencial s al final para plurales)
    prefijos.forEach(prefijo => {
      // Patrón para el nombre exacto
      const regexExacto = new RegExp(`${prefijo}${nombreComun}(\\b|s\\b)`, 'gi');
      
      // Reemplazar con el nombre común original + nombre científico en cursiva entre paréntesis
      texto = texto.replace(regexExacto, (match) => {
        return `${match} (*${nombreCientifico}*)`;
      });
    });
  }
  
  return texto;
}

// Texto de ejemplo para verificar cómo funciona nuestra función
const ejemploTexto = "- 3 pardillos\n- 2 rabilargas\n- 1 alondra\n- 2 trigueros\n- 1 ruiseñor";
console.log("Ejemplo de transformación:");
console.log(agregarNombreCientifico(ejemploTexto));
Salida-->




2. For every listen station (1 to 20) make a table with title by listen station centered, above the next rows: "Especie", latin name; above "<>25 m" assign number of specimens each specie, preferible >25; all of raptors or big birds to >25. Time -"Hora de Inicio"- extract every audio file like this: "..at 13.08.51.ogg" 

|   **Especie**     |**< 25 m**|**> 25 m** |
|-------------------|----------|-----------|
|                   |          |           |
|                   |          |           |
|                   |          |           |
|                   |          |           |

**Hora de Inicio:** _______


 3. I have 20 watch stations (from "Olalla1" or much better, Ol1" to Ol20) but I haven't birds in all of us. Take the doc added an follow the instructions: 
 - 1st order the obs from 1 to 20, make a list with latin name, number of specimens each specie. 
 - 2nd if not found species for a station empty, use the species of 2 listen stations before and 2 listen stations after this empty station and put with some random. 

 4. [[SACRE'26][Olalla ..0531]](https://share.gemini.google/7YPVUJWZ2kzG)Convert the table to excel file like this example of code:

 import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Create workbook
wb = openpyxl.Workbook()

# Setup sheets
ws_summary = wb.active
ws_summary.title = "Resumen"
ws_data = wb.create_sheet(title="Datos Consolidados")

# Data structure
raw_data = {
    "Ol1": {
        "time": "08.44.33.ogg",
        "birds": [
            ("Serinus serinus", 1, 0), ("Fringilla coelebs", 1, 0), ("Columba palumbus", 0, 2),
            ("Cettia cetti", 1, 0), ("Luscinia megarhynchos", 1, 0), ("Petronia petronia", 1, 0),
            ("Monticola sp.", 3, 0), ("Alauda arvensis", 2, 0), ("Sylvia atricapilla", 4, 0)
        ]
    },
    "Ol2": {
        "time": "08.48.00 (Interpolada)",
        "birds": [
            ("Monticola sp.", 2, 0), ("Alauda arvensis", 3, 0), ("Sylvia atricapilla", 2, 0),
            ("Linaria cannabina", 4, 0), ("Petronia petronia", 2, 0)
        ]
    },
    "Ol3": {
        "time": "08.52.24.ogg",
        "birds": [
            ("Hippolais polyglotta", 1, 0), ("Linaria cannabina", 7, 0), ("Petronia petronia", 4, 0),
            ("Sylvia sp.", 1, 0), ("Cuculus canorus", 1, 0), ("Sylvia undata", 1, 0),
            ("Ptyonoprogne rupestris", 1, 0), ("Upupa epops", 1, 0)
        ]
    },
    "Ol4": {
        "time": "08.58.30 (Interpolada)",
        "birds": [
            ("Linaria cannabina", 3, 0), ("Petronia petronia", 3, 0), ("Upupa epops", 1, 0),
            ("Jynx torquilla", 1, 0), ("Turdus merula", 1, 0)
        ]
    },
    "Ol5": {
        "time": "09.04.56.ogg",
        "birds": [
            ("Jynx torquilla", 1, 0), ("Petronia petronia", 5, 0), ("Turdus merula", 1, 0),
            ("Cettia cetti", 1, 0), ("Luscinia megarhynchos", 1, 0), ("Columba palumbus", 0, 4),
            ("Carduelis carduelis", 1, 0), ("Emberiza cia", 1, 0), ("Galerida sp.", 1, 0),
            ("Sylvia undata", 1, 0), ("Oriolus oriolus", 1, 0), ("Ptyonoprogne rupestris", 1, 0)
        ]
    },
    "Ol6": {
        "time": "09.22.53.ogg",
        "birds": [
            ("Phoenicurus ochruros", 1, 0), ("Monticola saxatilis", 1, 0), ("Alauda arvensis", 2, 0),
            ("Petronia petronia", 3, 0), ("Sturnus unicolor", 4, 0), ("Emberiza cia", 1, 0),
            ("Serinus serinus", 1, 0), ("Carduelis carduelis", 3, 0), ("Galerida sp.", 1, 0),
            ("Columba palumbus", 0, 1)
        ]
    },
    "Ol7": {
        "time": "09.25.31.ogg",
        "birds": [
            ("Emberiza calandra", 2, 0), ("Melanocorypha calandra", 1, 0)
        ]
    },
    "Ol8": {
        "time": "09.32.48.ogg",
        "birds": [
            ("Alauda arvensis", 7, 0), ("Coturnix coturnix", 2, 0), ("Emberiza calandra", 1, 0),
            ("Melanocorypha calandra", 1, 0), ("Columba palumbus", 0, 2)
        ]
    },
    "Ol9": {
        "time": "09.36.04.ogg",
        "birds": [
            ("Emberiza calandra", 2, 0), ("Coturnix coturnix", 1, 0), ("Alauda arvensis", 2, 0),
            ("Sturnus unicolor", 1, 0), ("Galerida sp.", 1, 0)
        ]
    },
    "Ol10": {
        "time": "09.39.31.ogg",
        "birds": [
            ("Emberiza calandra", 2, 0), ("Alauda arvensis", 3, 0), ("Coturnix coturnix", 1, 0),
            ("Galerida sp.", 1, 0), ("Pica pica", 0, 1), ("Cuculus canorus", 1, 0),
            ("Circus pygargus", 0, 1), ("Milvus migrans", 0, 1)
        ]
    },
    "Ol11": {
        "time": "09.44.49.ogg",
        "birds": [
            ("Emberiza calandra", 1, 0), ("Alauda arvensis", 4, 0), ("Galerida sp.", 1, 0),
            ("Coturnix coturnix", 1, 0), ("Corvus corone", 0, 1), ("Turdus merula", 2, 0)
        ]
    },
    "Ol12": {
        "time": "09.48.41.ogg",
        "birds": [
            ("Turdus merula", 3, 0), ("Cettia cetti", 1, 0), ("Luscinia megarhynchos", 1, 0),
            ("Carduelis carduelis", 1, 0), ("Fringilla coelebs", 1, 0), ("Columba palumbus", 0, 2),
            ("Sylvia atricapilla", 1, 0), ("Cuculus canorus", 1, 0), ("Picus sharpei", 1, 0),
            ("Troglodytes troglodytes", 1, 0)
        ]
    },
    "Ol13": {
        "time": "09.52.13.ogg",
        "birds": [
            ("Emberiza calandra", 2, 0), ("Alauda arvensis", 3, 0), ("Galerida sp.", 1, 0),
            ("Coturnix coturnix", 1, 0)
        ]
    },
    "Ol14": {
        "time": "09.56.53.ogg",
        "birds": [
            ("Alauda arvensis", 4, 0), ("Emberiza calandra", 2, 0), ("Galerida sp.", 1, 0),
            ("Coturnix coturnix", 1, 0), ("Corvus corone", 0, 1)
        ]
    },
    "Ol15": {
        "time": "10.02.25.ogg",
        "birds": [
            ("Alauda arvensis", 3, 0), ("Emberiza calandra", 1, 0), ("Coturnix coturnix", 1, 0),
            ("Sturnus unicolor", 1, 0), ("Circus pygargus", 0, 1), ("Upupa epops", 1, 0)
        ]
    },
    "Ol16": {
        "time": "10.06.08.ogg",
        "birds": [
            ("Coturnix coturnix", 1, 0), ("Alauda arvensis", 3, 0), ("Emberiza calandra", 1, 0),
            ("Cisticola juncidis", 1, 0), ("Pica pica", 0, 1), ("Milvus migrans", 0, 1)
        ]
    },
    "Ol17": {
        "time": "10.16.31.ogg",
        "birds": [
            ("Turdus merula", 3, 0), ("Cettia cetti", 1, 0), ("Luscinia megarhynchos", 1, 0),
            ("Fringilla coelebs", 1, 0), ("Carduelis carduelis", 2, 0), ("Columba palumbus", 0, 2),
            ("Sylvia atricapilla", 1, 0), ("Cuculus canorus", 1, 0), ("Picus sharpei", 1, 0),
            ("Upupa epops", 1, 0), ("Troglodytes troglodytes", 1, 0), ("Sturnus unicolor", 1, 0)
        ]
    },
    "Ol18": {
        "time": "10.20.07.ogg",
        "birds": [
            ("Emberiza calandra", 4, 0), ("Alauda arvensis", 6, 0), ("Galerida sp.", 2, 0),
            ("Coturnix coturnix", 1, 1), ("Falco naumanni", 0, 1)
        ]
    },
    "Ol19": {
        "time": "10.25.07.ogg",
        "birds": [
            ("Alauda arvensis", 8, 0), ("Emberiza calandra", 4, 0), ("Galerida sp.", 2, 0),
            ("Coturnix coturnix", 2, 0), ("Corvus corone", 0, 1), ("Sturnus unicolor", 1, 0),
            ("Pica pica", 0, 1)
        ]
    },
    "Ol20": {
        "time": "10.27.34.ogg",
        "birds": [
            ("Alauda arvensis", 14, 0), ("Emberiza calandra", 6, 0), ("Coturnix coturnix", 2, 1),
            ("Circus pygargus", 0, 2), ("Galerida sp.", 3, 0), ("Corvus corone", 0, 1),
            ("Pica pica", 0, 1)
        ]
    }
}

# Design Palettes
accent_fill = PatternFill(start_color="1B4D3E", end_color="1B4D3E", fill_type="solid") # Forest Green
zebra_fill = PatternFill(start_color="F4F7F5", end_color="F4F7F5", fill_type="solid")
header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
title_font = Font(name="Segoe UI", size=16, bold=True, color="1B4D3E")
regular_font = Font(name="Segoe UI", size=11)
italic_font = Font(name="Segoe UI", size=11, italic=True)
bold_font = Font(name="Segoe UI", size=11, bold=True)

thin_border = Border(
    left=Side(style='thin', color='DDDDDD'),
    right=Side(style='thin', color='DDDDDD'),
    top=Side(style='thin', color='DDDDDD'),
    bottom=Side(style='thin', color='DDDDDD')
)

# Build Tab 1: Consolidated Flat Data List
headers_data = ["Estación", "Hora/Archivo", "Especie (Nombre Científico)", "Cantidad < 25m", "Cantidad > 25m", "Total"]
ws_data.append(headers_data)
for cell in ws_data[1]:
    cell.fill = accent_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

row_idx = 2
for stat, info in raw_data.items():
    for bird, c1, c2 in info["birds"]:
        ws_data.cell(row=row_idx, column=1, value=stat).alignment = Alignment(horizontal="center")
        ws_data.cell(row=row_idx, column=2, value=info["time"])
        b_cell = ws_data.cell(row=row_idx, column=3, value=bird)
        b_cell.font = italic_font
        ws_data.cell(row=row_idx, column=4, value=c1).number_format = '#,##0'
        ws_data.cell(row=row_idx, column=5, value=c2).number_format = '#,##0'
        ws_data.cell(row=row_idx, column=6, value=f"=D{row_idx}+E{row_idx}").number_format = '#,##0'
        
        # Apply style to data row
        for col in range(1, 7):
            c = ws_data.cell(row=row_idx, column=col)
            c.font = italic_font if col == 3 else regular_font
            c.border = thin_border
            if row_idx % 2 == 1:
                c.fill = zebra_fill
        row_idx += 1

ws_data.views.sheetView[0].showGridLines = True

# Build Tab 0: Beautiful Styled Summary Table Layouts
ws_summary.cell(row=2, column=2, value="Muestreo de Aves - Programa SACRE '26 (Olalla)").font = title_font

curr_row = 4
for stat, info in raw_data.items():
    ws_summary.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
    title_cell = ws_summary.cell(row=curr_row, column=2, value=f"Estación de Escucha: {stat}")
    title_cell.font = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    title_cell.fill = accent_fill
    title_cell.alignment = Alignment(horizontal="center")
    
    curr_row += 1
    ws_summary.cell(row=curr_row, column=2, value="Especie").font = bold_font
    ws_summary.cell(row=curr_row, column=3, value="< 25 m").font = bold_font
    ws_summary.cell(row=curr_row, column=4, value="> 25 m").font = bold_font
    for col in range(2, 5):
        ws_summary.cell(row=curr_row, column=col).border = thin_border
        ws_summary.cell(row=curr_row, column=col).alignment = Alignment(horizontal="center" if col > 2 else "left")
        
    start_birds_row = curr_row + 1
    curr_row += 1
    
    for bird, c1, c2 in info["birds"]:
        b_c = ws_summary.cell(row=curr_row, column=2, value=bird)
        b_c.font = italic_font
        ws_summary.cell(row=curr_row, column=3, value=c1 if c1 > 0 else "").number_format = '#,##0'
        ws_summary.cell(row=curr_row, column=4, value=c2 if c2 > 0 else "").number_format = '#,##0'
        for col in range(2, 5):
            ws_summary.cell(row=curr_row, column=col).border = thin_border
            if col > 2:
                ws_summary.cell(row=curr_row, column=col).alignment = Alignment(horizontal="center")
        curr_row += 1
        
    # Add Total Row
    ws_summary.cell(row=curr_row, column=2, value="Total Individuos").font = bold_font
    ws_summary.cell(row=curr_row, column=3, value=f"=SUM(C{start_birds_row}:C{curr_row-1})").font = bold_font
    ws_summary.cell(row=curr_row, column=4, value=f"=SUM(D{start_birds_row}:D{curr_row-1})").font = bold_font
    for col in range(2, 5):
        ws_summary.cell(row=curr_row, column=col).border = thin_border
        if col > 2:
            ws_summary.cell(row=curr_row, column=col).alignment = Alignment(horizontal="center")
            
    curr_row += 1
    ws_summary.cell(row=curr_row, column=2, value=f"Hora de Inicio / Archivo: {info['time']}").font = Font(name="Segoe UI", size=10, italic=True)
    curr_row += 2 # Spacing between stations

ws_summary.views.sheetView[0].showGridLines = True

# Autofit column widths
for ws in [ws_summary, ws_data]:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

# Save Workbook
filename = "SACRE26_Olalla_Resultados.xlsx"
wb.save(filename)
print(f"Workbook successfully saved as {filename}")