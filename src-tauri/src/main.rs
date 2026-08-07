#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use once_cell::sync::Lazy;
use regex::Regex;
use serde_json::{json, Value};
use std::collections::HashMap;
use std::io::{Read as IoRead, Write};
use std::path::Path;
use std::sync::Mutex;

// ---------------------------------------------------------------------------
// Estado global: ruta del Excel seleccionado
// ---------------------------------------------------------------------------

static SELECTED_PATH: Lazy<Mutex<Option<String>>> = Lazy::new(|| Mutex::new(None));

fn get_selected_path() -> Option<String> {
    SELECTED_PATH.lock().unwrap().clone()
}

fn set_selected_path(p: Option<String>) {
    *SELECTED_PATH.lock().unwrap() = p;
}

fn find_default_excel_path() -> Option<String> {
    if let Ok(exe) = std::env::current_exe() {
        if let Some(dir) = exe.parent() {
            let candidate = dir.join("CCGG PLANTILLA - RECUv45.xlsx");
            if candidate.exists() {
                return Some(candidate.to_string_lossy().to_string());
            }
        }
    }
    None
}

fn require_selected_path() -> Result<String, String> {
    match get_selected_path().or_else(find_default_excel_path) {
        Some(p) => Ok(p),
        None => Err("No hay ningun archivo Excel seleccionado.".to_string()),
    }
}

// ---------------------------------------------------------------------------
// Utilidades de columnas Excel
// ---------------------------------------------------------------------------

fn col_name(idx: usize) -> String {
    let mut name = String::new();
    let mut v = idx + 1;
    while v > 0 {
        let r = (v - 1) % 26;
        name.insert(0, (b'A' + r as u8) as char);
        v = (v - 1) / 26;
    }
    name
}

fn col_index(name: &str) -> usize {
    name.chars()
        .fold(0usize, |acc, c| acc * 26 + (c as usize - 'A' as usize + 1))
        - 1
}

fn normalize_plain(s: &str) -> String {
    s.chars()
        .map(|c| match c {
            'á' | 'à' | 'â' | 'ä' | 'Á' | 'À' | 'Â' | 'Ä' => 'A',
            'é' | 'è' | 'ê' | 'ë' | 'É' | 'È' | 'Ê' | 'Ë' => 'E',
            'í' | 'ì' | 'î' | 'ï' | 'Í' | 'Ì' | 'Î' | 'Ï' => 'I',
            'ó' | 'ò' | 'ô' | 'ö' | 'Ó' | 'Ò' | 'Ô' | 'Ö' => 'O',
            'ú' | 'ù' | 'û' | 'ü' | 'Ú' | 'Ù' | 'Û' | 'Ü' => 'U',
            'ñ' | 'Ñ' => 'N',
            other => other.to_ascii_uppercase(),
        })
        .collect::<String>()
        .trim()
        .to_string()
}

fn parse_decimal(v: &Value) -> f64 {
    match v {
        Value::Number(n) => n.as_f64().unwrap_or(0.0),
        Value::String(s) => s.trim().replace('%', "").replace(',', ".").parse::<f64>().unwrap_or(0.0),
        _ => 0.0,
    }
}

// ---------------------------------------------------------------------------
// Leer XLSX con calamine
// ---------------------------------------------------------------------------

use calamine::{open_workbook_auto, Data, Reader};

fn cell_data_to_value(cell: &Data) -> Value {
    match cell {
        Data::Empty => Value::Null,
        Data::String(s) => Value::String(s.clone()),
        Data::Float(f) => json!(f),
        Data::Int(i) => json!(i),
        Data::Bool(b) => Value::Bool(*b),
        Data::DateTime(f) => json!(f.as_f64()),
        Data::Error(_) => Value::Null,
        _ => Value::Null,
    }
}

fn read_sheet_rows(path: &str, sheet: &str) -> Result<Vec<Vec<Value>>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| e.to_string())?;
    read_sheet_rows_from_wb(&mut wb, sheet)
}

// open_workbook_auto vuelve a leer y parsear TODO el zip (shared strings, styles...)
// cada vez que se llama. Cuando hay que leer varias hojas seguidas (p.ej. sincronizar
// varias hojas de evaluacion), abrir el libro una sola vez y reutilizarlo evita ese
// coste repetido (era el cuello de botella real del guardado de recuperaciones).
fn read_sheet_rows_from_wb(wb: &mut calamine::Sheets<std::io::BufReader<std::fs::File>>, sheet: &str) -> Result<Vec<Vec<Value>>, String> {
    let range = wb
        .worksheet_range(sheet)
        .map_err(|e| format!("Hoja '{sheet}' no encontrada: {e}"))?;

    let (row_start, col_start) = range.start().map(|(r, c)| (r as usize, c as usize)).unwrap_or((0, 0));
    let (row_end, _col_end) = range.end().map(|(r, c)| (r as usize, c as usize)).unwrap_or((0, 0));

    // Build a dense array where result[row_idx] corresponds to Excel row (row_idx + 1).
    // Calamine's range.rows() skips rows that have no cells in the XML, so we must
    // use get_value(r, c) to preserve the Excel row-to-index correspondence.
    let total_rows = if row_end >= row_start { row_end - row_start + 1 } else { 0 };
    let total_cols = range.width();

    // result[i] must correspond to Excel row (i+1), so pad with row_start empty rows first.
    let mut result: Vec<Vec<Value>> = vec![vec![]; row_start];
    for r in 0..total_rows {
        let abs_row = (row_start + r) as u32;
        let row: Vec<Value> = (0..total_cols)
            .map(|c| {
                let abs_col = (col_start + c) as u32;
                range.get_value((abs_row, abs_col))
                    .map(cell_data_to_value)
                    .unwrap_or(Value::Null)
            })
            .collect();
        // Prepend col_start empty cells so column indices match Excel column indices
        let mut full_row = vec![Value::Null; col_start];
        full_row.extend(row);
        result.push(full_row);
    }
    Ok(result)
}

fn sheet_names(path: &str) -> Result<Vec<String>, String> {
    let wb = open_workbook_auto(path).map_err(|e| e.to_string())?;
    Ok(wb.sheet_names().to_vec())
}

fn sheet_names_from_wb(wb: &calamine::Sheets<std::io::BufReader<std::fs::File>>) -> Vec<String> {
    wb.sheet_names().to_vec()
}

fn cell_val_str(v: &Value) -> String {
    match v {
        Value::String(s) => s.clone(),
        Value::Number(n) => {
            if let Some(i) = n.as_i64() { i.to_string() } else { n.as_f64().unwrap_or(0.0).to_string() }
        }
        Value::Bool(b) => b.to_string(),
        _ => String::new(),
    }
}

fn cell_str(rows: &[Vec<Value>], row: usize, col: usize) -> String {
    rows.get(row)
        .and_then(|r| r.get(col))
        .map(|v| cell_val_str(v).trim().to_string())
        .unwrap_or_default()
}

fn cell_f64(rows: &[Vec<Value>], row: usize, col: usize) -> Option<f64> {
    rows.get(row).and_then(|r| r.get(col)).and_then(|v| match v {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => s.replace(',', ".").parse().ok(),
        _ => None,
    })
}

fn normalise_date_for_ui(rows: &[Vec<Value>], row: usize, col: usize) -> String {
    let v = rows.get(row).and_then(|r| r.get(col)).cloned().unwrap_or(Value::Null);
    match &v {
        Value::String(s) if !s.is_empty() => s.clone(),
        Value::Number(n) => {
            if let Some(serial) = n.as_f64() {
                let days = serial as i64 - 25569;
                let secs = days * 86400;
                chrono::DateTime::from_timestamp(secs, 0)
                    .map(|dt: chrono::DateTime<chrono::Utc>| dt.format("%Y-%m-%d").to_string())
                    .unwrap_or_default()
            } else {
                String::new()
            }
        }
        _ => String::new(),
    }
}

// ---------------------------------------------------------------------------
// Buscar filas de secciones en hoja DATOS
// ---------------------------------------------------------------------------

fn find_header_row(rows: &[Vec<Value>], text: &str) -> Option<usize> {
    let target = text.to_uppercase();
    rows.iter().position(|row| {
        [0usize, 1].iter().any(|&col| row.get(col)
            .map(|v| cell_val_str(v).to_uppercase().contains(&target))
            .unwrap_or(false))
    })
}

// Devuelve los índices de fila de inicio de cada bloque "UNIDADES Xª" en col K.
// El Excel ESO tiene 3 bloques: UNIDADES 1ª, UNIDADES 2ª, UNIDADES 3ª.
fn find_unidades_bloques(rows: &[Vec<Value>]) -> Vec<(String, usize)> {
    let mut bloques = Vec::new();
    for (i, row) in rows.iter().enumerate() {
        let k = row.get(10).map(|v| cell_val_str(v).trim().to_string()).unwrap_or_default();
        let ku = k.to_uppercase();
        if ku.contains("UNIDADES") && (ku.contains('1') || ku.contains('2') || ku.contains('3')) {
            let eval = if ku.contains('1') { "1\u{00aa}" }
                       else if ku.contains('2') { "2\u{00aa}" }
                       else { "3\u{00aa}" };
            bloques.push((eval.to_string(), i + 1)); // start = fila siguiente a la cabecera
        }
    }
    bloques
}

// Mantener para compatibilidad con save_unidades_to_file
fn find_unidades_start(rows: &[Vec<Value>]) -> Option<usize> {
    find_unidades_bloques(rows).into_iter().map(|(_, s)| s).next()
}

// No hay sección RRAA en el Excel ESO — no se usa
fn find_rraa_start(_rows: &[Vec<Value>]) -> Option<usize> {
    None
}

// ---------------------------------------------------------------------------
// Comandos: selectFile / getSelectedFile / setSelectedFile
// ---------------------------------------------------------------------------

#[tauri::command]
fn excel_select_file() -> Result<Value, String> {
    let file = rfd::FileDialog::new()
        .set_title("Selecciona la plantilla Excel")
        .add_filter("Excel", &["xlsx", "xlsm", "xls"])
        .pick_file();
    match file {
        Some(path) => {
            let p = path.to_string_lossy().to_string();
            set_selected_path(Some(p.clone()));
            load_alumnos(&p)
        }
        None => Ok(Value::Null),
    }
}

#[tauri::command]
fn excel_get_selected_file() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() {
        Some(p) => load_alumnos(&p),
        None => Ok(Value::Null),
    }
}

#[tauri::command]
fn excel_set_selected_file(file_path: String) -> Result<Value, String> {
    if file_path.is_empty() { return Err("No se especificó ningún archivo.".to_string()); }
    if !Path::new(&file_path).exists() { return Err(format!("El archivo no existe: {file_path}")); }
    let name = Path::new(&file_path).file_name().unwrap_or_default().to_string_lossy().to_string();
    set_selected_path(Some(file_path.clone()));
    Ok(json!({ "filePath": file_path, "fileName": name }))
}

#[tauri::command]
fn excel_verify_file_exists(file_path: String) -> Result<bool, String> {
    Ok(Path::new(&file_path).exists())
}

// ---------------------------------------------------------------------------
// Cargar alumnos
// ---------------------------------------------------------------------------

fn load_alumnos(path: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let header_row = find_header_row(&rows, "ALUMNADO")
        .ok_or("No se encontro la seccion ALUMNADO en la hoja DATOS.")?;
    let mut alumnos = Vec::new();
    for i in (header_row + 1)..rows.len() {
        let nombre = cell_str(&rows, i, 1);
        if nombre.is_empty() { break; }
        let numero = cell_f64(&rows, i, 0).map(|f| json!(f as i64)).unwrap_or(json!(alumnos.len() + 1));
        alumnos.push(json!({ "numero": numero, "nombre": nombre }));
    }
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "alumnos": alumnos }))
}

#[tauri::command]
fn excel_get_alumnos() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() { Some(p) => load_alumnos(&p), None => Ok(Value::Null) }
}

#[tauri::command]
async fn excel_save_alumnos(alumnos: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_alumnos_impl(alumnos))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_alumnos_impl(alumnos: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let arr = alumnos.as_array().ok_or("La lista de alumnos no es valida.")?.clone();
    save_alumnos_to_file(&path, &arr)?;
    load_alumnos(&path)
}

// ---------------------------------------------------------------------------
// Cargar unidades
// ---------------------------------------------------------------------------

fn load_unidades(path: &str) -> Result<Value, String> {
    // Tabla fija I5:K20 (0-indexed: filas 4-19, cols 8=I, 9=J, 10=K)
    let rows = read_sheet_rows(path, "DATOS")?;
    let mut unidades = Vec::new();
    for ri in 4..=19 {
        let codigo = cell_str(&rows, ri, 8);
        let nombre = cell_str(&rows, ri, 9);
        let evaluacion = cell_str(&rows, ri, 10);
        if codigo.is_empty() && nombre.is_empty() { continue; }
        // "label" es para mostrar (con fallback al título de la hoja Ux si nombre
        // está vacío); "nombre" se deja tal cual para el editor de gestor-unidades.
        let label_nombre = resolve_unit_display_name(path, &codigo, &nombre);
        let label = if label_nombre.is_empty() { codigo.clone() } else { format!("{} - {}", codigo, label_nombre) };
        unidades.push(json!({ "codigo": codigo, "nombre": nombre, "evaluacion": evaluacion, "label": label }));
    }
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "unidades": unidades }))
}

#[tauri::command]
fn excel_get_unidades() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() { Some(p) => load_unidades(&p), None => Ok(Value::Null) }
}

#[tauri::command]
async fn excel_save_unidades(unidades: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_unidades_impl(unidades))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_unidades_impl(unidades: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let arr = unidades.as_array().ok_or("La lista de unidades no es valida.")?.clone();
    save_unidades_to_file(&path, &arr)?;
    load_unidades(&path)
}

// ---------------------------------------------------------------------------
// Cargar RRAA y criterios
// ---------------------------------------------------------------------------

// En ESO los códigos de criterio son del tipo "CR1.1", "CR2.3", etc.
fn is_cr_code(s: &str) -> bool {
    parse_cr_code(s).is_some()
}

fn parse_cr_code(s: &str) -> Option<(i64, i64)> {
    let re = Regex::new(r"(?i)^CR(\d+)\.(\d+)$").unwrap();
    let caps = re.captures(s.trim())?;
    let ce_num = caps.get(1)?.as_str().parse::<i64>().ok()?;
    let cr_num = caps.get(2)?.as_str().parse::<i64>().ok()?;
    Some((ce_num, cr_num))
}

fn canonical_cr_code(ce_num: i64, cr_num: i64) -> String {
    format!("CR{}.{}", ce_num, cr_num)
}

fn next_cr_code_for_ce(raw_code: &str, last_by_ce: &mut HashMap<i64, i64>) -> Option<(i64, i64, String)> {
    let (ce_num, mut cr_num) = parse_cr_code(raw_code)?;
    let last = last_by_ce.entry(ce_num).or_insert(0);
    if cr_num <= *last {
        cr_num = *last + 1;
    }
    *last = cr_num;
    Some((ce_num, cr_num, canonical_cr_code(ce_num, cr_num)))
}

// Compatibilidad: is_criterion_code no se usa en ESO pero se conserva para funciones heredadas
fn is_criterion_code(_s: &str) -> bool { false }

fn normalize_criterion_code(s: &str) -> String {
    s.trim().to_uppercase()
}

fn load_rraa_criterios(path: &str) -> Result<Value, String> {
    let (ce_list, criterios, ponderaciones_unidad) = extract_rraa_criterios_data(path)
        .ok_or("No se encontraron las hojas PESOS o DATOS.")?;
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    let rows = read_sheet_rows(path, "DATOS").unwrap_or_default();
    let mut instrumentos: Vec<Value> = Vec::new();
    for ri in 4..=13usize {
        let codigo = cell_str(&rows, ri, 13);
        let nombre = cell_str(&rows, ri, 14);
        if !codigo.is_empty() { instrumentos.push(json!({ "codigo": codigo, "nombre": nombre })); }
    }
    Ok(json!({ "filePath": path, "fileName": file_name, "rraa": ce_list, "criterios": criterios, "ponderacionesUnidad": ponderaciones_unidad, "instrumentos": instrumentos }))
}

#[tauri::command]
fn excel_get_rraa_criterios() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() { Some(p) => load_rraa_criterios(&p), None => Ok(Value::Null) }
}

#[tauri::command]
async fn excel_save_rraa_criterios(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_rraa_criterios_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_rraa_criterios_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let rraa = payload["rraa"].as_array().cloned().unwrap_or_default();
    let criterios = payload["criterios"].as_array().cloned().unwrap_or_default();
    let pond_unidad = payload["ponderacionesUnidad"].as_array().cloned().unwrap_or_default();
    save_rraa_criterios_to_file(&rraa, &criterios, &pond_unidad, &path)?;
    load_rraa_criterios(&path)
}

// Estructura PESOS en ESO:
//   Fila 3 (idx 2): cabeceras CE  — col B=CEspecíf.1, col C=CEspecíf.2 ... col AE=CEspecíf.6
//   Fila 4 (idx 3): códigos CR    — col C=CR1.1, D=CR1.2 ... AK=CR6.6
//   Filas 5+ (idx 4+): una por unidad, col A=nombre, cols CE/CR = ponderaciones
//
// ce_list = lista de CE con su índice de columna de inicio
// criterios = lista de CR con colIdx absoluto en PESOS
fn extract_rraa_criterios_data(path: &str) -> Option<(Vec<Value>, Vec<Value>, Vec<Value>)> {
    let datos_rows = read_sheet_rows(path, "DATOS").ok()?;

    // Ponderaciones: hoja PESOS — leer filas 3 y 4 via ZIP/XML para alcanzar columnas lejanas
    let pesos_rows = read_sheet_rows(path, "PESOS").ok();
    // Fila 3 (row_1=3) = CE headers, Fila 4 (row_1=4) = CR headers — via XML directo
    let pesos_ce_row = read_row_from_xml(path, "PESOS", 3);  // col_idx -> texto CE
    let pesos_cr_row = read_row_from_xml(path, "PESOS", 4);  // col_idx -> código CR

    // Texto de criterio: DATOS col W(22)=código CR, col X(23)=texto — una fila por CR
    let mut cr_text_map: HashMap<String, String> = HashMap::new();
    for row_idx in 0..datos_rows.len() {
        let code = cell_str(&datos_rows, row_idx, 22);
        if is_cr_code(&code) {
            let texto = cell_str(&datos_rows, row_idx, 23);
            cr_text_map.insert(normalize_criterion_code(&code), texto);
        }
    }

    // Mapa CR código → columna real en PESOS (fila 4)
    let mut cr_col_map: std::collections::HashMap<String, usize> = std::collections::HashMap::new();
    {
        let mut last_cr_by_ce: HashMap<i64, i64> = HashMap::new();
        let mut sorted_cols: Vec<usize> = pesos_cr_row.keys().cloned().collect();
        sorted_cols.sort();
        for ci in sorted_cols {
            let code = pesos_cr_row.get(&ci).cloned().unwrap_or_default();
            if let Some((_, _, fixed_code)) = next_cr_code_for_ce(&code, &mut last_cr_by_ce) {
                cr_col_map.insert(normalize_criterion_code(&fixed_code), ci);
            }
        }
    }

    // CE: leer de PESOS fila 3 via XML directo — formato "CE1. texto", "CE.1 texto", "CE10", etc.
    let re_ce_header = Regex::new(r"(?i)^CE\.?\s*(\d+)").unwrap();
    let mut ce_list: Vec<Value> = Vec::new();
    {
        let mut seen_ce: std::collections::HashSet<i64> = std::collections::HashSet::new();
        let mut sorted_cols: Vec<usize> = pesos_ce_row.keys().cloned().collect();
        sorted_cols.sort();
        for ci in sorted_cols {
            let text = pesos_ce_row.get(&ci).cloned().unwrap_or_default();
            if text.is_empty() || text == "UNIDADES" { continue; }
            if let Some(caps) = re_ce_header.captures(&text) {
                let num: i64 = caps.get(1).and_then(|m| m.as_str().parse().ok()).unwrap_or(0);
                if num > 0 && seen_ce.insert(num) {
                    ce_list.push(json!({ "numero": num, "descripcion": text, "colIdx": ce_list.len() }));
                }
            }
        }
    }
    // Fallback: inferir CE desde CR conocidos si PESOS fila 3 no tiene headers CE
    if ce_list.is_empty() {
        let mut seen: std::collections::BTreeSet<i64> = std::collections::BTreeSet::new();
        for key in cr_col_map.keys() {
            if let Some((ce_num, _)) = parse_cr_code(key) { seen.insert(ce_num); }
        }
        for num in seen {
            ce_list.push(json!({ "numero": num, "descripcion": format!("CE{}", num), "colIdx": ce_list.len() }));
        }
    }

    // CR: construir desde cr_col_map (PESOS fila 4 via XML)
    let mut criterios: Vec<Value> = Vec::new();
    let mut last_cr_by_ce2: HashMap<i64, i64> = HashMap::new();
    {
        let mut sorted_cols: Vec<usize> = pesos_cr_row.keys().cloned().collect();
        sorted_cols.sort();
        for ci in sorted_cols {
            let code = pesos_cr_row.get(&ci).cloned().unwrap_or_default();
            if !is_cr_code(&code) { continue; }
            let (ce_num, _, fixed_cr_code) = match next_cr_code_for_ce(&code, &mut last_cr_by_ce2) {
                Some(v) => v, None => continue,
            };
            let ce_desc = ce_list.iter().find(|ce| ce["numero"].as_i64() == Some(ce_num))
                .and_then(|ce| ce["descripcion"].as_str()).unwrap_or("").to_string();
            let texto = cr_text_map.get(&normalize_criterion_code(&fixed_cr_code))
                .or_else(|| cr_text_map.get(&normalize_criterion_code(&code)))
                .cloned().unwrap_or_default();
            criterios.push(json!({
                "numero": criterios.len() + 1,
                "codigo": fixed_cr_code.clone(),
                "nombre": fixed_cr_code.clone(),
                "originalCodigo": code,
                "raNumero": ce_num,
                "raDescripcion": ce_desc,
                "texto": texto,
                "colIdx": criterios.len(),  // índice secuencial (lo usa el frontend como key)
                "actualCol": ci  // columna real en PESOS (para leer ponderaciones)
            }));
        }
    }
    // Ponderación total por CE (fila idx 22) y por CR (fila idx 21) en PESOS
    // ce_first_col: ce_num → columna mínima de sus CR (puede ser col "Inst ev." anterior)
    let criterios: Vec<Value> = criterios.into_iter().filter(|c| {
        let cr_code = c["codigo"].as_str().unwrap_or("");
        let Some((ce_num, cr_num)) = parse_cr_code(cr_code) else { return false; };
        c["raNumero"].as_i64() == Some(ce_num)
            && cr_num <= 20
            && cr_col_map.contains_key(&normalize_criterion_code(cr_code))
    }).enumerate().map(|(idx, mut c)| {
        if let Value::Object(ref mut obj) = c {
            obj.insert("numero".to_string(), json!(idx + 1));
            obj.insert("colIdx".to_string(), json!(idx)); // recalcular secuencial post-filter
        }
        c
    }).collect();
    let mut ce_first_col: std::collections::HashMap<i64, usize> = std::collections::HashMap::new();
    for c in &criterios {
        let ce_num = c["raNumero"].as_i64().unwrap_or(0);
        let cr_code = c["codigo"].as_str().unwrap_or("");
        if let Some(&col) = cr_col_map.get(&normalize_criterion_code(cr_code)) {
            let entry = ce_first_col.entry(ce_num).or_insert(col);
            if col < *entry { *entry = col; }
        }
    }
    // instrColIdx por CE = columna "Inst ev." = first_cr_col - 1
    let ce_instr_col: std::collections::HashMap<i64, usize> = ce_first_col.iter()
        .filter_map(|(&ce, &col)| if col > 0 { Some((ce, col - 1)) } else { None })
        .collect();
    // Filas de totales en PESOS via XML: fila 22 = ponderacionTotal CE, fila 21 = % CR
    let pesos_row22 = read_row_from_xml(path, "PESOS", 23); // fila 23 = Ponderación final CE (%)
    let pesos_row21 = read_row_from_xml(path, "PESOS", 22); // fila 22 = %ponderación final C.R.
    // Añadir ponderacionTotal e instrColIdx a cada CE
    let ce_list: Vec<Value> = ce_list.into_iter().map(|ce| {
        let ce_num = ce["numero"].as_i64().unwrap_or(0);
        let pond_total = ce_first_col.get(&ce_num).and_then(|&col| {
            let v: f64 = pesos_row22.get(&col).and_then(|s| s.parse().ok()).unwrap_or(0.0);
            if v != 0.0 { return Some(v); }
            if col > 0 { pesos_row22.get(&(col-1)).and_then(|s| s.parse().ok()) } else { None }
        }).unwrap_or(0.0);
        let instr_col = ce_instr_col.get(&ce_num).copied().unwrap_or(0);
        let mut obj = ce.as_object().cloned().unwrap_or_default();
        obj.insert("ponderacionTotal".to_string(), json!(pond_total));
        obj.insert("instrColIdx".to_string(), json!(instr_col));
        Value::Object(obj)
    }).collect();
    // Añadir ponderacionCR y actualCol (columna real en PESOS) a cada criterio
    let criterios: Vec<Value> = criterios.into_iter().map(|c| {
        let cr_code = c["codigo"].as_str().unwrap_or("").to_string();
        let actual_col = cr_col_map.get(&normalize_criterion_code(&cr_code)).copied();
        let pct: f64 = actual_col
            .and_then(|col| pesos_row21.get(&col).and_then(|s| s.parse().ok()))
            .unwrap_or(0.0);
        let mut obj = c.as_object().cloned().unwrap_or_default();
        obj.insert("ponderacionCR".to_string(), json!(pct));
        obj.insert("actualCol".to_string(), json!(actual_col.unwrap_or(0)));
        Value::Object(obj)
    }).collect();

    // Leer ponderaciones por unidad desde PESOS filas 5-20 via XML directo
    let mut ponderaciones_unidad: Vec<Value> = Vec::new();
    for i in 0..16usize {
        let row_idx = i + 4;    // 0-indexed para calamine (DATOS)
        let row_1 = i + 5;     // 1-indexed para XML (PESOS: fila 5 = U1, ..., fila 20 = U16)
        let codigo_datos = cell_str(&datos_rows, row_idx, 8);
        let nombre_datos = cell_str(&datos_rows, row_idx, 9);
        let evaluacion_datos = cell_str(&datos_rows, row_idx, 10);
        // Leer fila de pesos vía XML para alcanzar columnas lejanas
        let pesos_xml_row = read_row_from_xml(path, "PESOS", row_1);
        let nombre_raw = pesos_xml_row.get(&0).cloned().unwrap_or_default();
        let nombre = if !nombre_raw.is_empty() && nombre_raw != "0" {
            nombre_raw
        } else if !nombre_datos.is_empty() && nombre_datos != "0" {
            nombre_datos.clone()
        } else if !codigo_datos.is_empty() && codigo_datos != "0" {
            codigo_datos.clone()
        } else {
            String::new()
        };
        let mut ponderaciones = serde_json::Map::new();
        for c in &criterios {
            let ci = c["colIdx"].as_u64().unwrap_or(0) as usize;
            let cr_code = c["codigo"].as_str().unwrap_or("");
            let actual_col = cr_col_map.get(&normalize_criterion_code(cr_code)).copied().unwrap_or(ci);
            let pond: f64 = pesos_xml_row.get(&actual_col)
                .and_then(|v| v.parse::<f64>().ok()).unwrap_or(0.0);
            ponderaciones.insert(ci.to_string(), json!({ "ponderacion": pond }));
        }
        let mut instr_por_ce = serde_json::Map::new();
        for (&ce_num, &instr_col) in &ce_instr_col {
            let instr = pesos_xml_row.get(&instr_col).cloned().unwrap_or_default();
            instr_por_ce.insert(ce_num.to_string(), json!({ "codigo": instr, "colIdx": instr_col }));
        }
        ponderaciones_unidad.push(json!({
            "numero": i + 1,
            "rowIdx": row_idx,
            "codigo": codigo_datos,
            "nombre": nombre,
            "evaluacion": evaluacion_datos,
            "ponderaciones": ponderaciones,
            "instrPorCe": instr_por_ce
        }));
    }

    Some((ce_list, criterios, ponderaciones_unidad))
}

// ---------------------------------------------------------------------------
// Tipos de actividad y bloques
// ---------------------------------------------------------------------------

struct ActivityType { key: &'static str, label: &'static str, base_col: usize }

const ACTIVITY_TYPES: &[ActivityType] = &[
    ActivityType { key: "practicas",  label: "Practicas",                base_col: 0   },
    ActivityType { key: "memorias",   label: "Memorias",                 base_col: 112 },
    ActivityType { key: "otros",      label: "Otras actividades",        base_col: 223 },
    ActivityType { key: "controles",  label: "Control teorico/practico", base_col: 334 },
];

struct ActivityFixedCols { start: usize, end: usize }

fn activity_fixed_cols(key: &str) -> ActivityFixedCols {
    match key {
        "practicas"  => ActivityFixedCols { start: 0,   end: 109 },
        "memorias"   => ActivityFixedCols { start: 112, end: 220 },
        "otros"      => ActivityFixedCols { start: 223, end: 331 },
        "controles"  => ActivityFixedCols { start: 334, end: 442 },
        _            => ActivityFixedCols { start: 0,   end: 109 },
    }
}

fn get_activity_type(key: &str) -> &'static ActivityType {
    ACTIVITY_TYPES.iter().find(|t| t.key == key).unwrap_or(&ACTIVITY_TYPES[0])
}

#[derive(Clone, Debug)]
#[allow(dead_code)]
struct ActivityBlock {
    tipo: String, tipo_label: String, numero: i64,
    title_row: usize, number_row: usize, included_row: usize,
    header_row: usize, first_student_row: usize,
    name_col: usize, note_col: usize, number_col: usize,
    name_value_col: usize, included_col: usize,
    nombre: String, incluida: bool,
    ce_cols: Vec<(String, usize)>,
}

fn find_activity_header_col(row: &[Value], start_col: usize, text: &str) -> Option<usize> {
    let target = normalize_plain(text);
    (start_col..(start_col + 8)).find(|&ci| {
        row.get(ci).map(|v| normalize_plain(&cell_val_str(v)).contains(&target)).unwrap_or(false)
    })
}

fn resolve_name_value_col(row: &[Value], start_col: usize) -> usize {
    find_activity_header_col(row, start_col, "NOMBRE").map(|c| c + 1).unwrap_or(start_col + 3)
}

fn find_activity_blocks(rows: &[Vec<Value>], tipo_key: &str) -> Vec<ActivityBlock> {
    let at = get_activity_type(tipo_key);
    let mut blocks = Vec::new();
    for row_idx in 0..rows.len().saturating_sub(4) {
        let number_cell = rows.get(row_idx + 1).and_then(|r| r.get(at.base_col + 1)).cloned().unwrap_or(Value::Null);
        let activity_number: i64 = match &number_cell {
            Value::Number(n) => match n.as_i64().or_else(|| n.as_f64().map(|f| f as i64)) { Some(v) => v, None => continue },
            Value::String(s) => match s.trim().parse::<f64>() { Ok(f) if f > 0.0 => f as i64, _ => continue },
            _ => continue,
        };
        let header_row_data = rows.get(row_idx + 3).cloned().unwrap_or_default();
        let name_col = match find_activity_header_col(&header_row_data, at.base_col, "NOMBRE Y APELLIDOS") { Some(c) => c, None => continue };
        let note_col = match find_activity_header_col(&header_row_data, at.base_col, "NOTA FINAL") { Some(c) => c, None => continue };
        let row1 = rows.get(row_idx + 1).cloned().unwrap_or_default();
        let name_value_col = resolve_name_value_col(&row1, at.base_col);
        let nombre = cell_val_str(row1.get(name_value_col).unwrap_or(&Value::Null)).trim().to_string();
        // CE codes are in row_idx+2 (the INCLUIDO/FECHA row), not the header row
        let ce_code_row = rows.get(row_idx + 2).cloned().unwrap_or_default();
        let included_val = ce_code_row.get(at.base_col + 1).map(cell_val_str).unwrap_or_default();
        let fixed = activity_fixed_cols(at.key);
        let mut ce_cols = Vec::new();
        for ci in (note_col + 1)..=fixed.end {
            if let Some(v) = ce_code_row.get(ci) {
                let s = cell_val_str(v);
                let strim = s.trim().to_lowercase();
                if strim.is_empty() { continue; }
                let code = if strim.ends_with(')') { strim.clone() } else { format!("{strim})") };
                if is_criterion_code(&code) { ce_cols.push((code, ci)); }
            }
        }
        blocks.push(ActivityBlock {
            tipo: at.key.to_string(), tipo_label: at.label.to_string(), numero: activity_number,
            title_row: row_idx, number_row: row_idx + 1, included_row: row_idx + 2,
            header_row: row_idx + 3, first_student_row: row_idx + 4,
            name_col, note_col, number_col: at.base_col + 1, name_value_col,
            included_col: at.base_col + 1, nombre, incluida: normalize_plain(&included_val) == "X",
            ce_cols,
        });
    }
    blocks
}

fn format_activity_block(b: &ActivityBlock) -> Value {
    let label = if b.nombre.is_empty() { format!("{} {}", b.tipo_label, b.numero) } else { format!("{} {} - {}", b.tipo_label, b.numero, b.nombre) };
    json!({ "numero": b.numero, "nombre": b.nombre, "incluida": b.incluida, "label": label, "firstStudentRow": b.first_student_row, "noteCol": b.note_col })
}

fn extract_activity_notes(rows: &[Vec<Value>], block: &ActivityBlock, max_alumnos: Option<usize>) -> Vec<Value> {
    let mut notes = Vec::new();
    for row_idx in block.first_student_row..rows.len() {
        let nombre = cell_str(rows, row_idx, block.name_col);
        if nombre.is_empty() || nombre == "0" { break; }
        let nota_raw = rows.get(row_idx).and_then(|r| r.get(block.note_col)).cloned().unwrap_or(Value::Null);
        let nota_str = match &nota_raw {
            Value::Number(n) => n.as_f64().map(|f| f.to_string().replace('.', ",")),
            Value::String(s) => Some(s.replace('.', ",")),
            _ => None,
        }.unwrap_or_default();
        let mut ce_notas = serde_json::Map::new();
        for (code, ci) in &block.ce_cols {
            let val = rows.get(row_idx).and_then(|r| r.get(*ci)).cloned().unwrap_or(Value::Null);
            let nota_ce = match &val {
                Value::Number(n) => n.as_f64().map(|f| f.to_string().replace('.', ",")),
                Value::String(s) if !s.trim().is_empty() => Some(s.replace('.', ",")),
                _ => None,
            };
            if let Some(s) = nota_ce { ce_notas.insert(code.clone(), json!(s)); }
        }
        notes.push(json!({ "numero": notes.len() + 1, "rowIdx": row_idx, "nombre": nombre, "nota": nota_str, "ceNotas": Value::Object(ce_notas) }));
        if let Some(max) = max_alumnos { if notes.len() >= max { break; } }
    }
    notes
}

// Nombre a mostrar para una unidad: usa el de DATOS si lo hay; si no, lee el
// título de la propia hoja Ux (fila 3, col A), que es una fórmula INDEX/MATCH
// cuyo caché calamine a veces no lee bien -> se lee también por XML directo.
fn resolve_unit_display_name(path: &str, codigo: &str, raw_nombre: &str) -> String {
    if !raw_nombre.is_empty() && raw_nombre.to_uppercase() != codigo.to_uppercase() {
        return raw_nombre.to_string();
    }
    let re_code_prefix = Regex::new(r"(?i)^U\d+[\s.\-\u{2013}]+").unwrap();
    read_row_from_xml(path, codigo, 3).get(&0).cloned()
        .filter(|t| !t.is_empty())
        .or_else(|| {
            read_sheet_rows(path, codigo).ok().and_then(|rows| {
                let t = cell_str(&rows, 2, 0);
                if t.is_empty() { None } else { Some(t) }
            })
        })
        .map(|t| {
            let limpio = re_code_prefix.replace(&t, "").trim().to_string();
            if limpio.is_empty() { t } else { limpio }
        })
        .unwrap_or_else(|| raw_nombre.to_string())
}

fn list_unit_sheets(path: &str) -> Result<Vec<Value>, String> {
    let names = sheet_names(path)?;
    let datos_rows = read_sheet_rows(path, "DATOS").unwrap_or_default();
    let mut units_from_datos: HashMap<String, String> = HashMap::new();
    let unidad_start = find_unidades_start(&datos_rows).unwrap_or(4);
    for idx in 0..16 {
        let code = { let s = cell_str(&datos_rows, unidad_start + idx, 8); if s.is_empty() { format!("U{}", idx + 1) } else { s } };
        let nombre = cell_str(&datos_rows, unidad_start + idx, 9);
        if !code.is_empty() || !nombre.is_empty() {
            units_from_datos.insert(code.to_uppercase(), nombre);
        }
    }
    let re = Regex::new(r"(?i)^U\d+$").unwrap();
    let mut unit_names: Vec<String> = names.iter().filter(|n| re.is_match(n)).cloned().collect();
    unit_names.sort_by_key(|n| n.chars().filter(|c| c.is_ascii_digit()).collect::<String>().parse::<i64>().unwrap_or(0));
    Ok(unit_names.iter().map(|codigo| {
        let raw = units_from_datos.get(&codigo.to_uppercase()).cloned().unwrap_or_default();
        let nombre = resolve_unit_display_name(path, codigo, &raw);
        let label = if nombre.is_empty() { codigo.clone() } else { format!("{} - {}", codigo, nombre) };
        json!({ "codigo": codigo, "nombre": nombre, "label": label })
    }).collect())
}

// ---------------------------------------------------------------------------
// Notas actividad
// ---------------------------------------------------------------------------

fn load_notas_actividad(path: &str, unidad: &str, tipo: &str, actividad: i64, max_alumnos: Option<usize>) -> Result<Value, String> {
    let rows = read_sheet_rows(path, unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let selected_block = blocks.iter().find(|b| b.numero == actividad).or_else(|| blocks.first());
    let notas = selected_block.map(|b| extract_activity_notes(&rows, b, max_alumnos)).unwrap_or_default();
    let tipos: Vec<Value> = ACTIVITY_TYPES.iter().map(|t| {
        let tblocks: Vec<Value> = find_activity_blocks(&rows, t.key).iter().map(format_activity_block).collect();
        let incluidas = tblocks.iter().filter(|b| b["incluida"].as_bool().unwrap_or(false)).count();
        let total = tblocks.len();
        json!({ "key": t.key, "label": t.label, "actividades": tblocks, "incluidas": incluidas, "total": total })
    }).collect();
    let unidades = list_unit_sheets(path)?;
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();

    // Cargar RRAA y criterios filtrados para la unidad actual
    // Filtro: CR presentes en la hoja Ux (fila idx=2), no por ponderación
    let (rraa, todas_criterios, ponderaciones_unidad, criterios_unidad) =
        if let Some((rraa, criterios, pu)) = extract_rraa_criterios_data(path) {
            // Detectar qué CR existen en la hoja Ux (fila 3 = row_1=3) via ZIP/XML directo
            let cr_en_hoja: std::collections::HashSet<String> = {
                let row_map = read_row_from_xml(path, unidad, 3);
                row_map.values().filter_map(|v| {
                    if is_cr_code(v) { Some(normalize_criterion_code(v)) } else { None }
                }).collect()
            };
            let unidad_idx = unidades.iter().position(|u| u["codigo"].as_str() == Some(unidad));
            let criterios_filtrados = criterios.iter().filter_map(|c| {
                let cr_code = c["codigo"].as_str().unwrap_or("");
                if !cr_en_hoja.contains(&normalize_criterion_code(cr_code)) { return None; }
                let mut c2 = c.clone();
                // Adjuntar pesos si existen
                if let Some(idx) = unidad_idx {
                    if let Some(pu_entry) = pu.get(idx) {
                        let ci_key = c["colIdx"].as_u64().unwrap_or(0).to_string();
                        if let Some(pesos) = pu_entry["ponderaciones"].get(&ci_key) {
                            if let Value::Object(ref mut obj) = c2 {
                                obj.insert("ponderacionUnidad".to_string(), pesos.clone());
                            }
                        }
                    }
                }
                Some(c2)
            }).collect();
            (rraa, criterios, pu, criterios_filtrados)
        } else {
            (vec![], vec![], vec![], vec![])
        };

    Ok(json!({
        "filePath": path, "fileName": file_name, "unidad": unidad, "tipo": at.key,
        "actividad": selected_block.map(|b| b.numero).unwrap_or(actividad),
        "unidades": unidades, "tipos": tipos,
        "actividades": blocks.iter().map(format_activity_block).collect::<Vec<_>>(),
        "notas": notas,
        "block": selected_block.map(format_activity_block),
        "rraa": rraa,
        "criterios": criterios_unidad,
        "todasCriterios": todas_criterios,
        "ponderacionesUnidad": ponderaciones_unidad
    }))
}

#[tauri::command]
fn excel_get_notas_actividad(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    let tipo = payload["tipo"].as_str().unwrap_or("practicas").to_string();
    let actividad = payload["actividad"].as_i64().unwrap_or(1);
    let max_alumnos = payload["maxAlumnos"].as_u64().map(|n| n as usize);
    load_notas_actividad(&path, &unidad, &tipo, actividad, max_alumnos)
}

fn load_notas_actividades_tipo(path: &str, unidad: &str, tipo: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let activities: Vec<Value> = blocks.iter().map(|b| {
        let notas = extract_activity_notes(&rows, b, None);
        let mut v = format_activity_block(b);
        v["tipo"] = json!(at.key); v["tipoLabel"] = json!(at.label); v["notas"] = json!(notas); v
    }).collect();
    let tipos: Vec<Value> = ACTIVITY_TYPES.iter().map(|t| {
        let tblocks: Vec<Value> = find_activity_blocks(&rows, t.key).iter().map(format_activity_block).collect();
        let incluidas = tblocks.iter().filter(|b| b["incluida"].as_bool().unwrap_or(false)).count();
        let total = tblocks.len();
        json!({ "key": t.key, "label": t.label, "actividades": tblocks, "incluidas": incluidas, "total": total })
    }).collect();
    let unidades = list_unit_sheets(path)?;
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "unidad": unidad, "unidades": unidades, "tipo": at.key, "tipos": tipos, "activities": activities }))
}

#[tauri::command]
fn excel_get_notas_actividades_tipo(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    let tipo = payload["tipo"].as_str().unwrap_or("practicas").to_string();
    load_notas_actividades_tipo(&path, &unidad, &tipo)
}

// ---------------------------------------------------------------------------
// Notas evaluación
// ---------------------------------------------------------------------------

fn find_evaluation_sheet_name(names: &[String], evaluacion: &str) -> Option<String> {
    let target = evaluacion.trim();
    if target == "final" { return names.iter().find(|n| normalize_plain(n) == "FINAL").cloned(); }
    if target == "2solo" || target == "3solo" {
        let num = &target[..1];
        return names.iter().find(|n| { let norm = normalize_plain(n); norm.contains(num) && norm.contains("EVA") && norm.contains("SOLO") }).cloned();
    }
    names.iter().find(|n| { let norm = normalize_plain(n); norm.contains(&target.to_uppercase()) && norm.contains("EVA") && !norm.contains("MAX") && !norm.contains("SOLO") }).cloned()
}

fn is_eval_criterion_code(s: &str) -> bool {
    let s = s.trim();
    is_cr_code(s)
        || Regex::new(r"^\d+\.?[a-z]\)?$").unwrap().is_match(s)
        || Regex::new(r"^\d+\.\d+$").unwrap().is_match(s)
}

fn is_nota_ce_cell(v: &Value) -> bool {
    let txt = normalize_plain(&cell_val_str(v));
    txt == "NOTA CE" || txt.starts_with("NOTA CE") || txt == "NOTA_CE"
}

fn evaluation_ce_number_from_value(v: &Value) -> Option<i64> {
    let txt = normalize_plain(&cell_val_str(v));
    Regex::new(r"^CE[^0-9]*(\d+)")
        .unwrap()
        .captures(&txt)
        .and_then(|caps| caps.get(1))
        .and_then(|m| m.as_str().parse::<i64>().ok())
}

fn find_evaluation_ce_anchors(rows: &[Vec<Value>], code_row_idx: usize) -> Vec<(usize, i64)> {
    let start_row = code_row_idx.saturating_sub(4);

    for row_idx in (start_row..=code_row_idx).rev() {
        let anchors: Vec<(usize, i64)> = rows.get(row_idx)
            .map(|row| {
                row.iter()
                    .enumerate()
                    .filter_map(|(ci, value)| evaluation_ce_number_from_value(value).map(|num| (ci, num)))
                    .collect()
            })
            .unwrap_or_default();

        if anchors.len() >= 2 {
            return anchors;
        }
    }

    Vec::new()
}

fn load_notas_evaluacion(path: &str, evaluacion: &str) -> Result<Value, String> {
    let names = sheet_names(path)?;
    let sheet_name = find_evaluation_sheet_name(&names, evaluacion)
        .ok_or_else(|| format!("No se encontro la hoja de la {evaluacion} evaluacion."))?;
    let rows = read_sheet_rows(path, &sheet_name)?;

    // Buscar layout:
    // ESO: NOTA CE y CR codes están en la MISMA fila (ej. fila 17); fila+1 = sub-cabecera "Rec"; fila+2 = alumnos
    // FP:  NOTA CE en fila N, CR codes en fila N+1; alumnos en N+2
    // Fallback: primera fila con >= 2 CR codes
    let (summary_row_idx, code_row_idx, first_student_row_idx) = {
        let mut found = None;

        // Estrategia 1 (ESO): misma fila tiene NOTA CE y >= 2 CR codes
        for row_idx in 0..rows.len() {
            let has_nota_ce = rows[row_idx].iter().any(|v| is_nota_ce_cell(v));
            let cr_count = rows[row_idx].iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count();
            if has_nota_ce && cr_count >= 2 {
                found = Some((row_idx, row_idx, row_idx + 2)); // +2 para saltar fila "Rec"
                break;
            }
        }

        // Estrategia 2 (FP): fila NOTA CE + CR codes en la siguiente (hasta 3 filas después)
        if found.is_none() {
            for row_idx in 0..rows.len() {
                if rows[row_idx].iter().filter(|v| is_nota_ce_cell(v)).count() == 0 { continue; }
                for offset in 1..=3 {
                    if let Some(next) = rows.get(row_idx + offset) {
                        if next.iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count() >= 2 {
                            found = Some((row_idx, row_idx + offset, row_idx + offset + 1));
                            break;
                        }
                    }
                }
                if found.is_some() { break; }
            }
        }

        // Estrategia 3 (fallback): primera fila con >= 2 CR codes
        if found.is_none() {
            for row_idx in 1..rows.len() {
                if rows[row_idx].iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count() >= 2 {
                    found = Some((row_idx.saturating_sub(1), row_idx, row_idx + 1));
                    break;
                }
            }
        }

        found.ok_or("No se encontro la cabecera de notas de evaluacion.")?
    };

    let summary_row = &rows[summary_row_idx];
    let mut ra_columns: Vec<Value> = Vec::new();
    let ce_anchors = find_evaluation_ce_anchors(&rows, code_row_idx);
    if !ce_anchors.is_empty() {
        for (ci, ra_num) in ce_anchors {
            let peso = cell_str(&rows, 13, ci);
            ra_columns.push(json!({ "colIdx": ci, "address": col_name(ci), "label": format!("RRAA {ra_num}"), "numero": ra_num, "peso": peso }));
        }
    } else {
        for (ci, cell) in summary_row.iter().enumerate() {
            if !is_nota_ce_cell(cell) { continue; }
            let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
            let ra_num = code_row[ci+1..].iter().find_map(|v| {
                let s = cell_val_str(v);
                let norm = normalize_plain(&s);
                if norm == "NOTA CE" || norm == "NOTA FINAL" { return None; }
                s.trim().chars().take_while(|c| c.is_ascii_digit()).collect::<String>().parse::<i64>().ok()
            }).unwrap_or(ra_columns.len() as i64 + 1);
            let peso = cell_str(&rows, 13, ci);
            ra_columns.push(json!({ "colIdx": ci, "address": col_name(ci), "label": format!("RRAA {ra_num}"), "numero": ra_num, "peso": peso }));
        }
    }

    let final_col = summary_row.iter().position(|v| normalize_plain(&cell_val_str(v)) == "NOTA FINAL")
        .or_else(|| {
            let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
            code_row.iter().enumerate().rev().find_map(|(ci, v)| {
                let s = cell_val_str(v);
                if is_eval_criterion_code(&s) || normalize_plain(&s) == "REC" { Some(ci + 1) } else { None }
            })
        })
        .unwrap_or_else(|| col_index("CB")); // columna CB fija en plantilla ESO
    let final_cb_xml = read_col_values_from_xml(path, &sheet_name, col_index("CB"));

    let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
    // Fila de sub-etiquetas "Rec" está justo debajo de la fila de cabecera
    let rec_label_row = rows.get(code_row_idx + 1).cloned().unwrap_or_default();
    let mut criteria: Vec<Value> = Vec::new();
    for (ra_idx, ra) in ra_columns.iter().enumerate() {
        let ra_ci = ra["colIdx"].as_u64().unwrap() as usize;
        let next_ra_ci = ra_columns.get(ra_idx + 1).and_then(|r| r["colIdx"].as_u64()).map(|n| n as usize).unwrap_or(final_col);
        for ci in (ra_ci + 1)..next_ra_ci {
            let code = cell_val_str(code_row.get(ci).unwrap_or(&Value::Null));
            if !is_eval_criterion_code(&code) { continue; }
            // Columna Rec adyacente: ci+1 si no tiene código CR ni NOTA CE
            let rec_ci: i64 = {
                let nc = ci + 1;
                let nc_code = cell_val_str(code_row.get(nc).unwrap_or(&Value::Null));
                let nc_is_nota_ce = is_nota_ce_cell(code_row.get(nc).unwrap_or(&Value::Null));
                if nc < next_ra_ci && !is_eval_criterion_code(&nc_code) && !nc_is_nota_ce {
                    nc as i64
                } else { -1 }
            };
            criteria.push(json!({ "colIdx": ci, "address": col_name(ci), "raColIdx": ra_ci, "raLabel": ra["label"], "codigo": code.trim(), "peso": cell_str(&rows, 12, ci), "recColIdx": rec_ci }));
        }
    }

    let mut alumnos: Vec<Value> = Vec::new();
    for row_idx in first_student_row_idx..rows.len() {
        let nombre = cell_str(&rows, row_idx, 0);
        if nombre.is_empty() { continue; }
        let norm = normalize_plain(&nombre);
        if norm.contains("MEDIA") || norm.contains("PONDERACION") { continue; }

        // Criterios: celdas planas (no fórmulas) — siempre correctos
        let crit_vals: Vec<Value> = criteria.iter().map(|c| {
            let ci = c["colIdx"].as_u64().unwrap() as usize;
            let rec_ci = c["recColIdx"].as_i64().filter(|&n| n >= 0).map(|n| n as usize);
            json!({
                "colIdx": ci, "raColIdx": c["raColIdx"], "raLabel": c["raLabel"], "codigo": c["codigo"],
                "nota": cell_f64(&rows, row_idx, ci), "display": cell_str(&rows, row_idx, ci),
                "recColIdx": c["recColIdx"],
                "recDisplay": rec_ci.map(|r| cell_str(&rows, row_idx, r)).unwrap_or_default(),
            })
        }).collect();

        // NOTA CE: calcular desde criterios brutos (evita cache stale de fórmulas Excel)
        let rraa_vals: Vec<Value> = ra_columns.iter().map(|ra| {
            let ra_ci = ra["colIdx"].as_u64().unwrap() as usize;
            let crs: Vec<&Value> = criteria.iter()
                .filter(|c| c["raColIdx"].as_u64().map(|x| x as usize) == Some(ra_ci))
                .collect();
            let (nota, display) = if crs.is_empty() {
                let n = cell_f64(&rows, row_idx, ra_ci).unwrap_or(0.0);
                (n, format!("{:.2}", n).replace('.', ","))
            } else {
                let mut ws = 0.0f64;
                let mut ps = 0.0f64;
                for cr in &crs {
                    let cr_ci = cr["colIdx"].as_u64().unwrap() as usize;
                    let peso = cell_f64(&rows, 12, cr_ci).unwrap_or(0.0);
                    let nota = cell_f64(&rows, row_idx, cr_ci).unwrap_or(0.0);
                    let rec_ci = cr["recColIdx"].as_i64().filter(|&n| n >= 0).map(|n| n as usize);
                    let rec = rec_ci.and_then(|r| cell_f64(&rows, row_idx, r));
                    let eff = match rec { Some(r) if r > 0.0 => r.max(nota), _ => nota };
                    ws += peso * eff;
                    ps += peso;
                }
                let n = if ps > 0.0 { ws / ps } else { cell_f64(&rows, row_idx, ra_ci).unwrap_or(0.0) };
                (n, format!("{:.2}", n).replace('.', ","))
            };
            json!({ "colIdx": ra_ci, "label": ra["label"], "numero": ra["numero"], "nota": nota, "display": display })
        }).collect();

        // NOTA FINAL: calcular desde NOTA CE con pesos de la fila 13 (evita cache stale)
        let (final_f64, final_display) = {
            let weighted: Vec<(f64, f64)> = ra_columns.iter().zip(rraa_vals.iter()).filter_map(|(ra, rv)| {
                let peso = ra["peso"].as_str()
                    .and_then(|s| s.replace(',', ".").parse::<f64>().ok())
                    .unwrap_or(0.0);
                let nota = rv["nota"].as_f64().unwrap_or(0.0);
                if peso > 0.0 { Some((peso, nota)) } else { None }
            }).collect();
            if weighted.is_empty() {
                let fd = final_cb_xml.get(&row_idx).cloned()
                    .filter(|s| !s.is_empty() && s.replace(',', ".").parse::<f64>().is_ok())
                    .unwrap_or_else(|| cell_str(&rows, row_idx, final_col));
                let ff = fd.replace(',', ".").parse::<f64>()
                    .unwrap_or_else(|_| cell_f64(&rows, row_idx, final_col).unwrap_or(0.0));
                (ff, fd)
            } else {
                let ws: f64 = weighted.iter().map(|(p, n)| p * n).sum();
                let ps: f64 = weighted.iter().map(|(p, _)| p).sum();
                let n = if ps > 0.0 { ws / ps } else { 0.0 };
                (n, format!("{:.2}", n).replace('.', ","))
            }
        };

        alumnos.push(json!({ "rowIdx": row_idx, "numero": alumnos.len() + 1, "nombre": nombre, "final": final_f64, "finalDisplay": final_display, "rraa": rraa_vals, "criterios": crit_vals }));
    }

    // Filtrar RRAA y criterios no incluidos en la evaluación (peso 0 o vacío).
    let active_crit_cols: std::collections::HashSet<usize> = criteria.iter()
        .filter(|c| parse_decimal(&c["peso"]) > 0.0)
        .map(|c| c["colIdx"].as_u64().unwrap_or(0) as usize)
        .collect();

    let active_ra_cols: std::collections::HashSet<usize> = ra_columns.iter().filter(|ra| {
        let ra_ci = ra["colIdx"].as_u64().unwrap_or(0) as usize;
        parse_decimal(&ra["peso"]) > 0.0 || criteria.iter().any(|c| {
            c["raColIdx"].as_u64().map(|x| x as usize) == Some(ra_ci)
                && active_crit_cols.contains(&(c["colIdx"].as_u64().unwrap_or(0) as usize))
        })
    }).map(|ra| ra["colIdx"].as_u64().unwrap_or(0) as usize).collect();

    let criteria_filtered: Vec<Value> = criteria.into_iter().filter(|c| {
        active_crit_cols.contains(&(c["colIdx"].as_u64().unwrap_or(0) as usize))
    }).collect();
    let ra_columns_filtered: Vec<Value> = ra_columns.into_iter().filter(|ra| {
        active_ra_cols.contains(&(ra["colIdx"].as_u64().unwrap_or(0) as usize))
    }).collect();
    let alumnos_filtered: Vec<Value> = alumnos.into_iter().map(|mut a| {
        if let Some(arr) = a["criterios"].as_array_mut() {
            *arr = arr.iter().filter(|v| {
                let ci = v["colIdx"].as_u64().unwrap_or(0) as usize;
                active_crit_cols.contains(&ci)
            }).cloned().collect();
        }
        if let Some(arr) = a["rraa"].as_array_mut() {
            *arr = arr.iter().filter(|v| {
                let ci = v["colIdx"].as_u64().unwrap_or(0) as usize;
                active_ra_cols.contains(&ci)
            }).cloned().collect();
        }
        a
    }).collect();

    let title = if normalize_plain(&sheet_name) == "FINAL" { "FINAL".to_string() } else { cell_str(&rows, 2, 0) };
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({
        "filePath": path, "fileName": file_name, "sheetName": sheet_name, "title": title, "evaluacion": evaluacion,
        "layout": { "summaryRowIdx": summary_row_idx, "codeRowIdx": code_row_idx, "firstStudentRowIdx": first_student_row_idx },
        "raColumns": ra_columns_filtered, "criteria": criteria_filtered, "alumnos": alumnos_filtered
    }))
}

#[tauri::command]
fn excel_get_notas_evaluacion(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    load_notas_evaluacion(&path, payload["evaluacion"].as_str().unwrap_or("1"))
}

#[tauri::command]
fn excel_get_notas_evaluacion_alumno(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let evaluacion = payload["evaluacion"].as_str().unwrap_or("1").to_string();
    let alumno = payload["alumno"].as_str().unwrap_or("").to_string();
    let mut data = load_notas_evaluacion(&path, &evaluacion)?;
    if let Some(arr) = data["alumnos"].as_array_mut() {
        let filtered: Vec<Value> = arr.iter().filter(|a| a["nombre"].as_str().unwrap_or("") == alumno).cloned().collect();
        data["alumnos"] = json!(filtered);
    }
    Ok(data)
}

// Lee la nota final de la unidad (col E = índice 4).
// Usa el primer bloque de prácticas para localizar name_col y first_student_row,
// garantizando que leemos exactamente las mismas filas de alumnos que el resto del sistema.
fn load_notas_unidad(path: &str, unidad: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, unidad)
        .map_err(|_| format!("No se encontró la hoja \"{unidad}\"."))?;
    let unidades = list_unit_sheets(path)?;

    let first_row: usize = 4; // fila 5 en Excel (0-indexed)

    // Cargar nombres desde load_alumnos (mismo código que el gestor de alumnos)
    let nombres_datos: Vec<String> = match load_alumnos(path) {
        Ok(v) => v["alumnos"].as_array()
            .unwrap_or(&vec![])
            .iter()
            .map(|a| a["nombre"].as_str().unwrap_or("").to_string())
            .collect(),
        Err(_) => Vec::new(),
    };

    // Detectar CRs en hoja Ux usando XML directo (calamine trunca columnas lejanas)
    // Fila 3 en Excel = row_1=3 (1-indexed)
    let mut cr_cols: Vec<(String, usize)> = Vec::new();
    for check_row_1 in [3usize, 4usize] {
        let xml_row = read_row_from_xml(path, unidad, check_row_1);
        if !xml_row.is_empty() {
            let mut sorted: Vec<(usize, String)> = xml_row.into_iter().collect();
            sorted.sort_by_key(|(ci, _)| *ci);
            for (ci, s) in sorted {
                if is_cr_code(&s) { cr_cols.push((s.to_uppercase(), ci)); }
            }
        }
        if !cr_cols.is_empty() { break; }
    }
    // Fallback a calamine si XML falla
    if cr_cols.is_empty() {
        for check_ri in 2..=3 {
            if let Some(row) = rows.get(check_ri) {
                for ci in 0..row.len().min(200) {
                    let s = cell_val_str(row.get(ci).unwrap_or(&Value::Null));
                    if is_cr_code(&s) { cr_cols.push((s.to_uppercase(), ci)); }
                }
            }
            if !cr_cols.is_empty() { break; }
        }
    }

    // Leer ponderaciones de PESOS usando XML directo para ambas filas (CR map y valores de unidad)
    let ponderaciones_por_cr: std::collections::HashMap<String, f64> = {
        let mut map = std::collections::HashMap::new();
        // CR col map desde PESOS fila 4 (1-indexed) vía XML
        let pesos_cr_row = read_row_from_xml(path, "PESOS", 4);
        let cr_col_map: std::collections::HashMap<String, usize> = pesos_cr_row.into_iter()
            .filter(|(_, v)| is_cr_code(v))
            .map(|(ci, v)| (v.to_uppercase(), ci))
            .collect();
        // Calcular fila directamente del nombre de unidad: U1→fila5, U2→fila6, ...
        // unidad tiene formato "U1", "U2", ..., "U16"
        let unidad_num: Option<usize> = {
            let u = unidad.to_uppercase();
            let u = u.trim_start_matches('U');
            u.parse::<usize>().ok()
        };
        if let Some(u_num) = unidad_num {
            let target_row_1 = 4 + u_num; // U1→5, U2→6, ...
            if target_row_1 <= 20 {
                let pesos_row = read_row_from_xml(path, "PESOS", target_row_1);
                for (cr_code, col) in &cr_col_map {
                    if let Some(pond_str) = pesos_row.get(col) {
                        let pond: f64 = pond_str.replace(',', ".").trim().parse().unwrap_or(0.0);
                        map.insert(cr_code.clone(), pond);
                    }
                }
            }
        }
        // Fallback: buscar por nombre en col A si el índice directo no funciona
        if map.is_empty() {
            let unidad_norm = normalize_plain(unidad);
            for unidad_row_1 in 5..=20usize {
                let pesos_row = read_row_from_xml(path, "PESOS", unidad_row_1);
                if let Some(nombre) = pesos_row.get(&0) {
                    let nombre_norm = normalize_plain(nombre);
                    if nombre_norm.contains(&unidad_norm) || unidad_norm.contains(&nombre_norm) {
                        for (cr_code, col) in &cr_col_map {
                            if let Some(pond_str) = pesos_row.get(col) {
                                let pond: f64 = pond_str.replace(',', ".").trim().parse().unwrap_or(0.0);
                                map.insert(cr_code.clone(), pond);
                            }
                        }
                        break;
                    }
                }
            }
        }
        map
    };

    let criterios_json: Vec<Value> = cr_cols.iter()
        .map(|(code, ci)| {
            let ponderacion = ponderaciones_por_cr.get(code).copied().unwrap_or(0.0);
            json!({ "codigo": code, "colIdx": ci, "recColIdx": ci + 1, "ponderacion": ponderacion })
        })
        .collect();

    // Rec: se guarda en la propia hoja de unidad, en la columna adyacente (ci+1) a cada CR.
    let mut alumnos: Vec<Value> = Vec::new();
    let max_alumnos = nombres_datos.len().max(37);
    for ri in first_row..(first_row + max_alumnos).min(rows.len()) {
        let alumno_idx = ri - first_row;
        let nombre = nombres_datos.get(alumno_idx).cloned().unwrap_or_default();
        if nombre.is_empty() { break; }

        let cr_notas: Vec<Value> = cr_cols.iter().map(|(code, ci)| {
            let n = cell_f64(&rows, ri, *ci);
            let d = cell_str(&rows, ri, *ci);
            let rec_display = cell_str(&rows, ri, *ci + 1);
            json!({ "codigo": code, "colIdx": ci, "recColIdx": ci + 1, "nota": n, "display": d, "recDisplay": rec_display })
        }).collect();

        alumnos.push(json!({ "nombre": nombre, "rowIdx": ri, "crNotas": cr_notas }));
    }

    let titulo = cell_str(&rows, 2, 0);
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "unidad": unidad, "titulo": titulo, "unidades": unidades, "criterios": criterios_json, "alumnos": alumnos }))
}

#[tauri::command]
fn excel_get_notas_unidad(payload: Value) -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    load_notas_unidad(&path, &unidad)
}

#[tauri::command]
fn excel_get_alumnos_informes() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    let path = match get_selected_path() { Some(p) => p, None => return Ok(Value::Null) };
    let result = load_alumnos(&path)?;
    Ok(json!(result["alumnos"].as_array().unwrap_or(&vec![]).iter().map(|a| a["nombre"].clone()).collect::<Vec<_>>()))
}

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('"', "&quot;").replace('\'', "&apos;")
}

fn unescape_xml(s: &str) -> String {
    s.replace("&apos;", "'").replace("&quot;", "\"").replace("&gt;", ">").replace("&lt;", "<").replace("&amp;", "&")
}

fn get_xml_attr(tag: &str, name: &str) -> Option<String> {
    let pattern = format!(r#"{}="([^"]*)""#, regex::escape(name).replace("\\:", ":").replace("\\r", "r"));
    // Use raw attribute search for namespaced attrs like r:id
    let search = format!("{}=\"", name);
    if let Some(pos) = tag.find(&search) {
        let after = &tag[pos + search.len()..];
        if let Some(end) = after.find('"') {
            return Some(unescape_xml(&after[..end]));
        }
    }
    let _ = pattern;
    None
}

fn set_xml_attr(tag: &str, name: &str, value: &str) -> String {
    let search = format!("{}=\"", name);
    if let Some(pos) = tag.find(&search) {
        let end = tag[pos + search.len()..].find('"').map(|e| pos + search.len() + e + 1).unwrap_or(tag.len());
        let replacement = format!("{}=\"{}\"", name, escape_xml(value));
        format!("{}{}{}", &tag[..pos], replacement, &tag[end..])
    } else if tag.ends_with("/>") {
        format!("{} {}=\"{}\"/>", &tag[..tag.len()-2], name, escape_xml(value))
    } else {
        let pos = tag.find('>').unwrap_or(tag.len());
        format!("{} {}=\"{}\"{}", &tag[..pos], name, escape_xml(value), &tag[pos..])
    }
}

fn get_xml_row(sheet_xml: &str, row_number: usize) -> Option<String> {
    let pattern = format!(r#"<row\b[^>/]*\br="{row_number}"[^>/]*>[\s\S]*?</row>"#);
    Regex::new(&pattern).ok()?.find(sheet_xml).map(|m| m.as_str().to_string())
}

fn insert_xml_row_at(sheet_xml: &str, row_number: usize) -> Result<String, String> {
    let new_row = format!("<row r=\"{row_number}\"></row>");
    // Expandir <sheetData/> a <sheetData></sheetData> si hace falta
    let xml = if sheet_xml.contains("<sheetData/>") {
        sheet_xml.replace("<sheetData/>", "<sheetData></sheetData>")
    } else {
        sheet_xml.to_string()
    };
    if xml.contains("</sheetData>") {
        Ok(xml.replace("</sheetData>", &format!("{new_row}</sheetData>")))
    } else {
        Err("La hoja no tiene una estructura sheetData valida.".to_string())
    }
}

fn col_index_from_ref(cell_ref: &str) -> usize {
    let col_part: String = cell_ref.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
    col_index(&col_part)
}

fn insert_xml_cell_in_row(row_xml: &str, cell_xml: &str, col_idx: usize) -> String {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    for m in cell_re.find_iter(row_xml) {
        if let Some(r) = get_xml_attr(m.as_str(), "r") {
            if col_index_from_ref(&r) > col_idx {
                return row_xml.replacen(m.as_str(), &format!("{cell_xml}{}", m.as_str()), 1);
            }
        }
    }
    row_xml.replace("</row>", &format!("{cell_xml}</row>"))
}

fn excel_serial_from_date(value: &str) -> Option<i64> {
    let parts: Vec<&str> = value.split('-').collect();
    if parts.len() == 3 {
        let (y, m, d) = (parts[0].parse::<i32>().ok()?, parts[1].parse::<u32>().ok()?, parts[2].parse::<u32>().ok()?);
        use chrono::NaiveDate;
        let date = NaiveDate::from_ymd_opt(y, m, d)?;
        let epoch = NaiveDate::from_ymd_opt(1970, 1, 1)?;
        Some((date - epoch).num_days() + 25569)
    } else {
        None
    }
}

fn build_xml_cell(cell_ref: &str, value: &Value, value_type: &str, style_id: Option<&str>) -> String {
    let style_attr = style_id.map(|s| format!(" s=\"{}\"", escape_xml(s))).unwrap_or_default();
    match value_type {
        "number" => {
            let num = match value {
                Value::Number(n) => n.as_f64().unwrap_or(0.0),
                Value::String(s) => s.replace(',', ".").parse().unwrap_or(0.0),
                _ => return format!("<c r=\"{cell_ref}\"{style_attr}/>"),
            };
            format!("<c r=\"{cell_ref}\"{style_attr}><v>{num}</v></c>")
        }
        "date" => {
            let s = match value { Value::String(s) => s.clone(), _ => return format!("<c r=\"{cell_ref}\"{style_attr}/>") };
            match excel_serial_from_date(&s) {
                Some(serial) => format!("<c r=\"{cell_ref}\"{style_attr}><v>{serial}</v></c>"),
                None => format!("<c r=\"{cell_ref}\"{style_attr}/>"),
            }
        }
        _ => {
            let text = match value {
                Value::String(s) => s.clone(),
                Value::Number(n) => { if let Some(i) = n.as_i64() { i.to_string() } else { n.as_f64().unwrap_or(0.0).to_string() } }
                Value::Bool(b) => b.to_string(),
                _ => return format!("<c r=\"{cell_ref}\"{style_attr}/>"),
            };
            let space_attr = if text.trim() != text { " xml:space=\"preserve\"" } else { "" };
            format!("<c r=\"{cell_ref}\"{style_attr} t=\"inlineStr\"><is><t{space_attr}>{}</t></is></c>", escape_xml(&text))
        }
    }
}

fn set_xml_cell(sheet_xml: &str, row_idx: usize, col_idx: usize, value: Option<&Value>, value_type: &str) -> Result<String, String> {
    let row_number = row_idx + 1;
    let cell_ref = format!("{}{}", col_name(col_idx), row_number);
    let row_pattern = format!(r#"<row\b[^>/]*\br="{row_number}"[^>/]*>[\s\S]*?</row>"#);
    let row_re = Regex::new(&row_pattern).unwrap();

    let mut xml = sheet_xml.to_string();
    if !row_re.is_match(&xml) { xml = insert_xml_row_at(&xml, row_number)?; }

    let original_row = row_re.find(&xml).map(|m| m.as_str().to_string()).unwrap();
    let cell_pattern = format!(r#"<c\b[^>/]*\br="{}"[^>/]*(?:>[\s\S]*?</c>|\s*/?>)"#, regex::escape(&cell_ref));
    let cell_re = Regex::new(&cell_pattern).unwrap();

    let is_empty_val = matches!(value, None) || matches!(value, Some(Value::Null)) || value.map(|v| v.as_str() == Some("")).unwrap_or(false);
    if is_empty_val {
        let cleared = cell_re.replace(&original_row, "").to_string();
        return Ok(xml.replacen(&original_row, &cleared, 1));
    }

    let val = value.unwrap();
    let style_id = cell_re.find(&original_row).and_then(|m| get_xml_attr(m.as_str(), "s"));
    let new_cell = build_xml_cell(&cell_ref, val, value_type, style_id.as_deref());
    let updated_row = if cell_re.is_match(&original_row) {
        cell_re.replace(&original_row, new_cell.as_str()).to_string()
    } else {
        insert_xml_cell_in_row(&original_row, &new_cell, col_idx)
    };
    Ok(xml.replacen(&original_row, &updated_row, 1))
}

fn set_xml_formula_cache_number(sheet_xml: &str, row_idx: usize, col_idx: usize, value: Option<f64>) -> Result<String, String> {
    let row_number = row_idx + 1;
    let cell_ref = format!("{}{}", col_name(col_idx), row_number);
    let row_pattern = format!(r#"<row\b[^>/]*\br="{row_number}"[^>/]*>[\s\S]*?</row>"#);
    let row_re = Regex::new(&row_pattern).unwrap();
    let original_row = match row_re.find(sheet_xml) {
        Some(m) => m.as_str().to_string(),
        None => {
            return match value {
                Some(n) => set_xml_cell(sheet_xml, row_idx, col_idx, Some(&json!(n)), "number"),
                None => Ok(sheet_xml.to_string()),
            };
        }
    };
    let cell_pattern = format!(r#"<c\b[^>/]*\br="{}"[^>/]*(?:>[\s\S]*?</c>|\s*/?>)"#, regex::escape(&cell_ref));
    let cell_re = Regex::new(&cell_pattern).unwrap();
    let cell_xml = match cell_re.find(&original_row) {
        Some(m) => m.as_str(),
        None => {
            return match value {
                Some(n) => set_xml_cell(sheet_xml, row_idx, col_idx, Some(&json!(n)), "number"),
                None => Ok(sheet_xml.to_string()),
            };
        }
    };

    if !cell_xml.contains("<f") {
        return match value {
            Some(n) => set_xml_cell(sheet_xml, row_idx, col_idx, Some(&json!(n)), "number"),
            None => set_xml_cell(sheet_xml, row_idx, col_idx, None, "number"),
        };
    }

    // Debe matchear tanto <v>valor</v> como <v/> (autocierre, valor vacio de
    // IFERROR(...,"")): si no matchea el self-closing, la rama "sin <v>" añade
    // un <v> nuevo sin quitar el existente -> celda con dos <v>, XML invalido,
    // Excel repara vaciando la hoja entera.
    let v_re = Regex::new(r#"<v\b[^>]*?(?:/>|>[\s\S]*?</v>)"#).unwrap();
    let updated_cell = match value {
        Some(n) => {
            let v = format!("<v>{}</v>", n);
            if v_re.is_match(cell_xml) {
                v_re.replace(cell_xml, v.as_str()).to_string()
            } else {
                cell_xml.replace("</c>", &format!("{v}</c>"))
            }
        }
        None => v_re.replace(cell_xml, "").to_string(),
    };
    let updated_row = original_row.replacen(cell_xml, &updated_cell, 1);
    Ok(sheet_xml.replacen(&original_row, &updated_row, 1))
}

// ---------------------------------------------------------------------------
// ZIP: editar hojas y guardar
// ---------------------------------------------------------------------------

fn find_worksheet_path_in_zip(workbook_xml: &str, rels_xml: &str, sheet_name: &str) -> Result<String, String> {
    let sheet_re = Regex::new(r#"<sheet\b[^>]*>"#).unwrap();
    for m in sheet_re.find_iter(workbook_xml) {
        let tag = m.as_str();
        if get_xml_attr(tag, "name").as_deref() != Some(sheet_name) { continue; }
        let rel_id = get_xml_attr(tag, "r:id").ok_or("No r:id en sheet tag")?;
        let rel_re = Regex::new(r#"<Relationship\b[^>]*>"#).unwrap();
        for rm in rel_re.find_iter(rels_xml) {
            if get_xml_attr(rm.as_str(), "Id").as_deref() == Some(&rel_id) {
                let target = get_xml_attr(rm.as_str(), "Target").ok_or("No Target")?;
                return Ok(if target.starts_with("xl/") { target } else { format!("xl/{target}") });
            }
        }
    }
    Err(format!("No se encontro la hoja \"{sheet_name}\" dentro del libro."))
}

// Lee los valores cacheados de una columna directamente del XML del ZIP
// (calamine puede no alcanzar columnas lejanas si el rango detectado es corto)
fn read_col_values_from_xml(path: &str, sheet_name: &str, col: usize) -> HashMap<usize, String> {
    let mut result = HashMap::new();
    let input = match std::fs::read(path) { Ok(b) => b, Err(_) => return result };
    let cursor = std::io::Cursor::new(input);
    let mut zip = match zip::ZipArchive::new(cursor) { Ok(z) => z, Err(_) => return result };
    let mut files: HashMap<String, Vec<u8>> = HashMap::new();
    for i in 0..zip.len() {
        let mut f = match zip.by_index(i) { Ok(f) => f, Err(_) => continue };
        let name = f.name().to_string();
        if !name.contains("workbook") && !name.contains("worksheet") { continue; }
        let mut buf = Vec::new();
        let _ = f.read_to_end(&mut buf);
        files.insert(name, buf);
    }
    let wb = match files.get("xl/workbook.xml") { Some(b) => String::from_utf8_lossy(b).to_string(), None => return result };
    let rels = match files.get("xl/_rels/workbook.xml.rels") { Some(b) => String::from_utf8_lossy(b).to_string(), None => return result };
    let sheet_path = match find_worksheet_path_in_zip(&wb, &rels, sheet_name) { Ok(p) => p, Err(_) => return result };
    let xml = match files.get(&sheet_path) { Some(b) => String::from_utf8_lossy(b).to_string(), None => return result };
    let col_str = col_name(col);
    let pattern = format!(r#"<c\b[^>]*\br="{}(\d+)"[^>]*>[\s\S]*?<v>([^<]+)</v>"#, regex::escape(&col_str));
    let re = match Regex::new(&pattern) { Ok(r) => r, Err(_) => return result };
    for cap in re.captures_iter(&xml) {
        let row_1: usize = match cap[1].parse() { Ok(n) => n, Err(_) => continue };
        let val = cap[2].trim().to_string();
        if !val.is_empty() { result.insert(row_1 - 1, val); }
    }
    result
}

// Lee todos los valores de texto de una fila (1-indexed) directamente del ZIP/XML
// Devuelve mapa col_0indexed -> valor_string (con shared strings resueltos)
fn parse_shared_strings_xml(xml: &str) -> Vec<String> {
    // Extraer todos los <si>...</si> usando posición de string (sin regex multilinea)
    let mut shared = Vec::new();
    let mut pos = 0;
    while let Some(si_start) = xml[pos..].find("<si>") {
        let abs_start = pos + si_start + 4; // justo después de <si>
        let end = match xml[abs_start..].find("</si>") { Some(e) => abs_start + e, None => break };
        let si_content = &xml[abs_start..end];
        // Concatenar todos los <t>...</t> dentro del <si>
        let mut text = String::new();
        let mut p2 = 0;
        while let Some(t_start) = si_content[p2..].find("<t") {
            let abs_t = p2 + t_start;
            let gt = match si_content[abs_t..].find('>') { Some(g) => abs_t + g + 1, None => break };
            let t_end = match si_content[gt..].find("</t>") { Some(e) => gt + e, None => break };
            text.push_str(&si_content[gt..t_end]);
            p2 = t_end + 4;
        }
        shared.push(text);
        pos = end + 5; // después de </si>
    }
    shared
}

fn extract_cell_value(cell_xml: &str, shared: &[String]) -> Option<String> {
    // cell_xml: contenido entre <c ...> y </c> (sin las etiquetas)
    let is_shared = cell_xml.contains("t=\"s\"") || cell_xml.contains("t='s'");
    let is_inline = cell_xml.contains("t=\"inlineStr\"") || cell_xml.contains("t='inlineStr'");
    if is_inline {
        let t_start = cell_xml.find("<t")?;
        let t_gt = cell_xml[t_start..].find('>')? + t_start + 1;
        let t_end = cell_xml[t_gt..].find("</t>")? + t_gt;
        return Some(cell_xml[t_gt..t_end].to_string());
    }
    // Buscar <v>...</v>
    let v_start = cell_xml.find("<v>")? + 3;
    let v_end = cell_xml[v_start..].find("</v>")? + v_start;
    let raw = cell_xml[v_start..v_end].trim();
    if is_shared {
        let idx: usize = raw.parse().ok()?;
        Some(shared.get(idx).cloned().unwrap_or_default())
    } else {
        Some(raw.to_string())
    }
}

fn read_row_from_xml(path: &str, sheet_name: &str, row_1: usize) -> HashMap<usize, String> {
    let mut result = HashMap::new();
    let input = match std::fs::read(path) { Ok(b) => b, Err(_) => return result };
    let cursor = std::io::Cursor::new(input);
    let mut zip = match zip::ZipArchive::new(cursor) { Ok(z) => z, Err(_) => return result };
    let mut files: HashMap<String, Vec<u8>> = HashMap::new();
    for i in 0..zip.len() {
        let mut f = match zip.by_index(i) { Ok(f) => f, Err(_) => continue };
        let name = f.name().to_string();
        if !name.contains("workbook") && !name.contains("worksheet") && !name.contains("sharedStrings") { continue; }
        let mut buf = Vec::new();
        let _ = f.read_to_end(&mut buf);
        files.insert(name, buf);
    }
    let shared: Vec<String> = if let Some(b) = files.get("xl/sharedStrings.xml") {
        parse_shared_strings_xml(&String::from_utf8_lossy(b))
    } else { vec![] };
    let wb = match files.get("xl/workbook.xml") { Some(b) => String::from_utf8_lossy(b).to_string(), None => return result };
    let rels = match files.get("xl/_rels/workbook.xml.rels") { Some(b) => String::from_utf8_lossy(b).to_string(), None => return result };
    let sheet_path = match find_worksheet_path_in_zip(&wb, &rels, sheet_name) { Ok(p) => p, Err(_) => return result };
    let xml = match files.get(&sheet_path) { Some(b) => String::from_utf8_lossy(b).to_string(), None => return result };
    // Encontrar la fila r="N" usando búsqueda de string (no regex — el XML puede tener newlines)
    let row_attr = format!("r=\"{}\"", row_1);
    let row_attr2 = format!("r='{}'", row_1); // por si acaso
    let mut pos = 0;
    let row_content = loop {
        let found = match xml[pos..].find("<row ") { Some(f) => pos + f, None => return result };
        let tag_gt = match xml[found..].find('>') { Some(g) => found + g, None => return result };
        let tag = &xml[found..=tag_gt];
        let has_attr = tag.contains(&row_attr) || tag.contains(&row_attr2);
        if has_attr {
            // Verificar que no es r="10" cuando buscamos r="1"
            let re_check = Regex::new(&format!(r#"\br="{}""#, row_1)).unwrap();
            if re_check.is_match(tag) {
                let content_start = tag_gt + 1;
                let content_end = match xml[content_start..].find("</row>") { Some(e) => content_start + e, None => return result };
                break xml[content_start..content_end].to_string();
            }
        }
        pos = found + 1;
    };
    // Parsear celdas — iterar sobre <c ...>...</c> usando string search
    // Maneja celdas self-closing (<c r="X4"/>) y celdas con contenido (<c r="X4">...</c>)
    let re_cell_ref = Regex::new(r#"\br="([A-Z]+)\d+""#).unwrap();
    let mut cpos = 0;
    while let Some(c_start) = row_content[cpos..].find("<c ") {
        let abs_c = cpos + c_start;
        let c_gt = match row_content[abs_c..].find('>') { Some(g) => abs_c + g, None => break };
        let c_tag = &row_content[abs_c..=c_gt];
        // Extraer referencia de celda (ej: B3, DG4)
        let col_str = match re_cell_ref.captures(c_tag) {
            Some(cap) => cap[1].to_string(),
            None => { cpos = c_gt + 1; continue; }
        };
        // Detectar celda self-closing (<c ... />) — no tiene contenido
        let is_self_closing = c_tag.ends_with("/>") || c_tag.ends_with("/ >");
        if is_self_closing {
            cpos = c_gt + 1;
            continue;
        }
        // Encontrar cierre </c> — si no existe, continuar (no hacer break)
        let cell_end = match row_content[c_gt+1..].find("</c>") {
            Some(e) => c_gt + 1 + e,
            None => { cpos = c_gt + 1; continue; } // <-- continuar, no break
        };
        let cell_content = &row_content[c_gt+1..cell_end];
        let val = match extract_cell_value(&format!("{}{}", c_tag, cell_content), &shared) {
            Some(v) if !v.is_empty() => v,
            _ => { cpos = cell_end + 4; continue; }
        };
        let col_idx = col_name_to_idx(&col_str);
        result.insert(col_idx, val);
        cpos = cell_end + 4;
    }
    result
}

fn col_name_to_idx(name: &str) -> usize {
    name.chars().fold(0usize, |acc, c| acc * 26 + (c as usize - 'A' as usize + 1)) - 1
}

fn assert_worksheet_xml_safe(sheet_xml: &str, sheet_name: &str) -> Result<(), String> {
    let row_re = Regex::new(r#"<row\b[^>/]*(?:/>|>[\s\S]*?</row>)"#).unwrap();
    for m in row_re.find_iter(sheet_xml) {
        let row_xml = m.as_str();
        let row_num = get_xml_attr(row_xml, "r").unwrap_or_default();
        let open2 = Regex::new(r"<c\b").unwrap().find_iter(row_xml).count();
        let closed = row_xml.matches("</c>").count();
        let self_closed = Regex::new(r"<c\b[^>/]*/>").unwrap().find_iter(row_xml).count();
        if open2 != closed + self_closed {
            return Err(format!("El XML generado para {sheet_name}, fila {row_num}, no es seguro."));
        }
    }
    Ok(())
}

fn force_workbook_recalc(xml: &str) -> String {
    let apply = |tag: &str| -> String {
        let t = set_xml_attr(tag, "calcMode", "auto");
        let t = set_xml_attr(&t, "fullCalcOnLoad", "1");
        set_xml_attr(&t, "forceFullCalc", "1")
    };
    let self_re = Regex::new(r"<calcPr\b[^>]*/\s*>").unwrap();
    let open_re = Regex::new(r"<calcPr\b[^>]*>").unwrap();
    if self_re.is_match(xml) {
        self_re.replace(xml, |caps: &regex::Captures| apply(&caps[0])).to_string()
    } else if open_re.is_match(xml) {
        open_re.replace(xml, |caps: &regex::Captures| apply(&caps[0])).to_string()
    } else {
        xml.replace("</workbook>", "<calcPr calcMode=\"auto\" fullCalcOnLoad=\"1\" forceFullCalc=\"1\"/></workbook>")
    }
}

fn edit_workbook_sheets_xml(path: &str, sheet_edits: Vec<(&str, Box<dyn Fn(&str) -> Result<String, String>>)>) -> Result<(), String> {
    let input = std::fs::read(path).map_err(|e| e.to_string())?;
    let backup_bytes = input.clone();
    let cursor = std::io::Cursor::new(input);
    let mut zip = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let mut files: HashMap<String, Vec<u8>> = HashMap::new();
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).map_err(|e| e.to_string())?;
        let name = f.name().to_string();
        let mut buf = Vec::new();
        f.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        files.insert(name, buf);
    }

    let workbook_xml = String::from_utf8_lossy(files.get("xl/workbook.xml").ok_or("No xl/workbook.xml")?).to_string();
    let rels_xml = String::from_utf8_lossy(files.get("xl/_rels/workbook.xml.rels").ok_or("No workbook.xml.rels")?).to_string();

    for (sheet_name, edit_fn) in &sheet_edits {
        let sheet_path = find_worksheet_path_in_zip(&workbook_xml, &rels_xml, sheet_name)?;
        let original = String::from_utf8_lossy(files.get(&sheet_path).ok_or_else(|| format!("No se encontro {sheet_path}"))?).to_string();
        let updated = edit_fn(&original)?;
        assert_worksheet_xml_safe(&updated, sheet_name)?;
        files.insert(sheet_path, updated.into_bytes());
    }

    files.remove("xl/calcChain.xml");
    if let Some(b) = files.get_mut("xl/_rels/workbook.xml.rels") {
        let s = String::from_utf8_lossy(b).to_string();
        *b = Regex::new(r#"<Relationship\b[^>]*Type="[^"]*/calcChain"[^>]*(?:/>|></Relationship>)"#).unwrap().replace_all(&s, "").into_owned().into_bytes();
    }
    if let Some(b) = files.get_mut("[Content_Types].xml") {
        let s = String::from_utf8_lossy(b).to_string();
        *b = Regex::new(r#"<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*(?:/>|></Override>)"#).unwrap().replace_all(&s, "").into_owned().into_bytes();
    }
    if let Some(b) = files.get_mut("xl/workbook.xml") {
        let s = String::from_utf8_lossy(b).to_string();
        *b = force_workbook_recalc(&s).into_bytes();
    }

    let names_ordered: Vec<String> = {
        let input2 = std::fs::read(path).map_err(|e| e.to_string())?;
        let mut z2 = zip::ZipArchive::new(std::io::Cursor::new(input2)).map_err(|e| e.to_string())?;
        (0..z2.len()).filter_map(|i| z2.by_index_raw(i).ok().map(|f| f.name().to_string())).collect()
    };
    let mut final_order: Vec<String> = names_ordered.into_iter().filter(|n| n != "xl/calcChain.xml").collect();
    for k in files.keys() {
        if !final_order.contains(k) { final_order.push(k.clone()); }
    }

    let out_buf = Vec::new();
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(out_buf));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6));
    for name in &final_order {
        if let Some(data) = files.get(name) {
            zw.start_file(name, options).map_err(|e| e.to_string())?;
            zw.write_all(data).map_err(|e| e.to_string())?;
        }
    }
    let out_cursor = zw.finish().map_err(|e| e.to_string())?;
    // Backup best-effort del estado previo a esta escritura: si la escritura nueva
    // sale corrupta (como el bug del <v/> duplicado), el usuario tiene un ".bak"
    // del que recuperar. No aborta el guardado si falla (p.ej. disco lleno).
    let _ = std::fs::write(format!("{path}.bak"), &backup_bytes);
    std::fs::write(path, out_cursor.into_inner()).map_err(|e| e.to_string())?;
    Ok(())
}

// ---------------------------------------------------------------------------
// save_alumnos_to_file
// ---------------------------------------------------------------------------

fn save_alumnos_to_file(path: &str, alumnos: &[Value]) -> Result<(), String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let header_row = find_header_row(&rows, "ALUMNADO").ok_or("No se encontro la seccion ALUMNADO.")?;
    let alumnos_start = header_row + 1;
    let existing_count = (alumnos_start..rows.len()).take_while(|&i| !cell_str(&rows, i, 1).is_empty()).count();
    let rows_to_clear = existing_count.max(alumnos.len());
    let alumnos_owned = alumnos.to_vec();

    edit_workbook_sheets_xml(path, vec![("DATOS", Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for idx in 0..rows_to_clear {
            let ri = alumnos_start + idx;
            s = set_xml_cell(&s, ri, 0, None, "text")?;
            s = set_xml_cell(&s, ri, 1, None, "text")?;
        }
        for (idx, alumno) in alumnos_owned.iter().enumerate() {
            let ri = alumnos_start + idx;
            s = set_xml_cell(&s, ri, 0, Some(&alumno["numero"]), "number")?;
            s = set_xml_cell(&s, ri, 1, Some(&alumno["nombre"]), "text")?;
        }
        Ok(s)
    }))])
}

// ---------------------------------------------------------------------------
// save_unidades_to_file
// ---------------------------------------------------------------------------

// En ESO no hay bloques de evaluación separados en DATOS — sólo la tabla principal
fn save_unidades_to_file(path: &str, unidades: &[Value]) -> Result<(), String> {
    // Tabla fija I5:K20 (0-indexed: filas 4-19, cols 8=I, 9=J, 10=K)
    let unidades_owned = unidades.to_vec();
    edit_workbook_sheets_xml(path, vec![
        ("DATOS", Box::new({
            let unidades_datos = unidades_owned.clone();
            move |xml: &str| {
                let mut s = xml.to_string();
                for idx in 0..16 {
                    let u = unidades_datos.get(idx);
                    let ri = 4 + idx;
                    let codigo = u.and_then(|v| v["codigo"].as_str()).unwrap_or("");
                    let codigo_val = json!(codigo);
                    s = set_xml_cell(&s, ri, 8, if codigo.is_empty() { None } else { Some(&codigo_val) }, "text")?;
                    let nombre = u.and_then(|v| v["nombre"].as_str()).unwrap_or("");
                    let nombre_val = json!(nombre);
                    s = set_xml_cell(&s, ri, 9, if nombre.is_empty() { None } else { Some(&nombre_val) }, "text")?;
                    let eval = u.and_then(|v| v["evaluacion"].as_str()).unwrap_or("");
                    let eval_val = json!(eval);
                    s = set_xml_cell(&s, ri, 10, if eval.is_empty() { None } else { Some(&eval_val) }, "text")?;
                }
                Ok(s)
            }
        })),
        ("PESOS", Box::new({
            let unidades_pesos = unidades_owned.clone();
            move |xml: &str| {
                let mut s = xml.to_string();
                for idx in 0..16 {
                    let u = unidades_pesos.get(idx);
                    let ri = 4 + idx;
                    let nombre = u.and_then(|v| v["nombre"].as_str()).unwrap_or("");
                    let nombre_val = json!(nombre);
                    s = set_xml_cell(&s, ri, 0, if nombre.is_empty() { None } else { Some(&nombre_val) }, "text")?;
                }
                Ok(s)
            }
        })),
    ])
}

// ---------------------------------------------------------------------------
// save_rraa_criterios_to_file (ESO: CE + CR en PESOS)
// ---------------------------------------------------------------------------

fn save_rraa_criterios_to_file(_rraa: &[Value], criterios: &[Value], pond_unidad: &[Value], path: &str) -> Result<(), String> {
    let criterios_owned = criterios.to_vec();
    let pond_owned = pond_unidad.to_vec();

    // Mapa código CR → fila en DATOS (col W=22) para persistir el texto (col X=23)
    let mut cr_datos_row: HashMap<String, usize> = HashMap::new();
    if let Ok(datos_rows) = read_sheet_rows(path, "DATOS") {
        for row_idx in 0..datos_rows.len() {
            let code = cell_str(&datos_rows, row_idx, 22);
            if is_cr_code(&code) {
                cr_datos_row.insert(normalize_criterion_code(&code), row_idx);
            }
        }
    }

    let mut code_col_map: HashMap<String, usize> = HashMap::new();
    if let Ok(rows) = read_sheet_rows(path, "PESOS") {
        if rows.len() > 3 {
            let mut last_cr_by_ce: HashMap<i64, i64> = HashMap::new();
            for ci in 0..rows[3].len() {
                let code = cell_str(&rows, 3, ci);
                if let Some((_, _, fixed_code)) = next_cr_code_for_ce(&code, &mut last_cr_by_ce) {
                    code_col_map.insert(normalize_criterion_code(&fixed_code), ci);
                }
            }
        }
    }

    edit_workbook_sheets_xml(path, vec![
        ("PESOS", Box::new({
            let criterios2 = criterios_owned.clone(); let pond2 = pond_owned.clone(); let code_col_map2 = code_col_map.clone();
            move |xml: &str| {
                let mut s = xml.to_string();
                // Mapa colIdx secuencial → columna real en PESOS
                let col_map: std::collections::HashMap<usize, usize> = criterios2.iter()
                    .filter_map(|c| {
                        let ci = c["colIdx"].as_u64()? as usize;
                        let by_actual = c["actualCol"].as_u64().map(|n| n as usize).filter(|&n| n > 0);
                        let by_code = c["codigo"].as_str()
                            .and_then(|code| code_col_map2.get(&normalize_criterion_code(code)).copied());
                        let actual = by_actual.or(by_code).unwrap_or(ci);
                        if actual > 0 { Some((ci, actual)) } else { None }
                    })
                    .collect();
                // Fila 4 (idx 3): escribir códigos CR en su columna real
                for criterio in &criterios2 {
                    if let Some(ac) = criterio["actualCol"].as_u64().map(|n| n as usize)
                        .filter(|&n| n > 0)
                        .or_else(|| criterio["codigo"].as_str()
                            .and_then(|code| code_col_map2.get(&normalize_criterion_code(code)).copied()))
                        .or_else(|| criterio["colIdx"].as_u64().map(|n| n as usize)) {
                        if ac > 0 {
                            s = set_xml_cell(&s, 3, ac, Some(&criterio["codigo"]), "text")?;
                        }
                    }
                }
                // Filas de unidades: ponderaciones por CR e instrumento por CE
                for unidad in &pond2 {
                    if let Some(ri) = unidad["rowIdx"].as_u64().map(|n| n as usize) {
                        let nombre = unidad["nombre"].as_str().unwrap_or("");
                        let nombre_val = json!(nombre);
                        s = set_xml_cell(&s, ri, 0, if nombre.is_empty() { None } else { Some(&nombre_val) }, "text")?;
                        if let Some(ponds) = unidad["ponderaciones"].as_object() {
                            for (col_key, vals) in ponds {
                                let ci: usize = col_key.parse().unwrap_or(0);
                                if let Some(&actual) = col_map.get(&ci) {
                                    if actual > 0 {
                                        s = set_xml_cell(&s, ri, actual, Some(&json!(parse_decimal(&vals["ponderacion"]))), "number")?;
                                    }
                                }
                            }
                        }
                        if let Some(instrs) = unidad["instrPorCe"].as_object() {
                            for (_ce_num, val) in instrs {
                                if let (Some(instr_str), Some(col_idx)) = (val["codigo"].as_str(), val["colIdx"].as_u64()) {
                                    let ci = col_idx as usize;
                                    if ci > 0 {
                                        s = set_xml_cell(&s, ri, ci, Some(&json!(instr_str)), "text")?;
                                    }
                                }
                            }
                        }
                    }
                }
                Ok(s)
            }
        })),
        ("DATOS", Box::new({
            let criterios3 = criterios_owned.clone(); let cr_datos_row2 = cr_datos_row.clone();
            move |xml: &str| {
                let mut s = xml.to_string();
                for criterio in &criterios3 {
                    if let Some(codigo) = criterio["codigo"].as_str() {
                        if let Some(&row_idx) = cr_datos_row2.get(&normalize_criterion_code(codigo)) {
                            let texto = criterio["texto"].as_str().unwrap_or("");
                            let texto_val = json!(texto);
                            s = set_xml_cell(&s, row_idx, 23, if texto.is_empty() { None } else { Some(&texto_val) }, "text")?;
                        }
                    }
                }
                Ok(s)
            }
        })),
    ])
}

// ---------------------------------------------------------------------------
// save_notas_actividad
// ---------------------------------------------------------------------------

// Replica SUMPRODUCT((vals<>"")*pesos*vals)/SUMPRODUCT((vals<>"")*pesos) de la
// plantilla: media ponderada de hasta 4 instrumentos, ignorando los que estan
// vacios (None) tanto en el numerador como en el denominador. Un valor 0
// explicito SI cuenta (distinto de vacio, igual que en Excel "0" <> "").
// Temporary: only called from test cfg until Task 3 wires it into evaluation sheet load.
#[allow(dead_code)]
fn compute_final_weighted(values: [Option<f64>; 4], weights: [f64; 4]) -> Option<f64> {
    let mut num = 0.0;
    let mut den = 0.0;
    for i in 0..4 {
        if let Some(v) = values[i] {
            num += weights[i] * v;
            den += weights[i];
        }
    }
    if den == 0.0 { None } else { Some(num / den) }
}

fn normalize_grade(value: &Value) -> Option<f64> {
    match value {
        Value::Number(n) => n.as_f64(),
        Value::String(s) => { let s2 = s.replace(',', "."); if s2.trim().is_empty() { None } else { s2.trim().parse().ok() } }
        _ => None,
    }
}

#[tauri::command]
async fn excel_save_notas_actividad(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_notas_actividad_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_notas_actividad_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let tipo = payload["tipo"].as_str().ok_or("Falta tipo")?.to_string();
    let actividad = payload["actividad"].as_i64().unwrap_or(1);
    let notas = payload["notas"].as_array().ok_or("Falta notas")?.clone();
    let nombre_actividad = payload["nombreActividad"].as_str().map(|s| s.to_string());
    let incluida = payload["incluida"].as_bool().unwrap_or(false);

    let rows = read_sheet_rows(&path, &unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(&tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let block = blocks.iter().find(|b| b.numero == actividad)
        .ok_or_else(|| format!("No se encontro la actividad {actividad} de {} en {unidad}.", at.label))?.clone();

    let path_clone = path.clone(); let unidad_clone = unidad.clone(); let tipo_clone = tipo.clone();

    edit_workbook_sheets_xml(&path, vec![(&unidad, Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        let nv = nombre_actividad.as_deref().map(|n| json!(n));
        s = set_xml_cell(&s, block.number_row, block.name_value_col, nv.as_ref(), "text")?;
        let x_val = json!("X");
        s = set_xml_cell(&s, block.included_row, block.included_col, if incluida { Some(&x_val) } else { None }, "text")?;
        for nota_item in &notas {
            if let Some(ri) = nota_item["rowIdx"].as_u64().map(|n| n as usize) {
                if ri < block.first_student_row { continue; }
                if cell_str(&rows, ri, block.name_col).is_empty() { continue; }
                match normalize_grade(&nota_item["nota"]) {
                    Some(n) => { s = set_xml_cell(&s, ri, block.note_col, Some(&json!(n)), "number")?; }
                    None    => { s = set_xml_cell(&s, ri, block.note_col, None, "number")?; }
                }
                // Guardar notas CE por alumno si vienen en el payload
                if let Some(ce_notas) = nota_item["ceNotas"].as_object() {
                    for (code, ci) in &block.ce_cols {
                        if let Some(val) = ce_notas.get(code) {
                            match normalize_grade(val) {
                                Some(n) => { s = set_xml_cell(&s, ri, *ci, Some(&json!(n)), "number")?; }
                                None    => { s = set_xml_cell(&s, ri, *ci, None, "number")?; }
                            }
                        }
                    }
                }
            }
        }
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    load_notas_actividad(&path_clone, &unidad_clone, &tipo_clone, actividad, None)
}

// ---------------------------------------------------------------------------
// excel_save_notas_unidad — ESO
// ---------------------------------------------------------------------------

fn find_evaluation_layout_indices(rows: &[Vec<Value>]) -> Option<(usize, usize, usize)> {
    for row_idx in 0..rows.len() {
        let has_nota_ce = rows[row_idx].iter().any(|v| is_nota_ce_cell(v));
        let cr_count = rows[row_idx].iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count();
        if has_nota_ce && cr_count >= 2 {
            return Some((row_idx, row_idx, row_idx + 2));
        }
    }

    for row_idx in 0..rows.len() {
        if rows[row_idx].iter().filter(|v| is_nota_ce_cell(v)).count() == 0 { continue; }
        for offset in 1..=3 {
            if let Some(next) = rows.get(row_idx + offset) {
                if next.iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count() >= 2 {
                    return Some((row_idx, row_idx + offset, row_idx + offset + 1));
                }
            }
        }
    }

    for row_idx in 1..rows.len() {
        if rows[row_idx].iter().filter(|v| is_eval_criterion_code(&cell_val_str(v))).count() >= 2 {
            return Some((row_idx.saturating_sub(1), row_idx, row_idx + 1));
        }
    }

    None
}

struct FieldUpdate { nota: Option<Option<f64>>, rec: Option<Option<f64>> }

fn build_eval_sheet_edits(path: &str, unidad: &str, notas: &[Value]) -> Result<Vec<(String, Box<dyn Fn(&str) -> Result<String, String>>)>, String> {
    let mut wb = open_workbook_auto(path).map_err(|e| e.to_string())?;
    let unit_rows = read_sheet_rows_from_wb(&mut wb, unidad)
        .map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let mut updates: HashMap<(String, String), FieldUpdate> = HashMap::new();

    for nota_item in notas {
        let Some(ri) = nota_item["rowIdx"].as_u64().map(|n| n as usize) else { continue; };
        let alumno = normalize_plain(&cell_str(&unit_rows, ri, 0));
        if alumno.is_empty() { continue; }
        let Some(cr_notas) = nota_item["crNotas"].as_object() else { continue; };
        for (code, val_obj) in cr_notas {
            let code_norm = normalize_plain(code);
            let entry = updates.entry((alumno.clone(), code_norm)).or_insert(FieldUpdate { nota: None, rec: None });
            if val_obj.get("nota").is_some() {
                entry.nota = Some(val_obj.get("nota").and_then(normalize_grade));
            }
            if val_obj.get("rec").is_some() {
                entry.rec = Some(val_obj.get("rec").and_then(normalize_grade));
            }
        }
    }

    if updates.is_empty() { return Ok(Vec::new()); }

    let names = sheet_names_from_wb(&wb);
    let eval_sheets: Vec<String> = names.into_iter().filter(|name| {
        let norm = normalize_plain(name);
        (norm.contains("EVA") || norm == "FINAL") && !norm.contains("MAX")
    }).collect();

    let mut edits: Vec<(String, Box<dyn Fn(&str) -> Result<String, String>>)> = Vec::new();

    for sheet_name in eval_sheets {
        let rows = match read_sheet_rows_from_wb(&mut wb, &sheet_name) {
            Ok(rows) => rows,
            Err(_) => continue,
        };
        let Some((_, code_row_idx, first_student_row_idx)) = find_evaluation_layout_indices(&rows) else { continue; };
        let code_row = rows.get(code_row_idx).cloned().unwrap_or_default();
        let code_cols: Vec<(usize, String)> = code_row.iter().enumerate().filter_map(|(ci, value)| {
            let code = cell_val_str(value);
            if is_eval_criterion_code(&code) {
                Some((ci, normalize_plain(&code)))
            } else {
                None
            }
        }).collect();
        if code_cols.is_empty() { continue; }

        let mut cells: Vec<(usize, usize, Option<f64>)> = Vec::new();
        for row_idx in first_student_row_idx..rows.len() {
            let alumno = normalize_plain(&cell_str(&rows, row_idx, 0));
            if alumno.is_empty() { continue; }
            for (col_idx, code_norm) in &code_cols {
                if let Some(update) = updates.get(&(alumno.clone(), code_norm.clone())) {
                    if let Some(nota_val) = update.nota {
                        cells.push((row_idx, *col_idx, nota_val));
                    }
                    if let Some(rec_val) = update.rec {
                        cells.push((row_idx, col_idx + 1, rec_val));
                    }
                }
            }
        }
        if cells.is_empty() { continue; }

        edits.push((sheet_name, Box::new(move |xml: &str| {
            let mut s = xml.to_string();
            for (row_idx, col_idx, value) in &cells {
                s = set_xml_formula_cache_number(&s, *row_idx, *col_idx, *value)?;
            }
            Ok(s)
        }) as Box<dyn Fn(&str) -> Result<String, String>>));
    }

    Ok(edits)
}

// Comando async: el guardado reescribe el zip completo del xlsx (varios MB, varias
// hojas de evaluacion) y puede tardar; si el comando no es async, Tauri lo ejecuta
// en el hilo principal y congela toda la ventana mientras dura el guardado.
#[tauri::command]
async fn excel_save_notas_unidad(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_notas_unidad_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_notas_unidad_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let notas = payload["notas"].as_array().ok_or("Falta notas")?.clone();
    // syncEval=false: guardado ligero (1 celda, sin recalcular caché de evaluación ni
    // releer la unidad entera) — usado en el guardado inmediato por celda para no
    // bloquear la escritura mientras el usuario sigue tecleando.
    let sync_eval = payload["syncEval"].as_bool().unwrap_or(true);
    let notas_for_unit = notas.clone();
    let notas_for_eval = notas.clone();

    let unit_edit_fn: Box<dyn Fn(&str) -> Result<String, String>> = Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for nota_item in &notas_for_unit {
            if let Some(ri) = nota_item["rowIdx"].as_u64().map(|n| n as usize) {
                if let Some(cr_notas) = nota_item["crNotas"].as_object() {
                    for (_code, val_obj) in cr_notas {
                        if let Some(ci) = val_obj.get("colIdx").and_then(|v| v.as_u64()).map(|n| n as usize) {
                            if let Some(val) = val_obj.get("nota") {
                                match normalize_grade(val) {
                                    Some(n) => { s = set_xml_cell(&s, ri, ci, Some(&json!(n)), "number")?; }
                                    None    => { s = set_xml_cell(&s, ri, ci, None, "number")?; }
                                }
                            }
                            if let Some(val) = val_obj.get("rec") {
                                match normalize_grade(val) {
                                    Some(n) => { s = set_xml_cell(&s, ri, ci + 1, Some(&json!(n)), "number")?; }
                                    None    => { s = set_xml_cell(&s, ri, ci + 1, None, "number")?; }
                                }
                            }
                        }
                    }
                }
            }
        }
        Ok(s)
    });

    // El guardado del propio Rec/nota en la hoja de unidad va SIEMPRE primero y por
    // separado: debe quedar en disco aunque la sincronizacion con las hojas de
    // evaluacion (mas fragil: varias hojas, layouts variables) falle.
    edit_workbook_sheets_xml(&path, vec![(unidad.as_str(), unit_edit_fn)])?;

    if !sync_eval {
        return Ok(Value::Null);
    }

    // Hojas de evaluacion afectadas: se editan y se escriben en UNA sola pasada de
    // zip (antes se repetia una vez por hoja, bloqueando el autoguardado). Un fallo
    // aqui no debe perder el guardado de la unidad, que ya esta en disco.
    let eval_edits = build_eval_sheet_edits(&path, &unidad, &notas_for_eval)?;
    if !eval_edits.is_empty() {
        let (eval_names, eval_fns): (Vec<String>, Vec<Box<dyn Fn(&str) -> Result<String, String>>>) = eval_edits.into_iter().unzip();
        let all_eval_edits: Vec<(&str, Box<dyn Fn(&str) -> Result<String, String>>)> = eval_names.iter()
            .zip(eval_fns.into_iter())
            .map(|(name, f)| (name.as_str(), f))
            .collect();
        edit_workbook_sheets_xml(&path, all_eval_edits)?;
    }

    load_notas_unidad(&path, &unidad)
}

// Vuelve a propagar TODAS las notas/Rec ya guardados en una hoja de unidad hacia
// las hojas de evaluacion. Existe porque, durante un tiempo, el autoguardado de
// recuperaciones no disparaba esa sincronizacion (bug ya corregido): los datos
// quedaron bien guardados en la hoja de unidad pero nunca llegaron a la hoja de
// evaluacion. Este comando repara ese historial sin que el usuario tenga que
// re-teclear cada nota.
#[tauri::command]
async fn excel_resync_unidad_eval(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_resync_unidad_eval_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_resync_unidad_eval_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();

    let unit_data = load_notas_unidad(&path, &unidad)?;
    let alumnos = unit_data["alumnos"].as_array().cloned().unwrap_or_default();

    let notas: Vec<Value> = alumnos.iter().filter_map(|a| {
        let row_idx = a["rowIdx"].as_u64()?;
        let cr_notas = a["crNotas"].as_array()?;
        let mut cr_obj = serde_json::Map::new();
        for cr in cr_notas {
            let codigo = cr["codigo"].as_str()?.to_string();
            let col_idx = cr["colIdx"].as_u64()?;
            let mut entry = serde_json::Map::new();
            entry.insert("colIdx".to_string(), json!(col_idx));
            if let Some(n) = cr["nota"].as_f64() { entry.insert("nota".to_string(), json!(n)); }
            let rec_str = cr["recDisplay"].as_str().unwrap_or("").trim().replace(',', ".");
            if let Ok(v) = rec_str.parse::<f64>() { entry.insert("rec".to_string(), json!(v)); }
            cr_obj.insert(codigo, Value::Object(entry));
        }
        Some(json!({ "rowIdx": row_idx, "crNotas": Value::Object(cr_obj) }))
    }).collect();

    let eval_edits = build_eval_sheet_edits(&path, &unidad, &notas)?;
    let hojas_afectadas = eval_edits.len();
    if !eval_edits.is_empty() {
        let (eval_names, eval_fns): (Vec<String>, Vec<Box<dyn Fn(&str) -> Result<String, String>>>) = eval_edits.into_iter().unzip();
        let all_eval_edits: Vec<(&str, Box<dyn Fn(&str) -> Result<String, String>>)> = eval_names.iter()
            .zip(eval_fns.into_iter())
            .map(|(name, f)| (name.as_str(), f))
            .collect();
        edit_workbook_sheets_xml(&path, all_eval_edits)?;
    }

    Ok(json!({ "unidad": unidad, "alumnos": notas.len(), "hojasAfectadas": hojas_afectadas }))
}

// ---------------------------------------------------------------------------
// add_actividad
// ---------------------------------------------------------------------------

const ACTIVITY_BLOCK_ROWS: usize = 41;
const ACTIVITY_BLOCK_STRIDE: usize = 44;

fn shift_formula_rows(formula: &str, row_delta: i64, src_start: i64, src_end: i64) -> String {
    Regex::new(r"(\$?[A-Z]{1,3})(\$?)(\d+)").unwrap().replace_all(formula, |caps: &regex::Captures| {
        let col = &caps[1]; let abs_row = &caps[2]; let row_num: i64 = caps[3].parse().unwrap_or(0);
        if !abs_row.is_empty() || row_num < src_start || row_num > src_end { caps[0].to_string() }
        else { format!("{}{}", col, row_num + row_delta) }
    }).to_string()
}

fn shift_formula_ref_attr(tag: &str, row_delta: i64, src_start: i64, src_end: i64) -> String {
    if let Some(ref_val) = get_xml_attr(tag, "ref") {
        set_xml_attr(tag, "ref", &shift_formula_rows(&ref_val, row_delta, src_start, src_end))
    } else { tag.to_string() }
}

fn extract_and_adjust_cells(source_row_xml: &str, dst_row: i64, row_delta: i64, src_start: i64, src_end: i64, col_start: usize, col_end: usize) -> Vec<String> {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    let mut cells = Vec::new();
    for m in cell_re.find_iter(source_row_xml) {
        let cell_xml = m.as_str();
        let ref_attr = match get_xml_attr(cell_xml, "r") { Some(r) => r, None => continue };
        let col_part: String = ref_attr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let ci = col_index(&col_part);
        if ci < col_start || ci > col_end { continue; }
        let new_ref = format!("{}{}", col_part, dst_row);
        let mut adjusted = cell_xml.replacen(&format!(r#"r="{ref_attr}""#), &format!(r#"r="{new_ref}""#), 1);
        adjusted = Regex::new(r#"(<f\b[^>/]*>)([\s\S]*?)(</f>)"#).unwrap().replace_all(&adjusted, |caps: &regex::Captures| {
            format!("{}{}{}", shift_formula_ref_attr(&caps[1], row_delta, src_start, src_end), shift_formula_rows(&caps[2], row_delta, src_start, src_end), &caps[3])
        }).to_string();
        cells.push(adjusted);
    }
    cells
}

fn insert_xml_row_ordered(sheet_xml: &str, row_xml: &str) -> Result<String, String> {
    let target: i64 = get_xml_attr(row_xml, "r").and_then(|s| s.parse().ok()).unwrap_or(0);
    let row_re = Regex::new(r#"<row\b[^>/]*\br="(\d+)"[^>/]*>[\s\S]*?</row>"#).unwrap();
    for m in row_re.find_iter(sheet_xml) {
        let n: i64 = get_xml_attr(m.as_str(), "r").and_then(|s| s.parse().ok()).unwrap_or(0);
        if n > target { return Ok(sheet_xml.replacen(m.as_str(), &format!("{row_xml}{}", m.as_str()), 1)); }
    }
    if sheet_xml.contains("</sheetData>") { Ok(sheet_xml.replace("</sheetData>", &format!("{row_xml}</sheetData>"))) }
    else { Err("La hoja no tiene una estructura sheetData valida.".to_string()) }
}

fn paste_range_cells(sheet_xml: &str, dst_row: usize, cells: &[String], col_start: usize, col_end: usize) -> Result<String, String> {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    match get_xml_row(sheet_xml, dst_row) {
        None => {
            let new_row = format!("<row r=\"{dst_row}\">{}</row>", cells.join(""));
            insert_xml_row_ordered(sheet_xml, &new_row)
        }
        Some(target_row_xml) => {
            let mut updated = cell_re.replace_all(&target_row_xml, |caps: &regex::Captures| {
                if let Some(r) = get_xml_attr(&caps[0], "r") {
                    let ci = col_index_from_ref(&r);
                    if ci >= col_start && ci <= col_end { return String::new(); }
                }
                caps[0].to_string()
            }).to_string();
            for cell in cells {
                let ref_attr = get_xml_attr(cell, "r").unwrap_or_default();
                updated = insert_xml_cell_in_row(&updated, cell, col_index_from_ref(&ref_attr));
            }
            Ok(sheet_xml.replacen(&target_row_xml, &updated, 1))
        }
    }
}

fn copy_activity_merges_xml(sheet_xml: &str, source_start: usize, source_end: usize, target_start: usize, type_start_col: usize, type_end_col: usize) -> Result<String, String> {
    let row_delta = (target_start as i64) - (source_start as i64);
    let src_s = (source_start + 1) as i64; let src_e = (source_end + 1) as i64;
    let tgt_s = (target_start + 1) as i64; let tgt_e = src_e + row_delta;

    let decode = |s: &str| -> Option<(usize, i64)> {
        let col: String = s.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let row: i64 = s.chars().skip_while(|c| c.is_ascii_alphabetic()).collect::<String>().parse().ok()?;
        Some((col_index(&col), row))
    };

    let merge_re = Regex::new(r#"<mergeCell\b[^>]*\bref="([^"]+)"[^>]*/>"#).unwrap();
    let mut keep: Vec<String> = Vec::new(); let mut new_refs: Vec<String> = Vec::new();

    for m in merge_re.find_iter(sheet_xml) {
        let ref_val = get_xml_attr(m.as_str(), "ref").unwrap_or_default();
        let parts: Vec<&str> = ref_val.split(':').collect();
        let (sc, sr) = match decode(parts[0]) { Some(v) => v, None => { keep.push(ref_val); continue; } };
        let (ec, er) = match decode(parts.get(1).copied().unwrap_or(parts[0])) { Some(v) => v, None => { keep.push(ref_val); continue; } };
        let tgt_overlap = !(er < tgt_s || sr > tgt_e || ec < type_start_col || sc > type_end_col);
        if !tgt_overlap { keep.push(ref_val.clone()); }
        if sr >= src_s && er <= src_e && sc >= type_start_col && ec <= type_end_col {
            let ns = format!("{}{}", col_name(sc), sr + row_delta);
            let ne = format!("{}{}", col_name(ec), er + row_delta);
            new_refs.push(if ns == ne { ns } else { format!("{ns}:{ne}") });
        }
    }

    let all_refs: Vec<String> = { let mut seen = std::collections::HashSet::new(); keep.into_iter().chain(new_refs).filter(|r| seen.insert(r.clone())).collect() };
    let merge_xml = if all_refs.is_empty() { String::new() } else {
        format!("<mergeCells count=\"{}\">{}</mergeCells>", all_refs.len(), all_refs.iter().map(|r| format!("<mergeCell ref=\"{}\"/>", escape_xml(r))).collect::<String>())
    };
    let full_re = Regex::new(r#"<mergeCells\b[\s\S]*?</mergeCells>"#).unwrap();
    if full_re.is_match(sheet_xml) { Ok(full_re.replace(sheet_xml, merge_xml.as_str()).to_string()) }
    else if merge_xml.is_empty() { Ok(sheet_xml.to_string()) }
    else { Ok(sheet_xml.replace("</sheetData>", &format!("</sheetData>{merge_xml}"))) }
}

// Expande las "shared formulas" de Excel dentro del rango dado.
// Celdas secundarias tienen <f t="shared" si=N/> (sin fórmula). Sin expandir, al copiar
// quedan vacías y el Excel se corrompe.
fn expand_shared_formulas_in_range(sheet_xml: &str, src_row_start: i64, src_row_end: i64, col_start: usize, col_end: usize) -> String {
    let cell_re = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();

    // Paso 1: recopilar masters (celdas con <f t="shared" si=N ref=...>FORMULA</f>)
    struct Master { formula: String, master_col: String, master_row: i64 }
    let mut masters: HashMap<String, Master> = HashMap::new();

    for m in cell_re.find_iter(sheet_xml) {
        let cell = m.as_str();
        let ref_attr = match get_xml_attr(cell, "r") { Some(r) => r, None => continue };
        let f_match = Regex::new(r#"<f\b([^>/]*)>([\s\S]*?)</f>"#).unwrap();
        if let Some(fm) = f_match.captures(cell) {
            let f_tag_inner = fm.get(1).map(|x| x.as_str()).unwrap_or("");
            let formula = fm.get(2).map(|x| x.as_str().trim().to_string()).unwrap_or_default();
            let f_tag = format!("<f {f_tag_inner}>");
            let t_attr = get_xml_attr(&f_tag, "t");
            let si_attr = get_xml_attr(&f_tag, "si");
            let ref2 = get_xml_attr(&f_tag, "ref");
            if t_attr.as_deref() == Some("shared") && si_attr.is_some() && ref2.is_some() && !formula.is_empty() {
                let col_part: String = ref_attr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
                let row_num: i64 = ref_attr.chars().skip_while(|c| c.is_ascii_alphabetic()).collect::<String>().parse().unwrap_or(0);
                masters.insert(si_attr.unwrap(), Master { formula, master_col: col_part, master_row: row_num });
            }
        }
    }

    if masters.is_empty() { return sheet_xml.to_string(); }

    // Paso 2: reemplazar celdas secundarias (self-closing <f t="shared" si=N/>) en el rango
    cell_re.replace_all(sheet_xml, |caps: &regex::Captures| {
        let cell = &caps[0];
        let ref_attr = match get_xml_attr(cell, "r") { Some(r) => r, None => return cell.to_string() };
        let col_part: String = ref_attr.chars().take_while(|c| c.is_ascii_alphabetic()).collect();
        let row_num: i64 = ref_attr.chars().skip_while(|c| c.is_ascii_alphabetic()).collect::<String>().parse().unwrap_or(0);
        let col_idx = col_index(&col_part);
        if col_idx < col_start || col_idx > col_end || row_num < src_row_start || row_num > src_row_end {
            return cell.to_string();
        }
        // Buscar self-closing shared formula: <f ... />
        let sc_re = Regex::new(r#"<f\b([^>/]*)/>"#).unwrap();
        if let Some(fm) = sc_re.captures(cell) {
            let f_inner = fm.get(1).map(|x| x.as_str()).unwrap_or("");
            let f_tag = format!("<f {f_inner}>");
            let t_attr = get_xml_attr(&f_tag, "t");
            let si_attr = get_xml_attr(&f_tag, "si");
            if t_attr.as_deref() == Some("shared") {
                if let Some(si) = si_attr {
                    if let Some(master) = masters.get(&si) {
                        let row_delta = row_num - master.master_row;
                        let concrete = Regex::new(r"(\$?[A-Z]{1,3})(\$?)(\d+)").unwrap().replace_all(&master.formula, |c2: &regex::Captures| {
                            let col = &c2[1]; let abs_row = &c2[2]; let rn: i64 = c2[3].parse().unwrap_or(0);
                            if !abs_row.is_empty() { c2[0].to_string() } else { format!("{}{}", col, rn + row_delta) }
                        }).to_string();
                        let expanded_f = format!("<f>{concrete}</f>");
                        return cell.replacen(fm.get(0).unwrap().as_str(), &expanded_f, 1);
                    }
                }
            }
        }
        cell.to_string()
    }).to_string()
}

fn copy_activity_block_xml(sheet_xml: &str, source_start: usize, source_end: usize, target_start: usize,
    type_start_col: usize, type_end_col: usize, number_col: usize, name_value_col: usize,
    included_row_offset: usize, included_col: usize, first_student_row_offset: usize, note_col: usize,
    notes_to_clear: usize, new_number: i64, nombre_actividad: &str, incluida: bool
) -> Result<String, String> {
    let row_delta = (target_start as i64) - (source_start as i64);
    let src_s = (source_start + 1) as i64; let src_e = (source_end + 1) as i64;
    // Expandir shared formulas antes de copiar para evitar corrupción del Excel
    let mut xml = expand_shared_formulas_in_range(sheet_xml, src_s, src_e, type_start_col, type_end_col);

    for row_idx in source_start..=source_end {
        let src_row = row_idx + 1;
        let dst_row = (src_row as i64 + row_delta) as usize;
        let source_row_xml = match get_xml_row(&xml, src_row) { Some(r) => r, None => continue };
        let cloned = extract_and_adjust_cells(&source_row_xml, dst_row as i64, row_delta, src_s, src_e, type_start_col, type_end_col);
        xml = paste_range_cells(&xml, dst_row, &cloned, type_start_col, type_end_col)?;
    }

    xml = copy_activity_merges_xml(&xml, source_start, source_end, target_start, type_start_col, type_end_col)?;
    xml = set_xml_cell(&xml, target_start + 1, number_col, Some(&json!(new_number)), "number")?;
    let nombre_val = json!(nombre_actividad);
    xml = set_xml_cell(&xml, target_start + 1, name_value_col, if nombre_actividad.is_empty() { None } else { Some(&nombre_val) }, "text")?;
    let x_val = json!("X");
    xml = set_xml_cell(&xml, target_start + included_row_offset, included_col, if incluida { Some(&x_val) } else { None }, "text")?;

    let student_start_row = target_start + first_student_row_offset + 1;
    let student_end_row = student_start_row + notes_to_clear;
    let row_re = Regex::new(r#"<row\b[^>/]*\br="(\d+)"[^>/]*>[\s\S]*?</row>"#).unwrap();
    let cell_re2 = Regex::new(r#"<c\b[^>/]*(?:/>|>[\s\S]*?</c>)"#).unwrap();
    xml = row_re.replace_all(&xml, |caps: &regex::Captures| {
        let row_xml = &caps[0];
        let rn: usize = get_xml_attr(row_xml, "r").and_then(|s| s.parse().ok()).unwrap_or(0);
        if rn < student_start_row || rn >= student_end_row { return row_xml.to_string(); }
        cell_re2.replace_all(row_xml, |caps2: &regex::Captures| {
            let cell = &caps2[0];
            if let Some(r) = get_xml_attr(cell, "r") {
                if col_index_from_ref(&r) == note_col {
                    let style = get_xml_attr(cell, "s").map(|s| format!(" s=\"{}\"", escape_xml(&s))).unwrap_or_default();
                    return format!("<c r=\"{r}\"{style}><v>0</v></c>");
                }
            }
            cell.to_string()
        }).to_string()
    }).to_string();

    Ok(xml)
}

#[tauri::command]
async fn excel_add_actividad(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_add_actividad_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_add_actividad_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().unwrap_or("U1").to_string();
    let tipo = payload["tipo"].as_str().unwrap_or("practicas").to_string();
    let nombre_actividad = payload["nombreActividad"].as_str().unwrap_or("").to_string();
    let incluida = payload["incluida"].as_bool().unwrap_or(true);

    let rows = read_sheet_rows(&path, &unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(&tipo);
    let fixed = activity_fixed_cols(at.key);
    let mut blocks = find_activity_blocks(&rows, at.key);
    blocks.sort_by_key(|b| b.numero);
    let prev_block = blocks.last().ok_or_else(|| format!("No se encontro ninguna actividad previa de {} en {unidad}.", at.label))?.clone();
    let new_number = prev_block.numero + 1;
    let source_start = prev_block.title_row;
    let source_end = source_start + ACTIVITY_BLOCK_ROWS - 1;
    let target_start = source_start + ACTIVITY_BLOCK_STRIDE;
    let notes_to_clear = extract_activity_notes(&rows, &prev_block, None).len();
    let included_row_offset = prev_block.included_row - prev_block.title_row;
    let first_student_row_offset = prev_block.first_student_row - prev_block.title_row;
    let path_clone = path.clone(); let unidad_clone = unidad.clone(); let tipo_clone = tipo.clone();

    edit_workbook_sheets_xml(&path, vec![(&unidad, Box::new(move |xml: &str| {
        copy_activity_block_xml(xml, source_start, source_end, target_start, fixed.start, fixed.end,
            prev_block.number_col, prev_block.name_value_col, included_row_offset, prev_block.included_col,
            first_student_row_offset, prev_block.note_col, notes_to_clear, new_number, &nombre_actividad, incluida)
    }))])?;

    load_notas_actividad(&path_clone, &unidad_clone, &tipo_clone, new_number, None)
}

// ---------------------------------------------------------------------------
// save_ce_notas (batch: escribe la nota de un CE en todas las filas de alumnos)
// ---------------------------------------------------------------------------

#[tauri::command]
async fn excel_save_ce_notas(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_ce_notas_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_ce_notas_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let tipo = payload["tipo"].as_str().ok_or("Falta tipo")?.to_string();
    let actividad = payload["actividad"].as_i64().unwrap_or(1);
    let ce_notas = payload["ceNotas"].as_object().cloned().unwrap_or_default();

    if ce_notas.is_empty() {
        return load_notas_actividad(&path, &unidad, &tipo, actividad, None);
    }

    let rows = read_sheet_rows(&path, &unidad).map_err(|_| format!("El archivo no tiene la hoja \"{unidad}\"."))?;
    let at = get_activity_type(&tipo);
    let blocks = find_activity_blocks(&rows, at.key);
    let block = blocks.iter().find(|b| b.numero == actividad)
        .ok_or_else(|| format!("No se encontro la actividad {actividad} de {} en {unidad}.", at.label))?.clone();

    let path_clone = path.clone(); let unidad_clone = unidad.clone(); let tipo_clone = tipo.clone();

    edit_workbook_sheets_xml(&path, vec![(&unidad, Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        let mut ri = block.first_student_row;
        loop {
            if ri >= rows.len() { break; }
            let nombre = cell_str(&rows, ri, block.name_col);
            if nombre.is_empty() || nombre == "0" { break; }
            for (code, ci) in &block.ce_cols {
                if let Some(val) = ce_notas.get(code) {
                    if let Some(n) = normalize_grade(val) {
                        s = set_xml_cell(&s, ri, *ci, Some(&json!(n)), "number")?;
                    }
                }
            }
            ri += 1;
        }
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    load_notas_actividad(&path_clone, &unidad_clone, &tipo_clone, actividad, None)
}

// ---------------------------------------------------------------------------
// Diario de actividades
// ---------------------------------------------------------------------------

const DIARIO_SHEET: &str = "Diario";

// Añade la hoja "Diario" al XLSX si no existe.
fn ensure_diario_sheet(path: &str) -> Result<(), String> {
    let names = sheet_names(path)?;
    if names.iter().any(|n| n == DIARIO_SHEET) {
        return Ok(());
    }

    // Plantilla XML mínima para una hoja nueva
    let blank_sheet_xml = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData/>
</worksheet>"#;

    let input = std::fs::read(path).map_err(|e| e.to_string())?;
    let backup_bytes = input.clone();
    let cursor = std::io::Cursor::new(input);
    let mut zip = zip::ZipArchive::new(cursor).map_err(|e| e.to_string())?;

    let mut files: HashMap<String, Vec<u8>> = HashMap::new();
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).map_err(|e| e.to_string())?;
        let name = f.name().to_string();
        let mut buf = Vec::new();
        f.read_to_end(&mut buf).map_err(|e| e.to_string())?;
        files.insert(name, buf);
    }

    // Determinar siguiente sheetId y rId
    let workbook_xml = String::from_utf8_lossy(files.get("xl/workbook.xml").ok_or("No xl/workbook.xml")?).to_string();
    let sheet_re = Regex::new(r#"sheetId="(\d+)""#).unwrap();
    let max_id: u32 = sheet_re.captures_iter(&workbook_xml)
        .filter_map(|c| c[1].parse::<u32>().ok())
        .max().unwrap_or(0);
    let new_sheet_id = max_id + 1;

    let rels_xml = String::from_utf8_lossy(files.get("xl/_rels/workbook.xml.rels").ok_or("No workbook.xml.rels")?).to_string();
    let rid_re = Regex::new(r#"Id="rId(\d+)""#).unwrap();
    let max_rid: u32 = rid_re.captures_iter(&rels_xml)
        .filter_map(|c| c[1].parse::<u32>().ok())
        .max().unwrap_or(0);
    let new_rid = max_rid + 1;
    let new_rid_str = format!("rId{}", new_rid);
    let new_sheet_path = format!("xl/worksheets/sheet{}.xml", new_sheet_id);

    // Añadir entrada en workbook.xml (antes de </sheets>)
    let sheet_tag = format!(r#"<sheet name="{}" sheetId="{}" r:id="{}"/>"#, DIARIO_SHEET, new_sheet_id, new_rid_str);
    let new_workbook = workbook_xml.replace("</sheets>", &format!("{}</sheets>", sheet_tag));
    files.insert("xl/workbook.xml".to_string(), new_workbook.into_bytes());

    // Añadir relación en workbook.xml.rels
    let rel_tag = format!(r#"<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>"#, new_rid_str, new_sheet_id);
    let new_rels = rels_xml.replace("</Relationships>", &format!("{}</Relationships>", rel_tag));
    files.insert("xl/_rels/workbook.xml.rels".to_string(), new_rels.into_bytes());

    // Añadir la hoja en blanco
    files.insert(new_sheet_path.clone(), blank_sheet_xml.as_bytes().to_vec());

    // Añadir override en [Content_Types].xml
    let ct_xml = String::from_utf8_lossy(files.get("[Content_Types].xml").ok_or("No [Content_Types].xml")?).to_string();
    let override_tag = format!(r#"<Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>"#, new_sheet_path);
    let new_ct = ct_xml.replace("</Types>", &format!("{}</Types>", override_tag));
    files.insert("[Content_Types].xml".to_string(), new_ct.into_bytes());

    // Reescribir ZIP
    let names_ordered: Vec<String> = {
        let input2 = std::fs::read(path).map_err(|e| e.to_string())?;
        let mut z2 = zip::ZipArchive::new(std::io::Cursor::new(input2)).map_err(|e| e.to_string())?;
        (0..z2.len()).filter_map(|i| z2.by_index_raw(i).ok().map(|f| f.name().to_string())).collect()
    };
    let mut final_order = names_ordered;
    for k in files.keys() {
        if !final_order.contains(k) { final_order.push(k.clone()); }
    }

    let out_buf = Vec::new();
    let mut zw = zip::ZipWriter::new(std::io::Cursor::new(out_buf));
    let options = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated)
        .compression_level(Some(6));
    for name in &final_order {
        if let Some(data) = files.get(name) {
            zw.start_file(name, options).map_err(|e| e.to_string())?;
            zw.write_all(data).map_err(|e| e.to_string())?;
        }
    }
    let out_cursor = zw.finish().map_err(|e| e.to_string())?;
    let _ = std::fs::write(format!("{path}.bak"), &backup_bytes);
    std::fs::write(path, out_cursor.into_inner()).map_err(|e| e.to_string())?;
    Ok(())
}

// Lee B2 de DATOS (nombre del módulo) y todas las entradas de la hoja Diario.
// Hoja Diario: col A=fecha(YYYY-MM-DD), B=hora, C=modulo, D=texto
#[tauri::command]
fn excel_get_diario() -> Result<Value, String> {
    let path = require_selected_path()?;

    // Nombre de la asignatura desde DATOS A2
    let datos_rows = read_sheet_rows(&path, "DATOS")?;
    let modulo = cell_str(&datos_rows, 1, 0); // A2

    // Intentar leer hoja Diario (puede no existir aún)
    // Columnas: A=fecha, B=hora, C=modulo, D=texto, E=tarea, F=estado,
    //           G=incidencia, H=incidenciaTipo, I=incidenciaAlumno, J=incidenciaDesc
    let entradas: Vec<Value> = match read_sheet_rows(&path, DIARIO_SHEET) {
        Ok(rows) => rows.iter().enumerate()
            .skip(1) // skip cabecera
            .filter(|(_, row)| {
                row.get(0).map(|v| !matches!(v, Value::Null)).unwrap_or(false)
                    && !cell_val_str(row.get(0).unwrap_or(&Value::Null)).is_empty()
            })
            .map(|(ri, row)| {
                let cv = |i: usize| cell_val_str(row.get(i).unwrap_or(&Value::Null));
                json!({
                    "fila": ri,
                    "fecha": cv(0), "hora": cv(1), "modulo": cv(2), "texto": cv(3),
                    "tarea": cv(4), "estado": cv(5),
                    "incidencia": cv(6), "incidenciaTipo": cv(7),
                    "incidenciaAlumno": cv(8), "incidenciaDesc": cv(9),
                })
            })
            .collect(),
        Err(_) => vec![],
    };

    Ok(json!({ "modulo": modulo, "entradas": entradas }))
}

// Añade una entrada al diario. Crea la hoja Diario si no existe (con cabecera).
#[tauri::command]
fn excel_save_diario_entrada(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let gs = |k: &str| payload[k].as_str().unwrap_or("").to_string();
    let fecha = gs("fecha"); if fecha.is_empty() { return Err("Falta fecha".into()); }
    let hora = gs("hora"); if hora.is_empty() { return Err("Falta hora".into()); }
    let modulo = gs("modulo"); let texto = gs("texto");
    let tarea = gs("tarea"); let estado = gs("estado");
    let incidencia = gs("incidencia"); let inc_tipo = gs("incidenciaTipo");
    let inc_alumno = gs("incidenciaAlumno"); let inc_desc = gs("incidenciaDesc");

    ensure_diario_sheet(&path)?;

    let existing_rows = match read_sheet_rows(&path, DIARIO_SHEET) {
        Ok(rows) => rows, Err(_) => vec![],
    };
    let need_header = existing_rows.len() < 2
        || cell_str(&existing_rows, 0, 0).to_uppercase() != "FECHA";
    let last_data_row = if existing_rows.len() <= 1 { 1 } else {
        let last = (1..existing_rows.len()).rev()
            .find(|&i| !cell_str(&existing_rows, i, 0).is_empty()).unwrap_or(0);
        last + 1
    };
    let insert_row = last_data_row.max(1);

    let (fv, hv, mv, tv, tav, ev, iv, itv, iav, idv) = (
        json!(fecha), json!(hora), json!(modulo), json!(texto),
        json!(tarea), json!(estado), json!(incidencia),
        json!(inc_tipo), json!(inc_alumno), json!(inc_desc),
    );
    let nhf = need_header;

    edit_workbook_sheets_xml(&path, vec![(DIARIO_SHEET, Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        if nhf {
            for (ci, h) in ["FECHA","HORA","MODULO","ACTIVIDAD","TAREA","ESTADO",
                             "INCIDENCIA","INC_TIPO","INC_ALUMNO","INC_DESC"].iter().enumerate() {
                s = set_xml_cell(&s, 0, ci, Some(&json!(h)), "text")?;
            }
        }
        s = set_xml_cell(&s, insert_row, 0, Some(&fv), "text")?;
        s = set_xml_cell(&s, insert_row, 1, Some(&hv), "text")?;
        s = set_xml_cell(&s, insert_row, 2, Some(&mv), "text")?;
        s = set_xml_cell(&s, insert_row, 3, Some(&tv), "text")?;
        s = set_xml_cell(&s, insert_row, 4, if tav.as_str()==Some("") { None } else { Some(&tav) }, "text")?;
        s = set_xml_cell(&s, insert_row, 5, if ev.as_str()==Some("") { None } else { Some(&ev) }, "text")?;
        s = set_xml_cell(&s, insert_row, 6, if iv.as_str()==Some("") { None } else { Some(&iv) }, "text")?;
        s = set_xml_cell(&s, insert_row, 7, if itv.as_str()==Some("") { None } else { Some(&itv) }, "text")?;
        s = set_xml_cell(&s, insert_row, 8, if iav.as_str()==Some("") { None } else { Some(&iav) }, "text")?;
        s = set_xml_cell(&s, insert_row, 9, if idv.as_str()==Some("") { None } else { Some(&idv) }, "text")?;
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    excel_get_diario()
}

// Borra una entrada del diario (borra las 10 celdas de la fila indicada).
#[tauri::command]
fn excel_delete_diario_entrada(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let fila = payload["fila"].as_u64().ok_or("Falta fila")? as usize;

    edit_workbook_sheets_xml(&path, vec![(DIARIO_SHEET, Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for ci in 0..10 { s = set_xml_cell(&s, fila, ci, None, "text")?; }
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    excel_get_diario()
}

// ---------------------------------------------------------------------------
// Instrumentos de evaluación (DATOS cols N/O, filas 5-9)
// ---------------------------------------------------------------------------

fn load_instrumentos(path: &str) -> Result<Value, String> {
    let rows = read_sheet_rows(path, "DATOS")?;
    let mut instrumentos: Vec<Value> = Vec::new();
    // DATOS filas 5-9 (0-indexed 4-8): col N(13)=código, O(14)=nombre
    for ri in 4..=13usize {
        let codigo = cell_str(&rows, ri, 13);
        let nombre = cell_str(&rows, ri, 14);
        if codigo.is_empty() && nombre.is_empty() { continue; }
        instrumentos.push(json!({ "codigo": codigo, "nombre": nombre }));
    }
    let file_name = Path::new(path).file_name().unwrap_or_default().to_string_lossy().to_string();
    Ok(json!({ "filePath": path, "fileName": file_name, "instrumentos": instrumentos }))
}

fn save_instrumentos_to_file(path: &str, instrumentos: &[Value]) -> Result<(), String> {
    let instrumentos_owned = instrumentos.to_vec();
    edit_workbook_sheets_xml(path, vec![("DATOS", Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for slot in 0..10usize {
            let ri = 4 + slot;
            if let Some(inst) = instrumentos_owned.get(slot) {
                let codigo = inst["codigo"].as_str().unwrap_or("");
                let nombre = inst["nombre"].as_str().unwrap_or("");
                let codigo_val = json!(codigo);
                let nombre_val = json!(nombre);
                s = set_xml_cell(&s, ri, 13, if codigo.is_empty() { None } else { Some(&codigo_val) }, "text")?;
                s = set_xml_cell(&s, ri, 14, if nombre.is_empty() { None } else { Some(&nombre_val) }, "text")?;
            } else {
                s = set_xml_cell(&s, ri, 13, None, "text")?;
                s = set_xml_cell(&s, ri, 14, None, "text")?;
            }
        }
        Ok(s)
    }))])
}

#[tauri::command]
fn excel_get_instrumentos() -> Result<Value, String> {
    if get_selected_path().is_none() { set_selected_path(find_default_excel_path()); }
    match get_selected_path() { Some(p) => load_instrumentos(&p), None => Ok(Value::Null) }
}

#[tauri::command]
fn excel_save_instrumentos(instrumentos: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let arr = instrumentos.as_array().ok_or("Lista de instrumentos no valida.")?.clone();
    save_instrumentos_to_file(&path, &arr)?;
    load_instrumentos(&path)
}

// ---------------------------------------------------------------------------
// open external + main
// ---------------------------------------------------------------------------

#[tauri::command]
fn app_open_external(url: String) -> Result<(), String> {
    webbrowser::open(&url).map_err(|e| format!("No se pudo abrir el enlace: {e}"))
}

static TEMPLATE_XLSX: &[u8] = include_bytes!("../../Plantilla_Notas_ESO.xlsx");

#[tauri::command]
fn excel_download_template() -> Result<bool, String> {
    let path = rfd::FileDialog::new()
        .set_title("Guardar plantilla Excel")
        .set_file_name("Plantilla_Notas_ESO.xlsx")
        .add_filter("Excel", &["xlsx"])
        .save_file();
    match path {
        Some(p) => {
            std::fs::write(&p, TEMPLATE_XLSX).map_err(|e| e.to_string())?;
            Ok(true)
        }
        None => Ok(false),
    }
}

#[tauri::command]
fn save_csv_template(filename: String, content: String) -> Result<bool, String> {
    let path = rfd::FileDialog::new()
        .set_file_name(&filename)
        .add_filter("CSV", &["csv"])
        .save_file();
    match path {
        Some(p) => {
            let bom = "\u{FEFF}";
            std::fs::write(&p, format!("{}{}", bom, content)).map_err(|e| e.to_string())?;
            Ok(true)
        }
        None => Ok(false),
    }
}

#[cfg(test)]
mod formula_cache_tests {
    use super::*;

    // Repro real: 'CCGG PLANTILLA - CE AMPLIADOSv3.xlsx' sheet4.xml (2ª EVA) fila 22,
    // celda E22 tenia <f ...>...</f><v/></c> (IFERROR(...,"") -> cache vacio,
    // self-closing). set_xml_formula_cache_number no reconocia el <v/> existente y
    // añadia un <v> nuevo sin quitar el viejo -> <v/><v>10</v></c>, dos <v> en la
    // misma celda -> Excel rechaza el XML y vacia la hoja entera al reparar.
    #[test]
    fn replaces_self_closing_v_instead_of_duplicating() {
        let xml = r#"<worksheet><sheetData><row r="22"><c r="E22" s="241" t="str"><f ca="1">IFERROR(1,"")</f><v/></c></row></sheetData></worksheet>"#;
        let updated = set_xml_formula_cache_number(xml, 21, 4, Some(10.0)).unwrap();
        assert_eq!(updated.matches("<v").count(), 1, "debe quedar un solo <v> en la celda: {updated}");
        assert!(updated.contains("<v>10</v>"), "debe contener el valor nuevo: {updated}");
        assert!(!updated.contains("<v/>"), "no debe quedar el <v/> viejo: {updated}");
    }

    #[test]
    fn clears_self_closing_v_without_error() {
        let xml = r#"<worksheet><sheetData><row r="22"><c r="E22" s="241" t="str"><f ca="1">IFERROR(1,"")</f><v/></c></row></sheetData></worksheet>"#;
        let updated = set_xml_formula_cache_number(xml, 21, 4, None).unwrap();
        assert_eq!(updated.matches("<v").count(), 0, "debe quedar sin <v>: {updated}");
        assert!(updated.contains("<f ca=\"1\">IFERROR(1,\"\")</f>"), "la formula debe seguir intacta: {updated}");
    }

    #[test]
    fn replaces_existing_paired_v() {
        let xml = r#"<worksheet><sheetData><row r="22"><c r="E22" s="241"><f ca="1">1+1</f><v>2</v></c></row></sheetData></worksheet>"#;
        let updated = set_xml_formula_cache_number(xml, 21, 4, Some(5.0)).unwrap();
        assert_eq!(updated.matches("<v").count(), 1, "debe quedar un solo <v>: {updated}");
        assert!(updated.contains("<v>5</v>"), "{updated}");
    }

    // Formulas compartidas (t="shared"): la celda "esclava" solo referencia el
    // indice (si="N"), el texto de la formula vive en la celda "maestra". Escribir
    // el cache no debe tocar ese tag <f>, solo el <v>.
    #[test]
    fn preserves_shared_formula_slave_tag() {
        let xml = r#"<worksheet><sheetData><row r="22"><c r="B22" s="238"><f t="shared" si="191"/><v>1.9999999999999998</v></c></row></sheetData></worksheet>"#;
        let updated = set_xml_formula_cache_number(xml, 21, 1, Some(3.0)).unwrap();
        assert!(updated.contains(r#"<f t="shared" si="191"/>"#), "el tag de formula compartida debe seguir intacto: {updated}");
        assert_eq!(updated.matches("<v").count(), 1, "{updated}");
        assert!(updated.contains("<v>3</v>"), "{updated}");
    }

    // Combinacion peor caso: formula compartida (self-closing <f>) Y cache vacio
    // (self-closing <v/>) a la vez, el patron real que corrompio sheet4/5 EVA.
    #[test]
    fn preserves_shared_formula_slave_tag_with_empty_cache() {
        let xml = r#"<worksheet><sheetData><row r="22"><c r="E22" s="241" t="str"><f t="shared" si="7"/><v/></c></row></sheetData></worksheet>"#;
        let updated = set_xml_formula_cache_number(xml, 21, 4, Some(10.0)).unwrap();
        assert!(updated.contains(r#"<f t="shared" si="7"/>"#), "{updated}");
        assert_eq!(updated.matches("<v").count(), 1, "{updated}");
        assert!(updated.contains("<v>10</v>"), "{updated}");
    }
}

#[cfg(test)]
mod eval_layout_tests {
    use super::*;

    fn row(cells: &[&str]) -> Vec<Value> {
        cells.iter().map(|c| json!(*c)).collect()
    }

    // Estrategia 1: "NOTA CE" y >=2 codigos CR en la MISMA fila (layout ESO tipico,
    // documentado en CLAUDE.md: fila 17 0-idx16 = cabecera).
    #[test]
    fn detects_layout_same_row_strategy() {
        let rows = vec![
            row(&["Alumno"]),
            row(&["NOTA CE", "CR1.1", "Rec", "CR1.2", "Rec"]),
            row(&[""]),
            row(&["Juan", "7", "", "8", ""]),
        ];
        let result = find_evaluation_layout_indices(&rows);
        assert_eq!(result, Some((1, 1, 3)));
    }

    // Estrategia 2: "NOTA CE" en una fila, los codigos CR aparecen 1-3 filas despues.
    #[test]
    fn detects_layout_offset_row_strategy() {
        let rows = vec![
            row(&["Alumno"]),
            row(&["NOTA CE"]),
            row(&["", "Rec"]),
            row(&["CR2.1", "CR2.2"]),
            row(&["Juan", "7", "8"]),
        ];
        let result = find_evaluation_layout_indices(&rows);
        assert_eq!(result, Some((1, 3, 4)));
    }

    // Estrategia 3 (fallback debil): sin "NOTA CE" en ningun sitio, solo una fila
    // con >=2 codigos CR — usa la fila anterior como cabecera.
    #[test]
    fn detects_layout_fallback_strategy() {
        let rows = vec![
            row(&["Cabecera generica"]),
            row(&["CR3.1", "CR3.2"]),
            row(&["Juan", "7", "8"]),
        ];
        let result = find_evaluation_layout_indices(&rows);
        assert_eq!(result, Some((0, 1, 2)));
    }

    // Layout no reconocible: ni "NOTA CE" ni >=2 codigos CR en ninguna fila.
    #[test]
    fn returns_none_when_no_layout_detected() {
        let rows = vec![
            row(&["Alumno"]),
            row(&["Solo texto", "sin codigos"]),
        ];
        assert_eq!(find_evaluation_layout_indices(&rows), None);
    }

}

#[cfg(test)]
mod final_weighted_tests {
    use super::*;

    // Replica la fórmula real de la plantilla (U1!L7):
    // =IFERROR(SUMPRODUCT((H7:K7<>"")*C$4:F$4*H7:K7)/SUMPRODUCT((H7:K7<>"")*C$4:F$4),"")
    const PESOS_REALES: [f64; 4] = [0.2, 0.4, 0.2, 0.2];

    #[test]
    fn media_ponderada_con_los_4_instrumentos() {
        let values = [Some(8.0), Some(7.0), Some(9.0), Some(6.0)];
        let result = compute_final_weighted(values, PESOS_REALES);
        // (0.2*8 + 0.4*7 + 0.2*9 + 0.2*6) / (0.2+0.4+0.2+0.2) = 7.4 / 1.0
        assert_eq!(result, Some(7.4));
    }

    #[test]
    fn ignora_instrumentos_vacios_en_numerador_y_denominador() {
        let values = [Some(8.0), None, Some(9.0), Some(6.0)];
        let result = compute_final_weighted(values, PESOS_REALES);
        // (0.2*8 + 0.2*9 + 0.2*6) / (0.2+0.2+0.2) = 4.6 / 0.6 = 7.666...
        assert!((result.unwrap() - 7.6666666666666).abs() < 1e-9);
    }

    #[test]
    fn valor_cero_explicito_cuenta_distinto_de_vacio() {
        // Un 0 tecleado SI cuenta (no es "" en Excel); solo None (celda vacia) se ignora.
        let values = [Some(0.0), None, None, None];
        let result = compute_final_weighted(values, PESOS_REALES);
        assert_eq!(result, Some(0.0));
    }

    #[test]
    fn todos_vacios_devuelve_none() {
        let result = compute_final_weighted([None, None, None, None], PESOS_REALES);
        assert_eq!(result, None);
    }

    #[test]
    fn peso_cero_en_slot_presente_no_distorsiona_el_resto() {
        let values = [Some(10.0), Some(4.0), None, None];
        let weights = [0.0, 0.5, 0.3, 0.2];
        let result = compute_final_weighted(values, weights);
        // (0*10 + 0.5*4) / (0 + 0.5) = 2.0 / 0.5 = 4.0
        assert_eq!(result, Some(4.0));
    }
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            excel_select_file, excel_get_selected_file, excel_set_selected_file,
            excel_verify_file_exists, excel_get_alumnos, excel_save_alumnos, excel_get_unidades,
            excel_save_unidades, excel_get_rraa_criterios, excel_save_rraa_criterios,
            excel_get_notas_actividad, excel_get_notas_actividades_tipo,
            excel_save_notas_actividad, excel_save_ce_notas, excel_add_actividad,
            excel_get_notas_evaluacion, excel_get_notas_evaluacion_alumno,
            excel_get_notas_unidad, excel_save_notas_unidad, excel_resync_unidad_eval, excel_get_alumnos_informes, app_open_external,
            excel_get_diario, excel_save_diario_entrada, excel_delete_diario_entrada,
            excel_get_instrumentos, excel_save_instrumentos, save_csv_template, excel_download_template
        ])
        .run(tauri::generate_context!())
        .expect("error al ejecutar la aplicacion Tauri");
}
