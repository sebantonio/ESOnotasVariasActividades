# Task 3: Backend - Comandos IPC para cargar/guardar pesos

**Files:**
- Modify: `src-tauri/src/main.rs` - agregar comandos

**Interfaces:**
- Consumes: unidad name (string), optional pesos array [f64; 4]
- Produces: JSON response with pesos array or success confirmation
- These commands are called by Task 2 (dropdown change) and Task 4 (save button)

**Steps:**

### 3.1: Comando excel_get_unidad_pesos

```rust
#[tauri::command]
async fn excel_get_unidad_pesos(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_get_unidad_pesos_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_get_unidad_pesos_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    
    let rows = read_sheet_rows(&path, &unidad)?;
    
    // Fila 5 (0-idx 4): C5:F5 = columnas 2:5
    let pesos = [
        cell_f64(&rows, 4, 2).unwrap_or(0.25),  // C5 = i1
        cell_f64(&rows, 4, 3).unwrap_or(0.25),  // D5 = i2
        cell_f64(&rows, 4, 4).unwrap_or(0.25),  // E5 = i3
        cell_f64(&rows, 4, 5).unwrap_or(0.25),  // F5 = i4
    ];
    
    Ok(json!({ "pesos": pesos }))
}
```

### 3.2: Comando excel_save_unidad_pesos

```rust
#[tauri::command]
async fn excel_save_unidad_pesos(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_unidad_pesos_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_unidad_pesos_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let pesos = payload["pesos"].as_array().ok_or("Falta pesos")?;
    
    if pesos.len() != 4 {
        return Err("Se requieren exactamente 4 pesos".to_string());
    }
    
    let pesos_f64: Vec<f64> = pesos.iter()
        .filter_map(|p| p.as_f64())
        .collect();
    
    if pesos_f64.len() != 4 {
        return Err("Todos los pesos deben ser números".to_string());
    }
    
    let suma: f64 = pesos_f64.iter().sum();
    if (suma - 1.0).abs() > 0.001 {
        return Err(format!("Suma debe ser 1.0 (100%), es {:.2}", suma));
    }
    
    let pesos_owned = pesos_f64.clone();
    let edit_fn: Box<dyn Fn(&str) -> Result<String, String>> = Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        // Fila 5 (0-idx 4), columnas C:F (0-idx 2:5)
        for (i, &peso) in pesos_owned.iter().enumerate() {
            s = set_xml_cell(&s, 4, 2 + i, Some(&json!(peso)), "number")?;
        }
        Ok(s)
    });
    
    edit_workbook_sheets_xml(&path, vec![(unidad.as_str(), edit_fn)])?;
    
    Ok(json!({ "pesos": pesos_f64 }))
}
```

**Key Notes:**
- Row 5 is 0-indexed as row 4 (see CLAUDE.md structure for units sheets)
- Columns C:F are 0-indexed as 2:5
- Tolerance for sum validation: 0.001 (allows floating point errors)
- Both functions are async with spawn_blocking pattern (existing pattern in main.rs)
- Use existing helpers: require_selected_path(), read_sheet_rows(), cell_f64(), set_xml_cell(), edit_workbook_sheets_xml()

**Global Constraints:**
- Pesos deben sumar 100% (tolerancia 0.001)
- Valores 0-1 (decimales: 0.2 = 20%)
- Fila 5 de cada hoja de unidad (C5:F5)
