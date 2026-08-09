(function () {
  if (window.electronExcel) {
    return;
  }

  const tauriCore = window.__TAURI__ && window.__TAURI__.core;

  if (!tauriCore || typeof tauriCore.invoke !== "function") {
    return;
  }

  const invoke = tauriCore.invoke;

  window.electronExcel = {
    selectFile: () => invoke("excel_select_file"),
    getSelectedFile: () => invoke("excel_get_selected_file"),
    getAlumnos: () => invoke("excel_get_alumnos"),
    saveAlumnos: (alumnos) => invoke("excel_save_alumnos", { alumnos }),
    getUnidades: () => invoke("excel_get_unidades"),
    saveUnidades: (unidades) => invoke("excel_save_unidades", { unidades }),
    getRraaCriterios: () => invoke("excel_get_rraa_criterios"),
    saveRraaCriterios: (payloadOrRraa, criterios, ponderacionesUnidad = []) => {
      const payload = Array.isArray(payloadOrRraa)
        ? { rraa: payloadOrRraa, criterios, ponderacionesUnidad }
        : payloadOrRraa;
      return invoke("excel_save_rraa_criterios", { payload });
    },
    getNotasActividad: (payload) => invoke("excel_get_notas_actividad", { payload }),
    getNotasActividadesTipo: (payload) => invoke("excel_get_notas_actividades_tipo", { payload }),
    saveNotasActividad: (payload) => invoke("excel_save_notas_actividad", { payload }),
    saveCeNotas: (payload) => invoke("excel_save_ce_notas", { payload }),
    addActividad: (payload) => invoke("excel_add_actividad", { payload }),
    getNotasUnidad: (payload) => invoke("excel_get_notas_unidad", { payload }),
    saveNotasUnidad: (payload) => invoke("excel_save_notas_unidad", { payload }),
    saveUnidadInstrumentos: (payload) => invoke("excel_save_unidad_instrumentos", { payload }),
    resyncUnidadEval: (payload) => invoke("excel_resync_unidad_eval", { payload }),
    getNotasEvaluacion: (payload) => invoke("excel_get_notas_evaluacion", { payload }),
    getNotasEvaluacionAlumno: (payload) => invoke("excel_get_notas_evaluacion_alumno", { payload }),
    getAlumnosInformes: () => invoke("excel_get_alumnos_informes"),
    setSelectedFile: (filePath) => invoke("excel_set_selected_file", { filePath }),
    verifyFileExists: (filePath) => invoke("excel_verify_file_exists", { filePath }),
    openExternal: (url) => invoke("app_open_external", { url }),
    getDiarioData: () => invoke("excel_get_diario"),
    saveDiarioEntrada: (payload) => invoke("excel_save_diario_entrada", { payload }),
    deleteDiarioEntrada: (payload) => invoke("excel_delete_diario_entrada", { payload }),
    getInstrumentos: () => invoke("excel_get_instrumentos"),
    saveInstrumentos: (instrumentos) => invoke("excel_save_instrumentos", { instrumentos })
  };
})();
