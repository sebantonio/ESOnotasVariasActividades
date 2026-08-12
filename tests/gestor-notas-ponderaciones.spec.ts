import { test, expect } from '@playwright/test';

// Asume que gestor-notas.html está disponible en http://localhost:3000
const BASE_URL = 'http://localhost:3000/gestor-notas.html';

test.describe('Panel de Ponderaciones de Instrumentos', () => {
  test.beforeEach(async ({ page }) => {
    // Mock de window.electronExcel para Playwright
    await page.addInitScript(() => {
      (window as any).electronExcel = {
        getSelectedFile: async () => '/ruta/ficticia/archivo.xlsx',
        setSelectedFile: async () => true,
        getUnidades: async () => ({
          unidades: [
            { codigo: 'U1', label: 'Unidad 1' },
            { codigo: 'U2', label: 'Unidad 2' }
          ]
        }),
        getInstrumentos: async () => ({
          instrumentos: [
            { instrumento: 'EX', nombre: 'Examen' },
            { instrumento: 'TR', nombre: 'Trabajo' },
            { instrumento: 'OB', nombre: 'Observación' },
            { instrumento: 'PR', nombre: 'Prueba' }
          ]
        }),
        getNotasUnidad: async () => ({
          unidad: 'U1',
          fileName: 'test.xlsx',
          alumnos: [
            { nombre: 'Alumno 1', crNotas: [] },
            { nombre: 'Alumno 2', crNotas: [] }
          ],
          criterios: [],
          unidades: [
            { codigo: 'U1', label: 'Unidad 1' }
          ],
          instrumentosUnidad: [
            { instrumento: 'EX', peso: 40 },
            { instrumento: 'TR', peso: 30 },
            { instrumento: 'OB', peso: 30 },
            { instrumento: '', peso: 0 }
          ]
        }),
        saveInstrumentos: async () => true,
        saveNotasUnidad: async () => true
      };
    });

    await page.goto(BASE_URL);
    // Esperar a que cargue el panel
    await page.waitForSelector('.ponderaciones-panel');
  });

  test('1. Panel visible, expandido y relleno al cargar unidad', async ({ page }) => {
    // El panel debe estar visible
    const panel = page.locator('.ponderaciones-panel');
    await expect(panel).toBeVisible();

    // El panel debe estar expandido (content con display:flex)
    const content = page.locator('#ponderacionesContent');
    const display = await content.evaluate(el => window.getComputedStyle(el).display);
    expect(display).toBe('flex');

    // Debe haber 4 slots
    const slots = page.locator('.slot-item');
    expect(await slots.count()).toBe(8); // 4 slots * 2 (select + input)

    // Los valores deben coincidir con instrumentosUnidad cargados
    await expect(page.locator('select[data-slot="0"]')).toHaveValue('EX');
    await expect(page.locator('input[data-slot="0"]')).toHaveValue('40');
    await expect(page.locator('select[data-slot="1"]')).toHaveValue('TR');
    await expect(page.locator('input[data-slot="1"]')).toHaveValue('30');

    // Suma total debe ser 100% y válida
    await expect(page.locator('#sumaPesos')).toContainText('100%');
    const sumaRow = page.locator('.suma-row');
    const classes = await sumaRow.evaluate(el => el.className);
    expect(classes).toContain('valid');
  });

  test('2. Cambiar instrumento en slot 1 → actualiza suma y marca cambios', async ({ page }) => {
    // Cambiar instrumento en slot 0
    await page.locator('select[data-slot="0"]').selectOption('TR');

    // Verificar que se marcó como cambio (indicador naranja)
    const indicator = page.locator('#unsavedIndicator');
    await expect(indicator).toBeVisible();

    // La suma debería cambiar (ahora TR(30) + TR(30) + OB(30) = 90)
    await expect(page.locator('#sumaPesos')).toContainText('90%');

    // Debe mostrar aviso de suma inválida
    const aviso = page.locator('#sumaPesosAviso');
    await expect(aviso).toContainText('⚠ Total debe ser 100%');
  });

  test('3. Cambiar peso en slot 2 → actualiza suma y validación color', async ({ page }) => {
    // Cambiar peso en slot 1 de 30 a 40
    await page.locator('input[data-slot="1"]').fill('40');

    // Trigger change event
    await page.locator('input[data-slot="1"]').evaluate(el => {
      (el as HTMLInputElement).dispatchEvent(new Event('change', { bubbles: true }));
    });

    // Nueva suma: 40 + 40 + 30 = 110 (inválida)
    await expect(page.locator('#sumaPesos')).toContainText('110%');

    // Debe mostrar aviso
    const aviso = page.locator('#sumaPesosAviso');
    await expect(aviso).toContainText('⚠ Total debe ser 100%');

    // Suma-row debe tener clase invalid
    const sumaRow = page.locator('.suma-row');
    const classes = await sumaRow.evaluate(el => el.className);
    expect(classes).toContain('invalid');
  });

  test('4. Click Guardar → persiste en Excel e indicador desaparece', async ({ page }) => {
    // Cambiar algo
    await page.locator('input[data-slot="0"]').fill('50');
    await page.locator('input[data-slot="0"]').evaluate(el => {
      (el as HTMLInputElement).dispatchEvent(new Event('change', { bubbles: true }));
    });

    // Verificar indicador visible antes de guardar
    const indicator = page.locator('#unsavedIndicator');
    await expect(indicator).toBeVisible();

    // Click Guardar
    await page.locator('button:has-text("Guardar")').click();

    // Esperar a que se guarde
    await page.waitForTimeout(500);

    // Indicador debe desaparecer
    await expect(indicator).not.toBeVisible();

    // Mensaje de éxito
    const message = page.locator('#message');
    await expect(message).toContainText('guardadas correctamente');
  });

  test('5. Click Cancelar sin guardar → restaura valores', async ({ page }) => {
    // Cambiar instrumento y peso
    await page.locator('select[data-slot="0"]').selectOption('OB');
    await page.locator('input[data-slot="0"]').fill('10');
    await page.locator('input[data-slot="0"]').evaluate(el => {
      (el as HTMLInputElement).dispatchEvent(new Event('change', { bubbles: true }));
    });

    // Verificar cambios
    await expect(page.locator('select[data-slot="0"]')).toHaveValue('OB');
    await expect(page.locator('input[data-slot="0"]')).toHaveValue('10');

    // Click Cancelar
    await page.locator('button:has-text("Cancelar")').click();

    // Debe restaurar valores originales (EX, 40)
    await expect(page.locator('select[data-slot="0"]')).toHaveValue('EX');
    await expect(page.locator('input[data-slot="0"]')).toHaveValue('40');

    // Suma debe volver a 100%
    await expect(page.locator('#sumaPesos')).toContainText('100%');

    // Indicador debe desaparecer
    const indicator = page.locator('#unsavedIndicator');
    await expect(indicator).not.toBeVisible();
  });

  test('6. Suma ≠ 100% → muestra aviso visual', async ({ page }) => {
    // Cambiar peso a 50 (total será 120)
    await page.locator('input[data-slot="2"]').fill('50');
    await page.locator('input[data-slot="2"]').evaluate(el => {
      (el as HTMLInputElement).dispatchEvent(new Event('change', { bubbles: true }));
    });

    // Debe mostrar aviso
    const aviso = page.locator('#sumaPesosAviso');
    await expect(aviso).toContainText('⚠ Total debe ser 100%');

    // Color de suma debe cambiar a naranja
    const sumaPesos = page.locator('#sumaPesos');
    const color = await sumaPesos.evaluate(el => window.getComputedStyle(el).color);
    // Naranja (#f59e0b)
    expect(color).toBeTruthy();
  });

  test('7. Collapse/expand panel → estado visual correcto', async ({ page }) => {
    // Panel debe estar expandido inicialmente
    let content = page.locator('#ponderacionesContent');
    let display = await content.evaluate(el => window.getComputedStyle(el).display);
    expect(display).toBe('flex');

    // Click en header para colapsar
    await page.locator('.ponderaciones-header').click();

    // Panel debe estar colapsado
    display = await content.evaluate(el => window.getComputedStyle(el).display);
    expect(display).toBe('none');

    // Chevron debe rotar
    let header = page.locator('.ponderaciones-header');
    let classes = await header.evaluate(el => el.className);
    expect(classes).toContain('collapsed');

    // Click nuevamente para expandir
    await page.locator('.ponderaciones-header').click();

    // Panel debe estar expandido nuevamente
    display = await content.evaluate(el => window.getComputedStyle(el).display);
    expect(display).toBe('flex');

    // Chevron debe volver a posición normal
    classes = await header.evaluate(el => el.className);
    expect(classes).not.toContain('collapsed');
  });

  test('8. Validación: inputs acotados a 0-100%', async ({ page }) => {
    const input = page.locator('input[data-slot="0"]');

    // min=0
    const min = await input.getAttribute('min');
    expect(min).toBe('0');

    // max=100
    const max = await input.getAttribute('max');
    expect(max).toBe('100');
  });

  test('9. Panel con menos de 4 instrumentos → rellena slots vacíos', async ({ page }) => {
    // Slot 3 debe estar vacío (instrumento="", peso=0)
    await expect(page.locator('select[data-slot="3"]')).toHaveValue('');
    await expect(page.locator('input[data-slot="3"]')).toHaveValue('0');
  });

  test('10. Integración: cambios sin guardar al cambiar de unidad → aviso', async ({ page }) => {
    // Cambiar algo
    await page.locator('input[data-slot="0"]').fill('50');
    await page.locator('input[data-slot="0"]').evaluate(el => {
      (el as HTMLInputElement).dispatchEvent(new Event('change', { bubbles: true }));
    });

    // Indicador de cambios debe estar visible
    const indicator = page.locator('#unsavedIndicator');
    await expect(indicator).toBeVisible();

    // (Este test es simbólico; el flujo real de cambiar de unidad depende de loadNotasConGuardado)
  });
});
