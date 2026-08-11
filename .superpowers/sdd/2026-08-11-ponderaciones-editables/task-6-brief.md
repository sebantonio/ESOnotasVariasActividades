# Task 6: Testing Manual

**Scope:** End-to-end testing of the complete weighting interface, from UI through backend to calculation

**Test Plan:**

1. **Setup:** Open the app with a test Excel file that has unit data
2. **UI Toggle Test:**
   - [ ] Open gestor-unidades.html
   - [ ] Verify "Configurar Ponderaciones" button is visible in toolbar
   - [ ] Click button — ponderaciones section should show
   - [ ] Click button again — section should hide

3. **Dropdown Test:**
   - [ ] Click "Configurar Ponderaciones" to show section
   - [ ] Verify dropdown populates with units from Excel
   - [ ] Select a unit — table should show with current weight values

4. **Validation Test:**
   - [ ] Edit weights to: i1=0.5, i2=0.3, i3=0.1, i4=0.1
   - [ ] Verify: display shows 50%, 30%, 10%, 10%
   - [ ] Verify: "Suma = 100% (válido)" message shows, button enabled
   - [ ] Change i1 to 0.4 (suma = 90%)
   - [ ] Verify: message shows "Suma = 90% (debe ser 100%)", button disabled
   - [ ] Correct i1 back to 0.5 — button re-enables

5. **Save Test:**
   - [ ] With valid weights (suma = 100%), click "Guardar Ponderaciones"
   - [ ] Verify: success message appears
   - [ ] Verify: section closes
   - [ ] Open gestor-unidades again, select same unit — weights should match what was saved

6. **Integration Test:**
   - [ ] Save weights for a unit: i1=0.5, i2=0.3, i3=0.1, i4=0.1
   - [ ] Open gestor-notas.html with that unit
   - [ ] Enter sample instrument notes for a student: i1=10, i2=8, i3=6, i4=4
   - [ ] Verify FINAL calculation: (0.5*10 + 0.3*8 + 0.1*6 + 0.1*4) / 1.0 = 5.0 + 2.4 + 0.6 + 0.4 = 8.4
   - [ ] If FINAL displays 8.4 (or close), integration is working

**Notes:**
- Test data should be in an actual Excel file that follows the new layout (units U1-U3 with rows 5 for weights)
- Use browser devtools console to inspect currentState.unidadPesos if needed
- Verify showMessage() displays correctly for both success/error cases

**Success Criteria:**
- All 6 test groups pass
- No console errors
- Weights persist after save/reload
- FINAL calculation uses new weights
