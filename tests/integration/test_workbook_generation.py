import os
import openpyxl
from excel_saas.application.generate_workbook import generate
from excel_saas.core.models.generation_request import GenerationRequest

def test_full_pipeline_light_and_dark(tmp_path):
    output_dir = str(tmp_path)
    
    req_light = GenerationRequest(template_id="finance_personal", year=2026, theme="light")
    path_light = generate(req_light, output_dir)
    
    assert os.path.exists(path_light)
    
    # Verify via openpyxl
    wb = openpyxl.load_workbook(path_light, data_only=False)
    sheet_names = wb.sheetnames
    
    assert "Comece Aqui" in sheet_names
    assert "Dashboard" in sheet_names
    assert "Lançamentos" in sheet_names
    assert "Configurações" in sheet_names
    
    ws = wb["Lançamentos"]
    
    # Check table existence (openpyxl supports this)
    assert "tblLancamentos" in ws.tables
    
    table = ws.tables["tblLancamentos"]
    assert table.ref == "B4:L5"
    
    # Check that protection was enabled
    assert ws.protection.sheet is True # Lançamentos is now protected
    assert wb["Comece Aqui"].protection.sheet is True
    
    req_dark = GenerationRequest(template_id="finance_personal", year=2026, theme="dark")
    path_dark = generate(req_dark, output_dir)
    assert os.path.exists(path_dark)
