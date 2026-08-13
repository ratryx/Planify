import os
from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.registry.template_registry import registry
from excel_saas.core.engine.workbook_engine import WorkbookEngine
from excel_saas.themes.light import LightTheme
from excel_saas.themes.dark import DarkTheme

# Auto-register templates
import excel_saas.templates.finance_personal.template

def generate(request: GenerationRequest, output_dir: str):
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Get Template
    template = registry.get(request.template_id)
    
    # 2. Build Plan
    plan = template.build_workbook_plan(request)
    
    # 3. Resolve Theme
    theme = LightTheme() if request.theme == "light" else DarkTheme()
    
    # 4. Render
    filename = f"{request.template_id}_{request.year}_{request.theme}.xlsx"
    filepath = os.path.join(output_dir, filename)
    
    engine = WorkbookEngine(plan, theme)
    engine.render(filepath)
    
    return filepath

if __name__ == "__main__":
    import sys
    # Example execution
    output = "output"
    
    req_light = GenerationRequest(template_id="finance_personal", year=2026, theme="light")
    print(f"Generating Light theme to {output}...")
    path1 = generate(req_light, output)
    print(f"Done: {path1}")
    
    req_dark = GenerationRequest(template_id="finance_personal", year=2026, theme="dark")
    print(f"Generating Dark theme to {output}...")
    path2 = generate(req_dark, output)
    print(f"Done: {path2}")
