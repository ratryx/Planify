import os
import sys
import shutil

# Add project root to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..")))

from excel_saas.core.models.generation_request import GenerationRequest
from excel_saas.core.engine.workbook_engine import WorkbookEngine
from excel_saas.templates.finance_personal.template import FinancePersonalTemplate
from excel_saas.themes.light import LightTheme
from excel_saas.themes.dark import DarkTheme

# Also add tests path to import fixture
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "tests")))
from fixtures.realistic_household_state import inject_realistic_household

def main():
    output_dir = os.path.abspath("runtime-output")
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    template = FinancePersonalTemplate()
    # A. light, empty product data, horizon=12
    req_a = GenerationRequest(
        template_id="finance_personal",
        year=2026,
        theme="light",
        with_sample_data=False,
        projection_horizon=12
    )
    plan_a = template.build_workbook_plan(req_a)
    path_a = os.path.join(output_dir, "finance_personal_runtime_a.xlsx")
    engine_a = WorkbookEngine(plan_a, LightTheme())
    engine_a.render(path_a)
    print(f"Generated {path_a}")

    # B. dark, realistic household fixture, horizon=12
    req_b = GenerationRequest(
        template_id="finance_personal",
        year=2026,
        theme="dark",
        with_sample_data=False,
        projection_horizon=12
    )
    plan_b = template.build_workbook_plan(req_b)
    inject_realistic_household(plan_b, 2026)
    path_b = os.path.join(output_dir, "finance_personal_runtime_b.xlsx")
    engine_b = WorkbookEngine(plan_b, DarkTheme())
    engine_b.render(path_b)
    print(f"Generated {path_b}")

    # C. light, realistic household fixture, horizon=60
    req_c = GenerationRequest(
        template_id="finance_personal",
        year=2026,
        theme="light",
        with_sample_data=False,
        projection_horizon=60
    )
    plan_c = template.build_workbook_plan(req_c)
    inject_realistic_household(plan_c, 2026)
    path_c = os.path.join(output_dir, "finance_personal_runtime_c.xlsx")
    engine_c = WorkbookEngine(plan_c, LightTheme())
    engine_c.render(path_c)
    print(f"Generated {path_c}")

if __name__ == "__main__":
    main()
