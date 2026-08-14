import tempfile
import os
import sys

# Verify import source
import excel_saas
pkg_path = getattr(excel_saas, "__file__", None)
if pkg_path is None:
    # Namespace package
    pkg_path = list(excel_saas.__path__)[0]

if "src" in pkg_path or "src\\excel_saas" in pkg_path or "src/excel_saas" in pkg_path:
    print(f"ERROR: Imported from source tree instead of installed wheel: {pkg_path}")
    sys.exit(1)

print(f"Successfully imported excel_saas from: {pkg_path}")

from excel_saas.application.generate_workbook import generate
from excel_saas.core.models.generation_request import GenerationRequest

def main():
    with tempfile.TemporaryDirectory() as temp_dir:
        request = GenerationRequest(
            template_id="finance_personal",
            year=2026,
            locale="pt_BR",
            currency="BRL",
            theme="light",
            with_sample_data=False,
            profile="couple",
            reserve_months=6,
            projection_horizon=12,
        )
        
        path = generate(request, temp_dir)
        
        assert os.path.exists(path), f"File {path} does not exist"
        assert os.path.isfile(path), f"Path {path} is not a file"
        assert path.endswith(".xlsx"), f"Path {path} does not have .xlsx suffix"
        
        file_size = os.path.getsize(path)
        assert file_size > 0, f"File {path} is empty"
        
        filename = os.path.basename(path)
        expected_filename = "finance_personal_2026_light.xlsx"
        assert filename == expected_filename, f"Expected filename {expected_filename}, got {filename}"
        
        print(f"Smoke test successful: Generated {filename} ({file_size} bytes)")

if __name__ == "__main__":
    main()
