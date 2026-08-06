from pathlib import Path

from docx_migration.extractor import Extractor
from docx_migration.builder import Builder

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE = str(PROJECT_ROOT / "samples" / "real_ams.docx")
OUTPUT = str(PROJECT_ROOT / "samples" / "roundtrip_output.docx")

model = Extractor(
    SOURCE
).extract()

Builder(project_root=PROJECT_ROOT).build(
    model,
    OUTPUT
)

print(
    "Round-trip completed successfully"
)