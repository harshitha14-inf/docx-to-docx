from extractor import Extractor
from builder import Builder

SOURCE = "samples/tables_images_only.docx"
OUTPUT = "samples/roundtrip_output.docx"

model = Extractor(
    SOURCE
).extract()

Builder().build(
    model,
    OUTPUT
)

print(
    "Round-trip completed successfully"
)