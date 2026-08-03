from extractor import Extractor
from branding import BrandingEngine
from builder import Builder


SOURCE = "input.docx"
OUTPUT = "output.docx"


model = Extractor(SOURCE).extract()

model = BrandingEngine().transform_document(
    model
)

Builder().build(
    model,
    OUTPUT
)