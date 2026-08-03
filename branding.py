class BrandingEngine:

    def __init__(self):

        self.replacements = {
            "ams": "Infineon",
            "AMS": "Infineon",
        }

    def transform_text(self, text):

        for old, new in self.replacements.items():

            text = text.replace(old, new)

        return text

    def transform_document(self, model):

        for paragraph in model.paragraphs:

            paragraph.text = self.transform_text(
                paragraph.text
            )

        return model