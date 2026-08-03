from docx import Document


class Builder:

    def build(self, model, out_file):

        doc = Document()

        # -------------------
        # Paragraphs
        # -------------------

        for paragraph in model.paragraphs:

            doc.add_paragraph(
                paragraph.text
            )

        # -------------------
        # Tables
        # -------------------

        for table in model.tables:

            if not table.data:
                continue

            rows = len(table.data)

            cols = max(
                len(row)
                for row in table.data
            )

            word_table = doc.add_table(
                rows=rows,
                cols=cols
            )

            for r, row_data in enumerate(table.data):

                for c, value in enumerate(row_data):

                    word_table.cell(
                        r,
                        c
                    ).text = value

        # -------------------
        # Images
        # -------------------

        for image in model.images:

            try:

                doc.add_picture(
                    image.path
                )

            except Exception as e:

                print(
                    f"Could not add image "
                    f"{image.name}: {e}"
                )

        doc.save(out_file)