from analyser import DocxAnalyzer


def validate(source_docx, output_docx):
    source_report = DocxAnalyzer(source_docx).analyze()
    output_report = DocxAnalyzer(output_docx).analyze()

    report = {
        "source": source_report,
        "output": output_report,
        "differences": {}
    }

    for key in source_report.keys():
        if source_report[key] != output_report[key]:
            report["differences"][key] = {
                "source": source_report[key],
                "output": output_report[key]
            }

    return report