from analyser import DocxAnalyzer
import json


def validate(source_docx, output_docx):

    src = DocxAnalyzer(source_docx).analyze()
    out = DocxAnalyzer(output_docx).analyze()

    critical_checks = [
        "tables",
        "images",
    ]

    warning_checks = [
        "paragraphs_non_empty",
        "drawings_total",
        "image_drawings",
        "inline_images",
        "floating_images",
        "textboxes",
        "headers",
        "footers",
        "charts",
        "ole_objects",
    ]

    report = {}

    print("\n========== VALIDATION REPORT ==========\n")

    critical_failed = 0
    warning_failed = 0

    # -----------------------------------
    # Critical Checks
    # -----------------------------------

    print("CRITICAL CHECKS")
    print("----------------")

    for item in critical_checks:

        source_value = src.get(item, 0)
        output_value = out.get(item, 0)

        passed = source_value == output_value

        report[item] = {
            "source": source_value,
            "output": output_value,
            "pass": passed,
            "severity": "CRITICAL"
        }

        if passed:

            print(
                f"[PASS] {item:<20} "
                f"Source={source_value} "
                f"Output={output_value}"
            )

        else:

            critical_failed += 1

            print(
                f"[FAIL] {item:<20} "
                f"Source={source_value} "
                f"Output={output_value}"
            )

    # -----------------------------------
    # Warning Checks
    # -----------------------------------

    print("\nWARNING CHECKS")
    print("--------------")

    for item in warning_checks:

        source_value = src.get(item, 0)
        output_value = out.get(item, 0)

        passed = source_value == output_value

        report[item] = {
            "source": source_value,
            "output": output_value,
            "pass": passed,
            "severity": "WARNING"
        }

        if passed:

            print(
                f"[PASS] {item:<20} "
                f"Source={source_value} "
                f"Output={output_value}"
            )

        else:

            warning_failed += 1

            print(
                f"[WARN] {item:<20} "
                f"Source={source_value} "
                f"Output={output_value}"
            )

    # -----------------------------------
    # Final Status
    # -----------------------------------

    if critical_failed == 0:

        overall_status = "PASS"

        if warning_failed > 0:
            overall_status = "PASS WITH WARNINGS"

    else:

        overall_status = "FAIL"

    print("\n======================================")
    print(f"Overall Status : {overall_status}")
    print(f"Critical Failed: {critical_failed}")
    print(f"Warning Failed : {warning_failed}")
    print("======================================\n")

    report["summary"] = {
        "status": overall_status,
        "critical_failed": critical_failed,
        "warning_failed": warning_failed
    }

    return report


if __name__ == "__main__":

    import sys

    if len(sys.argv) != 3:

        print(
            "Usage:\n"
            "python validator.py source.docx output.docx"
        )

        sys.exit(1)

    result = validate(
        sys.argv[1],
        sys.argv[2]
    )

    print(
        json.dumps(
            result,
            indent=4
        )
    )