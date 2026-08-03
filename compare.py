import json

from validator import Validator

import sys


if __name__ == "__main__":

    source = sys.argv[1]
    output = sys.argv[2]

    report = Validator.compare(
        source,
        output
    )

    print(
        json.dumps(report, indent=4)
    )