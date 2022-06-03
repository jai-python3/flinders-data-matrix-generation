#!/bin/bash
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" &> /dev/null && pwd )"
. $SCRIPT_DIR/../venv/bin/activate
python $SCRIPT_DIR/process_dr_worksheet.py $@
