@echo off
if '%1' == '1' goto ONE
if '%1' == '2' goto TWO

:ONE
    poetry run python extractxlsx.py
    goto END

:TWO
    poetry run python ExcelTestCaseExtractor.py
    goto END

:END