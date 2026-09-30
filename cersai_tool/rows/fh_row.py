"""FH row: the file header. Written once at the top of every .DAT file.
4 columns (sheet "introduction" of the format description)."""
from ..model import derived, fixed

FIELDS = ["FILE_HEADER", "FILE_TYPE", "NUMBER_OF_RECORDS", "FILE_DATE"]


# Column 1 - FILE_HEADER
# Always the text "FH".
def col_01_file_header(number_of_records, file_date, row):
    return fixed("FILE_HEADER", "FH")


# Column 2 - FILE_TYPE
# Always the text "SI_REGN" (security interest registration).
def col_02_file_type(number_of_records, file_date, row):
    return fixed("FILE_TYPE", "SI_REGN")


# Column 3 - NUMBER_OF_RECORDS
# The number of security interests (collaterals) in this file, 1 to 1000.
# In the Excel files it counts every collateral of the batch; when the .DAT
# file is written (step 3) it is recounted to the SIs actually included.
def col_03_number_of_records(number_of_records, file_date, row):
    return derived("NUMBER_OF_RECORDS", number_of_records, "number of SI blocks in the file")


# Column 4 - FILE_DATE
# The date the file is created, as YYYY-MM-DD (never a future date).
# In the Excel files this is the day the Excel was made; step 3 replaces it
# with the day the .DAT file is written.
def col_04_file_date(number_of_records, file_date, row):
    return derived("FILE_DATE", file_date, "date the file is created")


COLUMNS = [col_01_file_header, col_02_file_type, col_03_number_of_records, col_04_file_date]


def build(number_of_records: int, file_date: str):
    from . import build_row
    return build_row(FIELDS, COLUMNS, number_of_records, file_date)
