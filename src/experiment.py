from logger import AppLogger
from parse_docx import parse_docx
import pandas as pd

# pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)
pd.set_option('display.max_colwidth', -1)

logger = AppLogger("test")
parsed = parse_docx("/tmp/extract", logger)(
    "/home/xarlyle0/Work/Repos/CALM/test-docs/Sample_RFP_70US0921Q70090020.docx")

parsed
