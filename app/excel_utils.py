from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

class ExcelValidationError(Exception):
    pass

def validate_xlsx_bytes(file_bytes: bytes) -> None:
    try:
        wb = load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
        wb.close()
    except InvalidFileException as exc:
        raise ExcelValidationError("文件不是合法的 .xlsx 格式，或文件内容已损坏。") from exc
    except Exception as exc:
        raise ExcelValidationError("Excel 文件解析失败，请确认文件未损坏且内容可读取。") from exc
