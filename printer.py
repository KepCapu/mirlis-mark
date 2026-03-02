import win32print

def print_raw(printer_name: str, data: str):
    hPrinter = win32print.OpenPrinter(printer_name)
    try:
        job = win32print.StartDocPrinter(hPrinter, 1, ("MirlisMarkLabel", None, "RAW"))
        try:
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, data.encode("utf-8"))
            win32print.EndPagePrinter(hPrinter)
        finally:
            win32print.EndDocPrinter(hPrinter)
    finally:
        win32print.ClosePrinter(hPrinter)