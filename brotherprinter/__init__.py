class PrintException(Exception):
    pass


class BrotherPrinter:
    def __init__(self):
        self.constants = {
            'bpoHighResolution': 0x02000000
        }
        self.default_printer = None

    def get_printer_info(self, bpac):
        try:
            printer = self.default_printer
            if printer:
                support = True if bpac.Printer.IsPrinterSupported(printer) else False
                online = True if bpac.Printer.IsPrinterOnline(printer) else False
                return {
                    'printer_info': {
                        'name': printer,
                        'is_supported': support,
                        'online': online
                    }
                }
            else:
                return None
        except Exception as error:
            raise(PrintException(error))

    def get_label_info(self, bpac):
        try:
            printer = self.default_printer
            bpac.SetPrinter(printer, False)
            id = bpac.Printer.GetMediaId
            name = bpac.Printer.GetMediaName
            return {
                'label_info': {
                    'id': id,
                    'name': name if name else "Ingen media"
                }
            }
        except Exception as error:
            raise (PrintException(error))

    @staticmethod
    def get_available_printer(bpac):
        try:
            printers = bpac.Printer.GetInstalledPrinters
            for printer in printers:
                if bpac.Printer.IsPrinterSupported(printer) and \
                        bpac.Printer.IsPrinterOnline(printer):
                    return printer
            else:
                return None
        except Exception as error:
            raise (PrintException(error))

    def set_default_printer(self, bpac):
        try:
            self.default_printer = self.get_available_printer(bpac)
        except Exception as error:
            raise (PrintException(error))

    @staticmethod
    def get_printers(bpac):
        try:
            return bpac.Printer.GetInstalledPrinters
        except Exception as error:
            raise (PrintException(error))

    def print_callnumber(self, bpac, *, template, call_number, printer=None):
        try:
            printer = printer if printer else self.default_printer
            if not printer:
                raise PrintException('Ingen skrivare är tillgänglig')

            has_opened = bpac.Open(template)
            if not has_opened:
                raise PrintException('Mallfilen på sökvägen "' + template + '" är inte tillgänglig')

            if not bpac.StartPrint("", self.constants['bpoHighResolution']):
                raise PrintException('Misslyckades med att starta utskrift')

            bpac.GetObject("CallNumber").Text = call_number
            PrintOut = bpac.PrintOut(1, self.constants['bpoHighResolution'])

            if not PrintOut:
                raise PrintException('Misslyckades med att skriva ut (Felkod ' + bpac.ErrorCode + ')')
            if not bpac.EndPrint:
                raise PrintException('Misslyckades med att avsluta utskrift')
            if not bpac.Close:
                raise PrintException('Misslyckades med att stänga skrivaren')

            return True
        except Exception as error:
            raise (PrintException(error))
