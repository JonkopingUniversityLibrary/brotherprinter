class PrintException(Exception):
    pass


class BrotherPrinter:
    def __init__(self):
        self.constants = {
            'bpoHighResolution': 0x02000000
        }

    def get_printer_info(self, bpac):
        for printer in bpac.Printer.GetInstalledPrinters:
            support = True if bpac.Printer.IsPrinterSupported(printer) else False
            status = "Online" if bpac.Printer.IsPrinterOnline(printer) else "Offline"
            return {
                'printer_info': {
                    'name': printer,
                    'is_supported': support,
                    'status': status
                }
            }

    def get_label_info(self, bpac):
        for printer in bpac.Printer.GetInstalledPrinters:
            bpac.SetPrinter(printer, False)

            id = bpac.Printer.GetMediaId
            name = bpac.Printer.GetMediaName
            return {
                'label_info': {
                    'id': id,
                    'name': name if name else "Ingen media"
                }
            }

    def get_available_printer(self, bpac):
        printers = bpac.Printer.GetInstalledPrinters
        for printer in printers:
            if bpac.Printer.IsPrinterSupported(printer) \
                    and bpac.Printer.IsPrinterOnline(printer):
                return printer
        else:
            return None

    def print_callnumber(self, bpac, *, template, call_number):
        printer = self.get_available_printer(bpac)
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
