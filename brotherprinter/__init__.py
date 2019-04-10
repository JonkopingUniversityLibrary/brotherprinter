import win32com.client as win32com
import os


class PrintException(Exception):
    pass


class BrotherPrinter:
    def __init__(self, *, template_name):
        self.doc = win32com.DispatchEx('bpac.Document')
        self.printers = self.doc.Printer.GetInstalledPrinters
        self.enabled_printer = None
        self.template_name = template_name
        self.constants = {
            'bpoHighResolution': 0x02000000
        }

    def get_printer_info(self):
        for printer in self.printers:
            support = "Yes" if self.doc.Printer.IsPrinterSupported(printer) else "No"
            status = "Online" if self.doc.Printer.IsPrinterOnline(printer) else "Offline"
            return "{name} - Support: {support}, Status: {status}".format(name=printer, support=support, status=status)

    def get_label_info(self):
        for printer in self.printers:
            self.doc.SetPrinter(printer, False)

            id = self.doc.Printer.GetMediaId
            name = self.doc.Printer.GetMediaName
            return "Label - {id} : {name}".format(id=id, name=name) if name else "No Media"
        
    def get_available_printer(self):
        printers = self.doc.Printer.GetInstalledPrinters
        if printers:
            for printer in printers:
                if self.doc.Printer.IsPrinterSupported(printer) \
                        and self.doc.Printer.IsPrinterOnline(printer) \
                        and self.doc.Printer.GetMediaName:
                    return printer
            return None
        else:
            return None

    def set_printer(self, printer):
        if printer:
            self.doc.SetPrinter(printer, False)
            self.enabled_printer = printer
            return True
        return False

    def print_callnumber(self, call_number):
        # Set default printer
        if not self.enabled_printer:
            self.set_printer(self.get_available_printer())

        printer = self.enabled_printer
        if not printer:
            raise PrintException('No printers available')

        directory = os.path.abspath(os.path.dirname(__file__))
        template = os.path.join(directory, self.template_name)

        has_opened = self.doc.Open(template)
        if not has_opened:
            raise PrintException('No template with the name "' + self.template_name + '" available')

        if not self.doc.StartPrint("", self.constants['bpoHighResolution']):
            raise PrintException('Failed to start print')

        self.doc.GetObject("CallNumber").Text = call_number
        PrintOut = self.doc.PrintOut(1, self.constants['bpoHighResolution'])

        if not PrintOut:
            raise PrintException('Failed to Print Out (ErrorCode ' + self.doc.ErrorCode + ')')
        if not self.doc.EndPrint:
            raise PrintException('Failed to end print')
        if not self.doc.Close:
            raise PrintException('Failed to close print')

        return True


label_printer = BrotherPrinter(template_name='Dewey_12mm_16px.lbx')
print(label_printer.print_callnumber('073 Pythonscript'))

