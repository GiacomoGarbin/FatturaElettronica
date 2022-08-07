# imports

import sys
import traceback
import xml.etree.ElementTree as ET
import xlsxwriter
import win32api

# column values

class ColumnValue:
    def __init__(self, name, label):
        self.name = name
        self.label = label
        self.width = len(label)

    def UpdateWidth(self, value):
        self.width = max(self.width, len(value))

    def GetValue(self, row):
        found = row.find(self.name)
        value = found.text if (found != None) else str()
        self.UpdateWidth(value)
        return value

class Descrizione(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Descrizione")

class Commento(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Commento")

    def GetValue(self, row):
        comment = []
        for item in row.iter('AltriDatiGestionali'):
            found_type = item.find('TipoDato')
            if (found_type != None) and (found_type.text.upper() == 'COMMENTO'):
                found_data = item.find('RiferimentoTesto')
                if (found_data != None):
                    self.UpdateWidth(found_data.text)
                    comment.append(found_data.text)
        return '\n'.join(comment)

class CodiceArticolo(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Codice Articolo")

    def GetValue(self, row):
        value = str()

        found_article = row.find('CodiceArticolo')
        if (found_article != None):
            found_type = found_article.find('CodiceTipo')
            found_value = found_article.find('CodiceValore')

            if (found_type != None) and (found_value != None):
                value = '%s (%s)' % (found_value.text, found_type.text)
            elif (found_type != None):
                value = found_type.text
            elif (found_value != None):
                value = found_value.text

            # add some padding to get the right width
            self.UpdateWidth(value + '  ')
        
        return value

class DatiDDT(ColumnValue):
    def __init__(self, root):
        super().__init__(type(self).__name__, "DDT")

        self.ddt = []

        for item in root.iter('DatiDDT'):

            ddt = dict()

            found_number = item.find('NumeroDDT')
            found_date = item.find('DataDDT')

            ddt['number'] = found_number.text if found_number != None else None
            ddt['date'] = found_date.text if found_date != None else None
            ddt['lines'] = set(int(i.text) for i in item.iter('RiferimentoNumeroLinea') if i != None)
            
            for other in self.ddt:
                if len(ddt['lines'].intersection(other['lines'])):
                    raise Exception('len(ddt.intersection(other)) > 0')

            self.ddt.append(ddt)

    def GetValue(self, row):
        value = str()

        line = row.find('NumeroLinea')
        if line != None: 
            try:
                line = int(line.text)
            except ValueError:
                return str()

            for ddt in self.ddt:
                if line in ddt['lines']:
                    
                    if (ddt['number'] != None) and (ddt['date'] != None):
                        value = '%s del %s' % (ddt['number'], ddt['date'])
                    elif (ddt['number'] != None):
                        value = ddt['number']
                    elif (ddt['date'] != None):
                        value = ddt['date']
                    
                    self.UpdateWidth(value)
                    break

        return value

class UnitaMisura(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Unità Misura")

class Quantita(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Quantità")

    def GetValue(self, row):
        value = super().GetValue(row)

        try:
            value = float(value)
        except ValueError:
            return str()

        value = int(value)
        self.UpdateWidth(str(value))
        return value

class PrezzoUnitario(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Prezzo Unitario")
    
    def GetValue(self, row):
        value = super().GetValue(row)

        try:
            # float(value)
            value = float(value)
        except ValueError:
            return str()

        # i, f = value.split('.')
        # n = len(f)
        # k = n - 2
        # j = 0
        # while j < k:
        #     if f[(n-1)-j] > '0':
        #         break
        #     j += 1
        # return "%s.%s" % (i, f[:n-j])

        self.UpdateWidth(str(value))
        return value

class AliquotaIVA(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Aliquota IVA")
    
    def GetValue(self, row):
        value = super().GetValue(row)

        try:
            value = float(value)
        except ValueError:
            return str()

        # return "%.0f%%" % value
        
        value = int(value)
        self.UpdateWidth(str(value))
        return value

class PrezzoTotale(ColumnValue):
    def __init__(self):
        super().__init__(type(self).__name__, "Prezzo Totale")
    
    def GetValue(self, row):
        value = super().GetValue(row)

        try:
            value = float(value)
        except ValueError:
            return str()

        # return "%.2f" % value

        self.UpdateWidth(str(value))
        return value

class GeneralError(Exception):
    pass

# main

script_name = sys.argv[0].split('.')[0]

try:

    if len(sys.argv) < 2:
        raise GeneralError('no filename')

    in_filename = sys.argv[1]
    root = None

    try:
        tree = ET.parse(in_filename)
        root = tree.getroot()
    except ET.ParseError:
        raise GeneralError('failed to parse "%s"' % in_filename)

    if root.find('DatiBeniServizi'):
        raise GeneralError('no "DatiBeniServizi" elements in "%s"' % in_filename)

    def IsXML(filename):
        return filename.split('.')[-1].lower() == 'xml'
    
    out_filename = in_filename[:-4] if IsXML(in_filename) else in_filename
    workbook = xlsxwriter.Workbook(out_filename + '.xlsx')

    columns = [
        Descrizione(),
        Commento(),
        CodiceArticolo(),
        DatiDDT(root),
        UnitaMisura(),
        Quantita(),
        PrezzoUnitario(),
        AliquotaIVA(),
        PrezzoTotale()]
        

    # bold format for header fields
    bold = workbook.add_format({'bold': True})
    # vertical alignment
    valign = workbook.add_format({'valign': 'top'})
    valign.set_text_wrap()

    for table in root.iter('DatiBeniServizi'):
        # new sheet
        worksheet = workbook.add_worksheet()

        # write header
        for i in range(len(columns)):
            worksheet.write(0, i, columns[i].label, bold)

        # write rows
        i = 1
        for row in table.iter('DettaglioLinee'):
            height = 1
            for j, col in enumerate(columns):
                value = col.GetValue(row)
                height = max(height, len(str(value).split('\n')))
                worksheet.write(i, j, value)
            # adjust row height and alignment
            worksheet.set_row(i, 15*height, valign)
            i += 1

        # adjust columns width
        for j, col in enumerate(columns):
            worksheet.set_column(j, j, col.width)

    workbook.close()

except GeneralError as exc:
    win32api.MessageBox(0, str(exc), script_name)
except Exception:
    win32api.MessageBox(0, traceback.format_exc(), script_name)
    

# excel autofit

# import win32com.client as win32
# excel = win32.gencache.EnsureDispatch('Excel.Application')
# wb = excel.Workbooks.Open(r'file.xlsx')
# ws = wb.Worksheets("Sheet1")
# ws.Columns.AutoFit()
# wb.Save()
# excel.Application.Quit()

# header = dict()

# def Walker(obj, node):
#     if node.get(obj.tag) == None:
#         node[obj.tag] = dict()
#     for child in obj:
#         Walker(child, node[obj.tag])

# workbook = xlsxwriter.Workbook('IT01641790702_zPBc0.xlsx')
# for table in root.iter('DatiBeniServizi'):
#     worksheet = workbook.add_worksheet()
#     for row in table.iter('DettaglioLinee'):
#         for obj in row:
#             Walker(obj, header)
# workbook.close()

# def WalkerPrint(node, level=0):
#     for k, v in node.items():
#         print("%s%s" % ("\t"*level, k))
#         WalkerPrint(v, level+1)

# WalkerPrint(header)

# NumeroLinea
# CodiceArticolo
#         CodiceTipo
#         CodiceValore
# Descrizione
# Quantita
# UnitaMisura
# PrezzoUnitario
# PrezzoTotale
# AliquotaIVA
# AltriDatiGestionali
#         TipoDato
#             COMMENTO
#         RiferimentoTesto
# TipoCessionePrestazione

# for table in root.iter('DatiBeniServizi'):
#     for TipoDato in table.iter('TipoDato'):
#         print(TipoDato.text)