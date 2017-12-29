import requests
import re
import openpyxl
import time

class RegExAttributes:

    def __init__(self,regex,number):
        self.regex  = regex
        self.number = number

class USDAttributes:

    def __init__(self):
        self.__url                  = "http://finanzasenlinea.infosel.com/ElFinanciero/" \
                                      "IndicadoresMercadosDolar.aspx?i=COMPOSITE%2048H"
        self.regex_tup              = (RegExAttributes('LblValor">\\d+.\\d+',10),
                                       RegExAttributes('\\D\\d+.\\d+%',0),
                                       RegExAttributes('LblUltimoCierre">\\d+.\\d+',17),
                                       RegExAttributes('LblMinimoDia">\\d+.\\d+',14),
                                       RegExAttributes('LblMaximoDia">\\d+.\\d+',14)
                                       )
        self.utilities              = DataUtilities(self.__url)
    def search_values(self):
        temp_list = []
        for i in range(len(self.regex_tup)):
            temp_list.append(self.utilities.search_regex(self.regex_tup[i].regex))
            temp_list[i] = temp_list[i][self.regex_tup[i].number:]

        return temp_list

class GeneralIndexAttributes:

    def __init__(self):
        self.__url                  = "http://finanzasenlinea.infosel.com/ElFinanciero/" \
                                      "IndicadoresMercadosDestacados.html"

        self.regex_tup              = [RegExAttributes('id="".*',0),
                                       RegExAttributes('\\d+,\\d+.\\d+',0),
                                       RegExAttributes('\\D\\d+.\\d+%',0)
                                       ]
        self.utilities              = DataUtilities(self.__url)
    def search_values(self):
        temp_list = []
        text = self.utilities.search_regex(self.regex_tup[0].regex)
        for i in range(1,len(self.regex_tup)):
            temp_list.extend(self.utilities.search_regex(self.regex_tup[i].regex,'findall_ext',
                             text))
            temp_list[i] = temp_list[i][self.regex_tup[i].number:]

        return temp_list

class DJIAttributes(GeneralIndexAttributes):
    def __init__(self):
        GeneralIndexAttributes.__init__(self)
        self.regex_tup[0] =  RegExAttributes('id="DJI".*',0)

class IPCAttributes(GeneralIndexAttributes):
    def __init__(self):
        GeneralIndexAttributes.__init__(self)
        self.regex_tup[0] =  RegExAttributes('id="IPC".*',0)

class NASDAQAttributes(GeneralIndexAttributes):
    def __init__(self):
        GeneralIndexAttributes.__init__(self)
        self.regex_tup[0] =  RegExAttributes('id="COMPX".*',0)

class SandPAttributes(GeneralIndexAttributes):
    def __init__(self):
        GeneralIndexAttributes.__init__(self)
        self.regex_tup[0] =  RegExAttributes('id="INX".*',0)

class USDInterbancario:

    def __init__(self):
        self.__url                  = "http://www.banxico.org.mx/tipcamb/" \
                                      "llenarTiposCambioAction.do?idioma=sp"

        self.regex_tup              = [RegExAttributes('"INTER_DATO">\s*\d*.\d*',-8)]
        self.utilities              = DataUtilities(self.__url)
    def search_values(self):
        temp_list = []
        for i in range(len(self.regex_tup)):
            temp_list.append(self.utilities.search_regex(self.regex_tup[i].regex))
            temp_list[i] = temp_list[i][self.regex_tup[i].number:]

        return temp_list

class Monex:

    def __init__(self):
        self.__url                  = "https://www.monex.com.mx/portal/" 
        self.regex_tup              = (RegExAttributes('h5"><em>\d+.\d+',8),
                                       )
        self.utilities              = DataUtilities(self.__url)
    def search_values(self):
        temp_list = []
        for i in range(len(self.regex_tup)):
            temp_list.extend(self.utilities.search_regex(self.regex_tup[i].regex,'findall'))

        for j in range(len(temp_list)):
            temp_list[j] = temp_list[j][self.regex_tup[i].number:]

        temp_list = temp_list[:4]
        return temp_list


class DataUtilities:

    def __init__(self,url_string):
        self.res = ''
        self.url_string = url_string
        self.get_url_data()

    def get_url_data(self):
        try:
            self.res = requests.get(self.url_string)
        except:
            print("Invalid URL!\n")

    
    def search_regex(self,regex_pattern,option=None,txt=None):
        result = ''
        function_dict = {'normal':self.function_dict_normal,
                         'normal_ext':self.function_dict_normal_ext,
                         'findall':self.function_dict_findall,
                         'findall_ext':self.function_dict_findall_ext}

        regex = re.compile(regex_pattern)
        if option == None:
            option = 'normal'
        
        result = function_dict[option](regex,txt)

        return result

    def function_dict_normal(self,regular_expression,txt=None):
        mo = regular_expression.search(self.res.text)
        try:
            result = mo.group()
        except:
            print("No Results on Regex " + regex_pattern)

        return result

    def function_dict_normal_ext(self,regular_expression,txt=None):
        mo = regular_expression.search(txt)
        try:
            result = mo.group()
        except:
            print("No Results on Regex " + regex_pattern)

        return result

    def function_dict_findall(self,regular_expression,txt=None):
        mo = regular_expression.findall(self.res.text)
        result = mo

        return result

    def function_dict_findall_ext(self,regular_expression,txt=None):
        mo = regular_expression.findall(txt)
        result = mo

        return result

def build_workbook():
    wb = openpyxl.Workbook()
    sheet = wb.get_sheet_by_name('Sheet')
    sheet['A1'] = 'El Financiero'
    sheet['A2'] = 'USD'
    sheet['A3'] = 'Valor'
    sheet['A4'] = 'Variacion'
    sheet['A5'] = 'Cierre Anterior'
    sheet['A6'] = 'Minimo'
    sheet['A7'] = 'Maximo'
    sheet['A8'] = 'Dow Jones'
    sheet['A9'] = 'Actual'
    sheet['A10'] = 'Anterior'
    sheet['A11'] = 'Variacion'
    sheet['A12'] = 'IPC'
    sheet['A13'] = 'Actual'
    sheet['A14'] = 'Anterior'
    sheet['A15'] = 'Variacion'
    sheet['A16'] = 'NASDAQ'
    sheet['A17'] = 'Actual'
    sheet['A18'] = 'Anterior'
    sheet['A19'] = 'Variacion'
    sheet['A20'] = 'S&P'
    sheet['A21'] = 'Actual'
    sheet['A22'] = 'Anterior'
    sheet['A23'] = 'Variacion'
    sheet['C1'] = 'Banxico'
    sheet['C2'] = 'Interbancario 48 horas'
    sheet['E1'] = 'Monex'
    sheet['E2'] = 'USD Compra'
    sheet['E3'] = 'USD Venta'
    sheet['E4'] = 'Euro Compra'
    sheet['E5'] = 'Euro Venta'
    wb.save('indicadores'+time.strftime("%d.%m.%Y")+'.xlsx')

def flush_data(column,row,data):
    wb = openpyxl.load_workbook('indicadores'+time.strftime("%d.%m.%Y")+'.xlsx')
    sheet = wb.get_sheet_by_name('Sheet')
    for i in range(row,len(data)+row):
        sheet[column+str(i)]= data[i-row]
    wb.save('indicadores'+time.strftime("%d.%m.%Y")+'.xlsx')
    
if __name__ == "__main__":

    build_workbook()
    a = USDAttributes()
    val = a.search_values()
    flush_data('B',3,val)
    
    b = DJIAttributes()
    val = b.search_values()
    flush_data('B',9,val)
    
    c = IPCAttributes()
    val = c.search_values()
    flush_data('B',13,val)
    
    d = NASDAQAttributes()
    val = d.search_values()
    flush_data('B',17,val)
    
    e = SandPAttributes()
    val = e.search_values()
    flush_data('B',21,val)

    f = USDInterbancario()
    val = f.search_values()
    flush_data('D',2,val)
    
    g = Monex()
    val = g.search_values()
    flush_data('F',2,val)

    
    
    
