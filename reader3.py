import csv, sys, re, os, xlwt

class Reader3(object):
    def __init__(self):
        self.to_do = 0
        self.doing = 0
        self.done = 0
        self.fornecedores = 0
        self.despriorizado = 0
        self.block = 0
        self.paraquedas_todo = 0
        self.paraquedas_doing = 0
        self.paraquedas_done = 0
        self.paraquedas_fornecedores = 0
        self.paraquedas_block = 0
        self.paraquedas_despriorizado = 0
        self.burndown = {}


    def read_file(self):
        pattern = re.compile("^\d+")
        nome_ficheiro = 'teste.csv'
        cont = 1
        workbook = xlwt.Workbook()

        worksheet_metricas = workbook.add_sheet(u'Metricas', cell_overwrite_ok=True)

        worksheet_metricas.write(0,0,u'TO DO')
        worksheet_metricas.write(1,0,u'DOING')
        worksheet_metricas.write(2,0,u'DONE')
        worksheet_metricas.write(3,0,u'BLOCK')
        worksheet_metricas.write(4,0,u'FORNECEDORES')
        worksheet_metricas.write(5,0,u'DESPRIORIZADOS')
        worksheet_metricas.write(6,0,u'=======')
        worksheet_metricas.write(7,0,u'PARAQUEDAS TODO')
        worksheet_metricas.write(8,0,u'PARAQUEDAS DOING')
        worksheet_metricas.write(9,0,u'PARAQUEDAS DONE')
        worksheet_metricas.write(10,0,u'PARAQUEDAS BLOCK')
        worksheet_metricas.write(11,0,u'PARAQUEDAS FORNECEDORES')
        worksheet_metricas.write(12,0,u'PARAQUEDAS DESPRIORIZADO')



        worksheet_dados = workbook.add_sheet(u'Dados', cell_overwrite_ok=True)
        worksheet_dados.write(0,0,u'PONTOS')
        worksheet_dados.write(0,1,u'NAME')
        worksheet_dados.write(0,2,u'STATUS')
        worksheet_dados.write(0,3,u'TAG')
        worksheet_dados.write(0,4,u'DATA DE CRIAÇÃO')
        worksheet_dados.write(0,5,u'DATA DE CONCLUSÃO')

        worksheet_metricas.write(0,5,u'BURNDOWN')
        worksheet_metricas.write(1,5,u'DATA')
        worksheet_metricas.write(1,6,u'PONTOS')

        with open(nome_ficheiro, 'rt', encoding='utf8') as ficheiro:
            reader = csv.reader(ficheiro)
            for linha in reader:
                if linha[13] == '':

                    result = re.search(pattern, linha[4])
                    if result :
                       if re.search('PARAQUEDAS', linha[10]):
                          if linha[5] == 'BLOCK':
                              self.paraquedas_block += int(result.group(0))
                          elif linha[5] == 'TO DO':
                              self.paraquedas_todo += int(result.group(0))
                          elif linha[5] == 'DOING':
                              self.paraquedas_doing += int(result.group(0))
                          elif linha[5] == 'FORNECEDORES':
                              self.paraquedas_fornecedores += int(result.group(0))
                          elif linha[5] == 'DESPRIORIZADOS':
                              self.paraquedas_despriorizado += int(result.group(0))
                          elif linha[5] == 'DONE':
                              self.paraquedas_done += int(result.group(0))
                       else:
                          if linha[5] == 'BLOCK':
                              self.block += int(result.group(0))
                          elif linha[5] == 'TO DO':
                              self.to_do += int(result.group(0))
                          elif linha[5] == 'DOING':
                              self.doing += int(result.group(0))
                          elif linha[5] == 'FORNECEDORES':
                              self.fornecedores += int(result.group(0))
                          elif linha[5] == 'DESPRIORIZADOS':
                              self.despriorizado += int(result.group(0))
                          elif linha[5] == 'DONE':
                              self.done += int(result.group(0))

                       if linha[2] != '':
                          # print(linha[2])
                           if linha[2] in self.burndown:
                               self.burndown[linha[2]] = self.burndown[linha[2]] + int(result.group(0))
                           else:
                                self.burndown[linha[2]] = int(result.group(0))

                       worksheet_dados.write(cont,0,int(result.group(0)))
                       worksheet_dados.write(cont,1,linha[4])
                       worksheet_dados.write(cont,2,linha[5])
                       worksheet_dados.write(cont,3,linha[10])
                       worksheet_dados.write(cont,4,linha[1])
                       worksheet_dados.write(cont,5,linha[2])
                    else:
                       worksheet_dados.write(cont,1,linha[4])
                       worksheet_dados.write(cont,2,linha[5])
                       worksheet_dados.write(cont,3,linha[10])
                       worksheet_dados.write(cont,4,linha[1])
                       worksheet_dados.write(cont,5,linha[2])

                    cont+=1

        print(self.burndown)

        worksheet_metricas.write(0,1,self.to_do)
        worksheet_metricas.write(1,1,self.doing)
        worksheet_metricas.write(2,1,self.done)
        worksheet_metricas.write(3,1,self.block)
        worksheet_metricas.write(4,1,self.fornecedores)
        worksheet_metricas.write(5,1,self.despriorizado)
        worksheet_metricas.write(6,1,"=======")
        worksheet_metricas.write(7,1,self.paraquedas_todo)
        worksheet_metricas.write(8,1,self.paraquedas_doing)
        worksheet_metricas.write(9,1,self.paraquedas_done)
        worksheet_metricas.write(10,1,self.paraquedas_block)
        worksheet_metricas.write(11,1,self.paraquedas_fornecedores)
        worksheet_metricas.write(12,1,self.paraquedas_despriorizado)

        cont = 2
        for dia in self.burndown:
            worksheet_metricas.write(cont,5,dia)
            worksheet_metricas.write(cont,6,self.burndown[dia])
            cont+=1

        workbook.save('metrica.xls')

if "__main__" == __name__:
    reader = Reader3()
    reader.read_file()
