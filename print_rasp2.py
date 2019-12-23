import win32com.client

def str_count(text, substr):    #количество вхождений подстроки в строку
    return len(text.split(substr))-1


pos=2
fn=[u'c:\\rasp_ekz\\1.xlsx',
    u'c:\\rasp_ekz\\2.xlsx',
    u'c:\\rasp_ekz\\3.xlsx',
    u'c:\\rasp_ekz\\4.xlsx',
    u'c:\\rasp_ekz\\5.xlsx']

ofn=u'c:\\rasp_ekz\\ekz.xlsx'
mas=[5, 8, 11, 15, 18, 21, 24, 27, 30, 34, 37, 40, 43, 46, 49, 53, 56, 59, 62, 65]
prepods=['Алешкин', 'Алёшкин', 'Алексеенко', 'Волович', 'Выжигин', 'Головин', 'Ермакова', 'Корягин', 'Кульков', 'Лесько', 'Лось', 'Лукьянчиков', 'Мацнев', 'Мельников', 'Мерсов', 'Никольский', 'Никул','Нкульчев', 'Пушкин', 'Русаков', 'Серов', 'Филатов', 'Шпунт']


Excel = win32com.client.Dispatch("Excel.Application")
owb = Excel.Workbooks.Open(ofn)
osheet = owb.ActiveSheet
#for fi in range(0,5):
for fi in range(0,5):
    print(fn[fi])
    wb = Excel.Workbooks.Open(fn[fi])
    sheet = wb.ActiveSheet
    i=1
    while i<300:
        rw=3 # строка с шифрами групп
        #gr=str(currentSheet.cell(row=rw, column=i).value)
        gr=str(sheet.Cells(rw,i).value)
        
        if len(gr)>4:
            if str_count(gr, '-')>1:
               # print(gr,i)
                for zz in range(0,len(mas)):
                    z=mas[zz]
                    predm=str(sheet.Cells(z,i).value)
                    if len(predm)>4:
                        #print(predm,z)
                        kto=str(sheet.Cells(z+1,i).value)
                        #print(kto)
                        for x in range(0,len(prepods)):
                            if (str_count(kto, prepods[x])>0):
                                den=str(sheet.Cells(z-1,3).value)                      
                                gde=str(sheet.Cells(z-1,i+2).value)
                                gde=gde.replace('.0','')
                                vremya=str(sheet.Cells(z-1,i+1).value)
                                kurs=str(fi+1)
                                chto=str(sheet.Cells(z-1,i).value)
                                osheet.Cells(pos,1).value=kto
                                osheet.Cells(pos,2).value=den
                                osheet.Cells(pos,3).value=vremya
                                osheet.Cells(pos,4).value=predm
                                osheet.Cells(pos,5).value=chto       
                                osheet.Cells(pos,6).value=kurs
                                osheet.Cells(pos,7).value=gr
                                osheet.Cells(pos,8).value=gde
                                pos=pos+1
                                #print(kto+'  '+den+'  '+vremya+'  '+gr+'  '+ predm+'  '+chto+'  '+ gde+'  '+ vremya)
                                break
        i=i+1
    wb.Close()    
owb.Save()
owb.Close()
#закрываем COM объект
Excel.Quit()