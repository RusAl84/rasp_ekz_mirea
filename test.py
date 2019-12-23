def str_count(text, substr):    #количество вхождений подстроки в строку
    return len(text.split(substr))-1


text='БАСО-01-19 10.05.03(КБ-1)'
substr='БАСО0'
print(str_count(text, substr))

#5	8	11	65
#for i in range(5,68,3):
#    print(i)
#mas=[5, 8, 11, 15, 18, 21, 24, 27, 30, 34, 37, 40, 43, 46, 49, 53, 56, 59, 62, 65]
#for zz in range(0,len(mas)):
#    z=mas[zz]
#    print(z)