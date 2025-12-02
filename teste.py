import datetime



data_atual = datetime.datetime.now()
print(data_atual)
horas =  data_atual.hour
print(horas)

if int(horas) < 18:
   data_inicial = (data_atual - datetime.timedelta(1)).strftime("%d/%m/%Y")
   data_final = data_inicial
else:
    data_inicial = data_atual.strftime("%d/%m/%Y")
    data_final = data_inicial
print(data_inicial)