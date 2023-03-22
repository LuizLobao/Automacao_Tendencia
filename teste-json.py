from datetime import date, datetime, timedelta
import json

hoje = (datetime.today()- timedelta(days=0)).strftime('%d/%m/%Y') 
dicio = {hoje: [
                {
                "BOV_1066.TXT":'01/01/1900 00:00:00',
                "BOV_1065.TXT":'01/01/1900 00:00:00',
                "BOV_1059.TXT":'01/01/1900 00:00:00',
                "BOV_1067.TXT":'01/01/1900 00:00:00',
                "BOV_1064.TXT":'01/01/1900 00:00:00',
                "BOV_1058.TXT":'01/01/1900 00:00:00',
                "HADOOP_6162.TXT":'01/01/1900 00:00:00',
                "HADOOP_6163.TXT":'01/01/1900 00:00:00'
                }
                ]
        }
#print(dicio)

json_string = json.dumps(dicio, indent=2, sort_keys=True)

with open('cargas.json','a') as f:
    f.write(json_string)


with open('cargas.json', 'r') as g:
    dicionario = json.loads(g.read())

print(dicionario)
print('')
print('')
print(dicionario.keys())
print('')
print('')
print(dicionario['22/03/2023'][0])
print('')
print('')
print(dicionario['22/03/2023'][0]['BOV_1065.TXT'])
print('')
print('')
print(dicionario['22/03/2023'][0].keys())
print(dicionario['22/03/2023'][0].values())