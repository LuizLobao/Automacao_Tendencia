import time

def continuar():
    try:
        resposta = input("Deseja continuar? (Sim ou Não) ")
        return resposta.lower() == "sim"
    except KeyboardInterrupt:
        return True

while True:
    if continuar():
        print("O código está sendo executado...")
        time.sleep(10)
    else:
        print("O código foi interrompido.")
        break
