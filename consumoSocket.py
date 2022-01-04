import websocket, json, xlwings, pandas, time


#Aqui coloco o endpoint do websocket!! para obter mais endpoint consulte https://binance-docs.github.io/apidocs/spot/en/#websocket-market-streams
socket = "wss://stream.binance.com:9443/ws/!ticker@arr"
#Aqui abro uma planilha do excel. Obs.: É importante que a planilha esteja na mesma página do arquivo python
wb = xlwings.Book('dadosSocket.xlsx').sheets('Sheet1')

#Aqui inicio a função do websocket, ele tem 4 parametros:
#on_message -> Recebe a resposta da solicitação do endpoint
#on_error -> Caso a solicitação falhe ou dê algum problema ele recebe o erro
#on_close -> Quando o websocket é fechado
#on_open -> É chamada quando o conexão é iniciada

def on_message(ws, message):
    #Aqui carrego os dados recebidos do websocket
    dados = json.loads(message)

    #Aqui passo os dados para o dataFrame para poder tratar
    dadosTabulados = pandas.DataFrame(dados)

    #Aqui renomeio as colunas
    dadosTabulados.columns = ['Tipo', 'Tempo', 'Símbolo', 'Mudança de preço', 'Porcentagem de mudança de preço', 
    'Preço médio ponderado', 'primeira negociação ', 'Último preço', 'Última quantidade', 'Preço do melhor lance', 'Melhor quantidade de oferta',
    'Melhor preço de venda', 'Melhor quantidade de perguntas', 'Preço aberto', 'Preço Alto', 'Preço baixo', 'Volume de base total de ativos negociados',
    'Volume total de ativos de cotação negociados', 'Tempo de abertura das estatísticas', 'Tempo de fechamento das estatísticas',
    'ID da primeira negociação', 'Última Id de negociação','Número total de negociações']

    #Aqui excluo as colunas tipo e tempo
    dadosTabulados.drop(['Tipo', 'Tempo'], axis=1, inplace=True)

    #Aqui gravos as alterações na planilha em excel
    #Caso queira exibir os dados sem alterar as colunas, basta substituir o atributo dadosTabulados por dados
    wb.range('a1').options(pandas.DataFrame, index=False).value = dadosTabulados
    time.sleep(5)
    print("Recebendo informações da binance....")

def on_error(ws, error):
    print(error)

def on_close(ws, close_status_code, close_msg):
    print("Conexão Fechada")
    

def on_open(ws):
    print("Iniciando a Conexão")
    def run(*args):
        for i in range(3):
            time.sleep(1)
        time.sleep(1)
        ws.close()
        print("thread terminating...")
    _thread.start_new_thread(run, ())
    

if __name__ == "__main__":
    
     ws = websocket.WebSocketApp(socket,
                                on_message=on_message,
                                on_error=on_error,
                                on_close=on_close,)
     ws.on_open = on_open
    
     ws.run_forever()
