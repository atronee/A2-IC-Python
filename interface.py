def interface():
    escolha = None
    while escolha != 2:
        print("""
=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
=                        Bem vindo ao Py.Bank                        =
=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
[ 1 ] Ver Carteira de Investimento
[ 2 ] Sair
=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
            """) # menu
        escolha = verifica_int("Escolha sua opção: ") #A função vai verificar se o número digitado é um int
        if escolha == 1:
            #Aqui o programa irá pedi a URL e verificará se ela existe e funciona.
            url = str(input("Por favor, digite a url da sua carteira: "))
            if verifica_url(url):
                print("pode chamar o módulo")
                #Caso exista, o módulo que desenvolve o excel com as informações será chamado
                break
            else:
                print("Desculpa, a carteira não foi encontrada")
        elif escolha == 2:
           break
        else:
            print("Por favor, escolha uma opção válida!")
    print("Obrigado, Volte Sempre!!")


def verifica_int(msg):
    """
    -> Verifica se o caracter passado no input é um int.
    :param msg: mensagem personalizada solicitando o usuário que digite um número inteiro
    :return: número inteiro
    """
    while True:
        try: #Ele tenta transformar a resposta do usuário em inteiro
            inteiro = int(input(msg))
        except: #Se houver erro, a looping continua até digitar um int
            print('ERRO! por favor, digite um inteiro válido!')

            print('''
=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
[ 1 ] Ver Carteira de Investimento
[ 2 ] Sair
=-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-==-=-=
            ''')
        else:
            break #Se ele tiver sucesso, o programa segue normalmente.
    return inteiro


def verifica_url(url):
    pass

interface()