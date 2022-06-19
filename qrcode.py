import qrcode

def criar_img(_valor):
    data =f'O valor da carteira é R${_valor}'
    img = qrcode.make(data)

    img.save('qrcode.png')
