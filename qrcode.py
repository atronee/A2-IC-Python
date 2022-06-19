import qrcode

valor = 1000
data =f'O valor da carteira é R${valor}'
img = qrcode.make(data)

img.save('qrcode.png')
