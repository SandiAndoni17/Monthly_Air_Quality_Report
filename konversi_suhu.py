print('##  Program Python Konversi Suhu  ##')
print('====================================')
print()

celc = float(input('Input suhu celsius: '))

fahr = (9 / 5 * celc) + 32
kelv = celc + 273.15
ream = celc * (4 / 5)

print(celc, 'derajat Celsius =', fahr, 'derajat Fahrenheit')
print(celc, 'derajat Celsius =', kelv, 'derajat Kelvin')
print(celc, 'derajat Celsius =', ream, 'derajat Reamur')