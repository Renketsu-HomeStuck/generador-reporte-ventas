from PIL import Image

img = Image.open("icono.png")
img.save("icono.ico", format="ICO", sizes=[(16,16), (32,32), (48,48), (256,256)])