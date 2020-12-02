import olefile
from os import listdir
from os.path import isfile, join
#D:/workspace/python/add/Scripts/hwp5html  --output=D:/workspace/python/add "C:/Users/jhahn/Google 드라이브/2019ADD/ADD보고서양식.hwp"

home_dir = "C:/Users/jhahn/Google 드라이브/2019ADD/보고서양식"

onlyfiles = [f for f in listdir(home_dir) if isfile(join(home_dir, f))]

for file in onlyfiles:
    f = olefile.OleFileIO(join(home_dir,file))
    print(f.listdir())
    encoded_text = f.openstream('BodyText/Section0').read()
    print(encoded_text.decode('UTF-8'))

    encoded_text = f.openstream('PrvText').read()
    decoded_text = encoded_text.decode('UTF-16')
    print(decoded_text)