import requests


### close PYQT App

# pyinstaller --onefile --noconsole --name CapZa --icon .\assets\icon_logo.ico .\capza\main.py

### Delete PYQT APP

### Download installer
url = 'https://www.mediafire.com/file/b5lky932dj1zhzt/CapZa_0.1.1.zip'
r = requests.get(url, allow_redirects=True)

open('../CapZa-0.1.1.zip', 'wb').write(r.content)



### Update Versionsnummer

### Open

