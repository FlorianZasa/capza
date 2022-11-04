import requests


### close PYQT App

# pyinstaller --onefile --noconsole --name CapZa --icon .\assets\icon_logo.ico .\capza\main.py

### Delete PYQT APP

### Download installer
url = 'https://download851.mediafire.com/asmcmx1qzgjg/ddmykrfge4nt0ep/CapZa-0.1.2.zip'
r = requests.get(url, allow_redirects=True)

open('CapZa-0.1.2.zip', 'wb').write(r.content)



### Update Versionsnummer

### Open

