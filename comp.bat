del -r dist
del -r build
pyinstaller --path=d:\temp\eclipse\LibA5\A5 CdI.py -w -F -y
copy CdICfg.xlsx dist\.
cd dist\CdI
pause
