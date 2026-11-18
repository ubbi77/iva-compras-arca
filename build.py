# Une app.html + core.js:
#  - IVA_Compras_TIEMPO_ARCA.html : archivo suelto (doble clic, sin internet)
#  - web/index.html               : versión para publicar como página web (sin <html>/<head>/<body>)
import pathlib, re
d = pathlib.Path(__file__).parent
html = (d/'app.html').read_text(encoding='utf-8').replace('/*__CORE__*/', (d/'core.js').read_text(encoding='utf-8'))
(d/'IVA_Compras_TIEMPO_ARCA.html').write_text(html, encoding='utf-8')
title = re.search(r'<title>.*?</title>\s*<meta name="claude:viewport-fit-sides"[^>]*>', html, re.S).group(0)
style = re.search(r'<style>.*?</style>', html, re.S).group(0)
body = re.search(r'<body>(.*)</body>', html, re.S).group(1)
(d/'web').mkdir(exist_ok=True)
(d/'web'/'index.html').write_text(title + '\n' + style + '\n' + body.strip() + '\n', encoding='utf-8')
print('ok', len(html))
