# Genera index.html (un solo archivo) a partir de src/app.html y src/core.js
from pathlib import Path
app = Path('src/app.html').read_text(encoding='utf-8')
core = Path('src/core.js').read_text(encoding='utf-8')
Path('index.html').write_text(app.replace('/*CORE*/', core), encoding='utf-8')
print('index.html generado')
