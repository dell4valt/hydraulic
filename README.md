# Hydraulic
![GitHub last commit](https://img.shields.io/github/last-commit/dell4valt/hydraulic)

Гидравлические расчеты уровней и скоростей воды в русле реки.


## Примеры
Пример запуска расчётов.
Входные данные в Excel файле _example/example_profile.xlsx_, результирующий отчёт записать в файл _result/test.docx'_.

```python
#!/usr/bin/env python3
from hydraulic import profile

in_filename = 'example/example_profile.xlsx'
out_filename = 'result/test.docx'

profile.xls_calculate_hydraulic(in_filename, out_filename)
```

