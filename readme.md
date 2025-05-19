## Como construir o projeto

Para construir o seu projeto em um executável `.exe`, siga os seguintes passos:

**Crie o executável:**

Após garantir que o PyInstaller está instalado, navegue até o diretório onde o arquivo `main.py` está localizado. Em seguida, execute o seguinte comando no terminal:

<!-- pyinstaller --name "saida-facil" --onefile --windowed main.py  -->

pyinstaller --onefile --name "saida facil" --add-data "src/template_base/*;template_base" --windowed  main.py


