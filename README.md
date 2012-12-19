Sobre
=====

Script criado para monitorar uma pasta e copiar os arquivos novos para outra pasta. Como data de comparação estamos utilizando a data de modificação dos arquivos.

É possível adicionar o arquivo scan_files.bat no agendador de tarefas para execução recorrente.

Configuração
============

* Para preparar o script é necessário setar a data do arquivo newest_file.txt para 1/1/1970 00:00:00
* Dentro do arquivo scan_files.vbs setar o caminho de origem dos arquivos a serem scaneados na constante origin_path
* Dentro do arquivo scan_files.vbs setar o caminho de destino dos arquivos scaneados na constante destiny_path

OBS: Verificar se o usuário tem permissão de criação de arquivos no path onde se encontra.

