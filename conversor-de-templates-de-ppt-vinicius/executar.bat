@echo off
echo ========================================
echo  Conversor de Apresentacoes PowerPoint
echo  Fundacao Vanzolini
echo ========================================
echo.

echo Verificando dependencias...
pip install -r requirements.txt

echo.
echo Iniciando aplicacao web...
echo Acesse: http://localhost:5000
echo.
echo Pressione Ctrl+C para parar a aplicacao
echo.

python app.py

pause

