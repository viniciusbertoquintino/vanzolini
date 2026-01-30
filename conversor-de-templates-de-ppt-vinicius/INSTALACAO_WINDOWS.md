# Guia de Instalação - Windows

## Pré-requisitos

### 1. Microsoft PowerPoint
- **OBRIGATÓRIO**: Microsoft PowerPoint deve estar instalado
- Versões suportadas: PowerPoint 2016, 2019, 2021, Microsoft 365
- O PowerPoint deve estar licenciado e funcionando normalmente

### 2. Python
- **Versão recomendada**: Python 3.8 ou superior
- Download: https://www.python.org/downloads/windows/
- **IMPORTANTE**: Durante a instalação, marque "Add Python to PATH"

### 3. Verificação dos Pré-requisitos
Abra o Prompt de Comando (cmd) e execute:
```cmd
python --version
```
Deve retornar algo como: `Python 3.x.x`

## Instalação

### Passo 1: Extrair os Arquivos
1. Extraia todos os arquivos do projeto em uma pasta de sua escolha
2. Exemplo: `C:\ConversorPPT\`

### Passo 2: Instalar Dependências
1. Abra o Prompt de Comando **como Administrador**
2. Navegue até a pasta do projeto:
   ```cmd
   cd C:\ConversorPPT
   ```
3. Instale as dependências:
   ```cmd
   pip install -r requirements.txt
   ```

### Passo 3: Executar a Aplicação
**Opção 1 - Usando o arquivo batch:**
1. Clique duas vezes em `executar.bat`

**Opção 2 - Via linha de comando:**
1. No Prompt de Comando, execute:
   ```cmd
   python app.py
   ```

### Passo 4: Acessar a Interface Web
1. Abra seu navegador (Chrome, Firefox, Edge)
2. Acesse: `http://localhost:5000`

## Uso da Aplicação

### 1. Preparar Arquivos
- **Template**: Um arquivo .ppt ou .pptx com o design da Fundação Vanzolini
- **Apresentações**: Crie um arquivo ZIP contendo todas as apresentações que deseja converter

### 2. Processo de Conversão
1. Na interface web, clique em "Selecione o Template"
2. Escolha seu arquivo de template (.ppt ou .pptx)
3. Clique em "Selecione o arquivo ZIP"
4. Escolha o ZIP com suas apresentações
5. Clique em "Converter Apresentações"
6. Aguarde o processamento
7. Clique em "Baixar Apresentações Convertidas"

## Solução de Problemas

### Erro: "PowerPoint não encontrado"
**Solução:**
1. Verifique se o PowerPoint está instalado
2. Execute o comando como Administrador
3. Reinicie o computador após instalar o PowerPoint

### Erro: "Acesso negado" ou "Permission denied"
**Solução:**
1. Execute o Prompt de Comando como Administrador
2. Feche todas as instâncias do PowerPoint antes de executar
3. Verifique se o antivírus não está bloqueando

### Erro: "Porta 5000 em uso"
**Solução:**
1. Feche outros programas que possam usar a porta 5000
2. Ou edite o arquivo `app.py` e altere a linha:
   ```python
   app.run(host='0.0.0.0', port=5000, debug=True)
   ```
   Para:
   ```python
   app.run(host='0.0.0.0', port=5001, debug=True)
   ```
3. Acesse então: `http://localhost:5001`

### Erro: "pip não é reconhecido"
**Solução:**
1. Reinstale o Python marcando "Add Python to PATH"
2. Ou adicione manualmente o Python ao PATH do Windows

### PowerPoint não responde
**Solução:**
1. Feche todas as instâncias do PowerPoint
2. Abra o Gerenciador de Tarefas (Ctrl+Shift+Esc)
3. Finalize todos os processos "POWERPNT.EXE"
4. Execute a aplicação novamente

## Comandos Úteis para Windows

### Verificar se a aplicação está rodando:
```cmd
netstat -ano | findstr :5000
```

### Verificar processos Python:
```cmd
tasklist | findstr python
```

### Matar processo específico:
```cmd
taskkill /PID [número_do_processo] /F
```

## Estrutura de Arquivos

```
ConversorPPT/
├── app.py                 # Aplicação principal
├── requirements.txt       # Dependências
├── executar.bat          # Script de execução
├── README.md             # Documentação geral
├── INSTALACAO_WINDOWS.md # Este arquivo
├── static/
│   ├── logo-vanzolini.svg # Logotipo
│   ├── uploads/          # Uploads temporários
│   └── downloads/        # Downloads gerados
└── templates/
    └── index.html        # Interface web
```

## Suporte

Para problemas técnicos:
1. Verifique se todos os pré-requisitos estão instalados
2. Execute como Administrador
3. Consulte a seção "Solução de Problemas"
4. Entre em contato com a equipe de TI da Fundação Vanzolini

---
**Fundação Vanzolini** - Conversor de Apresentações PowerPoint  
Versão 1.0 - 2024

