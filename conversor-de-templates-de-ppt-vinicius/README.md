# Conversor de Apresentações PowerPoint - Fundação Vanzolini

Uma aplicação web para converter apresentações PowerPoint aplicando automaticamente templates da Fundação Vanzolini.

## Documentação Técnica

Consulte a documentação técnica completa de arquitetura e APIs em:
- `docs/ARQUITETURA_E_API.md`

## Características

- Interface web moderna e responsiva
- Upload de template e arquivos ZIP com apresentações
- Conversão em lote usando automação do PowerPoint
- Download automático dos arquivos convertidos
- Design com identidade visual da Fundação Vanzolini

## Requisitos do Sistema

- **Windows** (obrigatório - usa automação COM do PowerPoint)
- **Microsoft PowerPoint** instalado
- **Python 3.7+**
- **pip** (gerenciador de pacotes Python)

## Instalação

1. **Clone ou baixe o projeto**
   ```
   Extraia os arquivos em uma pasta de sua escolha
   ```

2. **Instale as dependências Python**
   ```cmd
   pip install -r requirements.txt
   ```

3. **Execute a aplicação**
   ```cmd
   python app.py
   ```
   
   Ou use o arquivo batch incluído:
   ```cmd
   executar.bat
   ```

4. **Acesse a aplicação**
   - Abra seu navegador
   - Vá para: `http://localhost:5000`

## Como Usar

1. **Prepare seus arquivos:**
   - Um arquivo template (.ppt ou .pptx) com o design da Fundação Vanzolini
   - Um arquivo ZIP contendo todas as apresentações que deseja converter

2. **Na interface web:**
   - Selecione o arquivo template
   - Selecione o arquivo ZIP com as apresentações
   - Clique em "Converter Apresentações"

3. **Aguarde a conversão:**
   - O sistema mostrará o progresso da conversão
   - Quando concluído, um link de download aparecerá

4. **Baixe o resultado:**
   - Clique no botão de download
   - Você receberá um ZIP com todas as apresentações convertidas

## Estrutura do Projeto

```
conversor-ppt/
├── app.py                 # Aplicação Flask principal
├── requirements.txt       # Dependências Python
├── executar.bat          # Script para executar no Windows
├── README.md             # Esta documentação
├── static/
│   ├── logo-vanzolini.svg # Logotipo da Fundação Vanzolini
│   ├── uploads/          # Arquivos temporários de upload
│   └── downloads/        # Arquivos convertidos para download
└── templates/
    └── index.html        # Interface web principal
```

## Funcionalidades Técnicas

- **Backend Flask** com endpoints para upload, conversão e download
- **Automação PowerPoint** usando win32com.client
- **Interface responsiva** com Tailwind CSS
- **Upload drag-and-drop** para melhor experiência do usuário
- **Limpeza automática** de arquivos temporários
- **Feedback visual** do progresso da conversão

## Solução de Problemas

### PowerPoint não encontrado
- Certifique-se de que o Microsoft PowerPoint está instalado
- Execute o Python como administrador se necessário

### Erro de permissão
- Verifique se o PowerPoint não está sendo usado por outro processo
- Execute o comando como administrador

### Porta em uso
- Se a porta 5000 estiver em uso, edite o arquivo `app.py`
- Altere a linha: `app.run(host='0.0.0.0', port=5000, debug=True)`
- Para: `app.run(host='0.0.0.0', port=5001, debug=True)`

## Suporte

Para suporte técnico ou dúvidas sobre o uso da aplicação, entre em contato com a equipe de TI da Fundação Vanzolini.

---

**Fundação Vanzolini** - Conversor de Apresentações PowerPoint  
Versão 1.0 - 2024

