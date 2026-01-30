# Conversor de Apresentações PowerPoint — Documentação Técnica

Este documento descreve a arquitetura, os componentes, os fluxos críticos, as APIs, os formatos de dados e as considerações operacionais do Conversor de Apresentações PowerPoint.

## Visão Geral

- **Objetivo**: Aplicar automaticamente um template (.ppt/.pptx) a uma coleção de apresentações PowerPoint contidas em um arquivo .zip, entregando um .zip com as apresentações convertidas.
- **Plataforma**: Windows (obrigatório), devido à automação COM do Microsoft PowerPoint.
- **Stack**: Flask (Python) + pywin32 (COM PowerPoint) + HTML/Tailwind para UI.

## Componentes

- `app.py`: Servidor Flask, rotas HTTP, orquestração de conversão, logging e persistência de progresso.
- `templates/index.html`: Interface web (upload, progresso, resultado, UX).
- `static/uploads/`: Área temporária por conversão (template, zip recebido, extração em `presentations/`).
- `static/downloads/`: Saída por conversão (arquivos convertidos e zip final).
- `data/progress/`: Estado de progresso por conversão (`conversion_YYYYMMDD_HHMMSS.json`).
- `data/logs/conversions.jsonl`: Trilhas de auditoria em formato JSONL (um evento por linha).

## Fluxo de Alto Nível

1. Usuário envia `template.pptx` e `apresentacoes.zip` via UI (`POST /upload`).
2. Backend cria `conversion_id`, persiste arquivos, escreve estado `queued` e dispara uma thread.
3. Thread:
   - Valida o ZIP (deve conter ao menos 1 `.ppt`/`.pptx`).
   - Extrai para `static/uploads/<conversion_id>/presentations/`.
   - Para cada apresentação: abre no PowerPoint via COM, aplica o template, normaliza layouts candidatos (ver Lógica), salva convertido em `static/downloads/<conversion_id>/`.
   - Gera `static/downloads/<conversion_id>_convertidos.zip`.
   - Atualiza progresso para `done` e registra evento de conclusão.
4. UI faz polling de `GET /progress/<conversion_id>` até `done` e então oferece `GET /download/<conversion_id>`.
5. Após o download, ocorre limpeza de artefatos temporários.

## Endpoints HTTP

- `GET /`
  - Retorna a UI (`index.html`).

- `GET /status`
  - Payload: `{ status: 'online', powerpoint_available: boolean }`.
  - Verifica inicialização COM e instanciamento do PowerPoint.

- `POST /upload`
  - Form-data: `template` (.ppt/.pptx), `presentations` (.zip).
  - Respostas:
    - 200 `{ success: true, conversion_id }` ao enfileirar conversão.
    - 400 em erros de validação (extensões ausentes/inválidas).
    - 500 para falhas internas.

- `GET /progress/<conversion_id>`
  - Respostas:
    - 200 com conteúdo do arquivo `data/progress/<conversion_id>.json`.
    - 404 `{ status: 'unknown' }` se não existir.

- `GET /download/<conversion_id>`
  - Fornece o arquivo `static/downloads/<conversion_id>_convertidos.zip` como anexo.
  - Agenda rotina de limpeza de diretórios e progresso ao fechar a resposta.

## Modelo de Progresso (data/progress/*.json)

Estados principais e campos observados:

```json
{
  "ts": "2025-08-20T15:44:04.123456",
  "status": "queued | processing | done | error",
  "stage": "opening | applying_template | saving | error",
  "current_file": "AULA 1.pptx",
  "converted_count": 2,
  "total_files": 5,
  "error": "Mensagem de erro (quando status=error)",
  "converted_files": ["AULA 1.pptx", "AULA 2.pptx"]
}
```

- Atualizado por: `write_progress()` e callback `progress_cb` durante `convert_presentations()`.

## Logging de Auditoria (data/logs/conversions.jsonl)

- Formato: uma linha por evento, cada linha um JSON.
- Eventos: `conversion_start`, `conversion_done`, `conversion_error`, `download`, `cleanup_error`.
- Campos comuns: `ts`, `event`, `conversion_id`, e metadados (ex.: `files`, `error`).

Exemplo (linhas):

```json
{"ts":"2025-08-20T15:44:05.001","event":"conversion_start","conversion_id":"conversion_20250820_154404","template":"MD_Template.pptx","zip":"OneDrive.zip"}
{"ts":"2025-08-20T15:45:10.321","event":"conversion_done","conversion_id":"conversion_20250820_154404","total":5,"files":["AULA 1.pptx","AULA 2.pptx", "..."]}
```

## Lógica de Conversão (COM PowerPoint)

- Inicialização: `pythoncom.CoInitialize()` e `win32com.client.Dispatch("PowerPoint.Application")`.
- Para cada apresentação:
  - Abre em modo `ReadOnly=1`, `WithWindow=0`.
  - Tira um snapshot dos índices de slides sem placeholder de título no arquivo ORIGINAL (`collect_no_title_indices`).
  - Aplica template: `pres.ApplyTemplate(template_path)`.
  - Normaliza layouts de forma conservadora: `normalize_layouts_with_sem_secao_fallback(pres, orig_no_title_idx, also_last_n=3, use_keywords=true)`:
    - Busca um layout chamado variações de "SEM_SEÇÃO" no master do template.
    - Aplica somente a slides candidatos (sem título no original) e que sejam dos últimos N ou que contenham palavras-chave de encerramento (`refer`, `crédit`, `bibliograf`, etc.).
    - Não há fallback genérico: se não existir layout "SEM_SEÇÃO", não força ajuste.
  - Salva no diretório de saída.
- Encerramento: `pp.Quit()` e `pythoncom.CoUninitialize()` em bloco `finally`.

## Regras de Validação

- `template` deve ser `.ppt` ou `.pptx`.
- `presentations` deve ser `.zip` e conter pelo menos um `.ppt`/`.pptx` (checado antes de extrair).
- Arquivos não PowerPoint dentro do zip são ignorados.

## Pastas e Convenções

- `static/uploads/<conversion_id>/`:
  - `presentations/` (extração do zip)
  - `*.pptx` template e zip originais
- `static/downloads/<conversion_id>/`: arquivos convertidos.
- `static/downloads/<conversion_id>_convertidos.zip`: pacote final.
- `data/progress/<conversion_id>.json`: estado da conversão.

## Considerações Operacionais

- Requisitos:
  - Windows com Microsoft PowerPoint (2016+ ou Microsoft 365).
  - Python 3.8+ e `pywin32` (instalado via `requirements.txt`).
- Execução em produção local (single-node): `python app.py`.
- O processo de conversão roda em thread dedicada; o Flask retorna imediatamente após `POST /upload`.
- Limpeza pós-download remove diretórios temporários e o zip final, além do arquivo de progresso.
- Observabilidade:
  - Consultar `/status` para checar disponibilidade do PowerPoint.
  - Consultar `data/logs/conversions.jsonl` para auditoria.

## Tratamento de Erros

- Erros de inicialização COM / instância PowerPoint: retornam `Falha ao iniciar PowerPoint/COM` em `convert_presentations` e registram log.
- Erros por arquivo (COM, IO): reportados via `progress_cb(stage='error', error=...)` e não interrompem os demais arquivos.
- Falhas gerais na thread: status `error` e log de `conversion_error`.

## Segurança e Limites

- Uploads aceitam apenas extensões esperadas; sanitização de nomes com `secure_filename`.
- Conversões rodam no mesmo processo; para concorrência elevada, recomenda-se isolar por processo/serviço e fila externa.
- Tamanho de arquivo: depende do servidor/host; ajustar limites por servidor reverse-proxy se necessário.

## Roadmap Técnico (sugestões)

- Modo serviço Windows com watchdog para resiliência.
- Fila externa (ex.: Redis/RQ) para múltiplas conversões simultâneas e retries.
- Métricas (Prometheus) e logs estruturados em arquivo diário rotacionado.
- Upload assinado e expiração de artefatos.
- Testes de integração com mocks de COM.

## Versões e Dependências

- Python: 3.8+
- Flask: 3.1.x
- pywin32: 306
- Werkzueg: 3.1.x

## Licença e Créditos

Aplicação desenvolvida para a Fundação Vanzolini. Direitos reservados conforme diretrizes internas.
