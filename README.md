# 📄 Renomear Certificados PDF

Renomeia automaticamente certificados PDF usando IA (Claude), com preview interativo antes de confirmar.

**Formato gerado:** `AAAA - Nome da Palestra - Emissor.pdf`

---

## ⚙️ Instalação

```bash
pip install pypdf anthropic flask
```

---

## 🔑 Configurar chave da API

Obtenha sua chave em https://console.anthropic.com/

```bash
# Linux / macOS
export ANTHROPIC_API_KEY="sk-ant-..."

# Windows (PowerShell)
$env:ANTHROPIC_API_KEY="sk-ant-..."

# Windows (CMD)
set ANTHROPIC_API_KEY=sk-ant-...
```

> Alternativamente, edite o arquivo `renomear.py` e coloque sua chave diretamente na linha:
> ```python
> ANTHROPIC_API_KEY = "sk-ant-..."
> ```

---

## ▶️ Uso

```bash
python renomear.py --pasta "/caminho/para/sua/pasta/OneDrive"
```

O script vai:
1. Ler todos os PDFs da pasta
2. Extrair campos com IA (ano, palestra, emissor)
3. Abrir o browser automaticamente com o preview
4. Aguardar sua confirmação antes de renomear qualquer arquivo

### Opções

| Argumento | Descrição | Padrão |
|-----------|-----------|--------|
| `--pasta` | Caminho da pasta com os PDFs | (obrigatório) |
| `--porta` | Porta do servidor local | `5555` |

---

## 🖥️ Interface de Preview

Na interface você pode:
- ✏️ **Editar** qualquer nome antes de confirmar
- ☑️ **Selecionar/desselecionar** arquivos individualmente
- ↺ **Resetar** para o nome sugerido pela IA
- 🚫 **Ignorar** arquivos que não deseja renomear
- ✦ **Renomear** todos os selecionados de uma vez

---

## 📁 Exemplo

Antes:
```
certificado_abc123.pdf
doc_final_v2.pdf
evento_2024.pdf
```

Depois:
```
2024 - Inteligência Artificial na Prática - FIAP.pdf
2023 - Workshop de Python - Alura.pdf
2024 - DevOps Fundamentals - AWS.pdf
```

---

## ⚠️ Observações

- O script **não apaga** arquivos, apenas renomeia
- O OneDrive sincroniza automaticamente os arquivos renomeados
- PDFs escaneados (imagem) podem ter extração menos precisa — edite manualmente no preview se necessário
- A análise consome créditos da API Anthropic (muito baixo: ~$0.01 por 10 certificados)
