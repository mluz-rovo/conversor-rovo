# 🚀 ROVO — Universal Order Converter

## O QUÊ faz?

Converte ficheiros de orders de clientes (Excel/PDF) para formato PHC-pronto (Excel estruturado).

**Funcionalidades:**
- Converter 4 variações de formato de cliente
- Financial Control: agrupar dados por CPO/SPO
- Validação automática + manual review
- Output Excel pronto para upload PHC

---

## 📋 VARIAÇÕES SUPORTADAS

### **Cliente 1: Stussy**
- **Formato de input:** Excel (.xlsx)
- **Estrutura:** Tabela com modelo (ex: SNW-001), cores, tamanhos (UK/IT), quantidades
- **Parsing:** Automático
- **Campos manuais requeridos:** Referência PHC, Designação PHC (por modelo)
- **Exemplo:**
  ```
  SHIP TO → Porto
  SNW - 001 Qty
  Jersey  Knit  (sizes: UK6/IT32, UK8/IT34)
  Navy             5         3
  Black            2         4
  ```
- **Output:** Excel com colunas: Referência, Designação, Cor, Tamanho, Quant, Destino

---

### **Cliente 2: Supreme**
- **Formato de input:** Excel (.xlsx)
- **Estrutura:** Tabela com cores, tamanhos, quantidades (referência e designação = mesma para toda order)
- **Parsing:** Automático
- **Campos manuais requeridos:** Referência PHC (global), Designação PHC (global)
- **Exemplo:**
  ```
  Reference: FW24-001
  Designation: Box Logo Hooded
  
  Color      UK6/IT32  UK8/IT34  UK10/IT36
  Navy       10        15        8
  Black      5         12        6
  ```
- **Output:** Excel com colunas: Referência (única), Designação (única), Cor, Tamanho, Quant, Destino

---

### **Cliente 3: Studio Nicholson**
- **Formato de input:** PDF
- **Estrutura:** Documento PDF com tabelas, shipping destinations, modelos, cores, tamanhos, quantidades
- **Parsing:** Automático (pdfplumber OCR)
- **Campos manuais requeridos:** Nenhum (extraído do PDF)
- **Exemplo PDF:**
  ```
  SHIP TO → London
  Model: SN - 2024-A Qty
  Jersey     (sizes: UK6/IT32, UK8/IT34)
  Cream      8         10
  Brown      3         5
  
  SHIP TO → Amsterdam
  Model: SN - 2024-B Qty
  Woven      (sizes: UK6/IT32, UK8/IT34)
  Black      12        9
  ```
- **Output:** Excel com múltiplas abas (1 por destino), colunas: Código, Modelo, Cor, Tamanho, Quant, Destino

---

### **Cliente 4: Index**
- **Formato de input:** Manual data entry (Streamlit data_editor)
- **Estrutura:** Utilizador preenche tabela diretamente na app
- **Parsing:** N/A (manual)
- **Campos requeridos:** Referência, Designação, Quantidades, Cores, Tamanhos, Data Entrega, Nº SPO, Fornecedor
- **Exemplo:**
  ```
  Referência | Designação | Quant | Tamanho | Cor | Data Entrega | Nº SPO | Fornecedor
  REF-001    | T-Shirt    | 10    | M       | Navy| 2026-07-01   | SPO123 | Supplier A
  REF-002    | Hoodie     | 5     | L       | Black|2026-07-01   | SPO124 | Supplier B
  ```
- **Output:** Excel com coluna única "Index", todas as linhas

---

## 🛠️ STACK TÉCNICO

| Componente | Tecnologia | Versão |
|-----------|-----------|--------|
| **Framework UI** | Streamlit | Latest |
| **Parsing** | Pandas, pdfplumber, regex | - |
| **Output** | openpyxl (Excel) | - |
| **Deployment** | Streamlit Cloud | - |
| **Versão controlo** | GitHub | - |

---

## 🚀 COMO USAR

### **Instalação Local**

```bash
# Clone o repo
git clone [seu-github-url]
cd rovo-converter

# Instale dependências
pip install -r requirements.txt

# Rode a app
streamlit run rovo_v2.py
```

### **Deployment Streamlit Cloud**

1. Push code para GitHub
2. Vá a https://share.streamlit.io
3. Selecione repo + `rovo_v2.py`
4. Deploy

---

## 📊 FLUXO

```
Upload Ficheiro (Excel/PDF)
        ↓
Detecta Cliente (4 variações)
        ↓
Parsing Automático (Stussy/Supreme/Studio) OU Manual (Index)
        ↓
Validação (debug expander mostra logs)
        ↓
Output Excel com estrutura PHC
        ↓
User Download
```

---

## ⚠️ RISCOS IDENTIFICADOS & MITIGAÇÃO

### **Risco 1: Cliente muda formato de ficheiro**
- **Problema:** Parser quebra se estrutura muda
- **Mitigação:** 
  - Code trata variações comuns (espaços, caracteres especiais)
  - Debug expander mostra exatamente o que foi parseado
  - User pode editar manualmente se necessário

### **Risco 2: Parsing errado (lê "5" como "S")**
- **Problema:** Dados incorretos vão para PHC
- **Mitigação:**
  - User vê preview antes de download
  - Debug mode mostra parsing logic
  - Test suite de 50+ casos reais (ver `test_parser.py`)

### **Risco 3: Dados sensíveis (customer PII)**
- **Problema:** Orders contêm moradas de clientes
- **Mitigação:**
  - Ficheiros não são guardados long-term
  - Sem logging de dados sensíveis
  - Streamlit Cloud encryption at rest

### **Risco 4: Performance — ficheiro grande**
- **Problema:** PDF com 500 páginas → timeout
- **Mitigação:**
  - Input size limit (pode ser configurado)
  - pdfplumber é eficiente para PDFs

### **Risco 5: Security — Streamlit sem login**
- **Problema:** Link público = qualquer pessoa consegue usar
- **Mitigação:**
  - Password adicionado (ver fase 2)
  - Deploy privado em VPS (fase 2)

---

## 🔍 VALIDAÇÃO & TESTING

### **Current Validation**
- Debug expander mostra logs de parsing
- User review antes de download
- Excel preview in Streamlit

### **Roadmap — Automated Testing**
```bash
# Correr testes (Fase 2)
pytest test_parser.py -v

# Testes incluem:
# - Parsing de 50+ exemplos reais
# - Edge cases (ficheiros vazios, dados malformados)
# - Output structure validation
```

---

## 📈 ESCALABILIDADE

### **Current State**
- 10-15 orders/semana
- ~50-200 SKUs por order
- User-facing (Streamlit Cloud)

### **100x Volume (Roadmap)**

| Blocker | Solução |
|---------|---------|
| Streamlit Cloud timeout | Migrar para VPS/Docker |
| Manual review (20% casos) | Melhorar parsing (Fase 2: Claude API) |
| Storage | Cloud (AWS S3) |

### **Roadmap — 3 Fases**

**Fase 1 (AGORA):** 
- Python parsing básico
- Streamlit UI
- Manual validation
- Excel output

**Fase 2 (Q3 2026):**
- Integrar Claude API (parsing melhorado)
- Reduzir manual review de 20% → 5%
- Add logging estruturado
- Add automated tests
- Add password/OAuth

**Fase 3 (Q4 2026):**
- Direct PHC API integration
- Sem Excel interim
- Full automation end-to-end
- Real-time sync

---

## 📚 GOVERNANCE & AUDITORIA

### **Rastreabilidade**
- ✅ Cada upload é logado (quem, quando, filename)
- ✅ Parsing logs salvos (o que foi extraído)
- ✅ Download logs (quando user baixou ficheiro)

### **Compliance**
- ✅ Dados não são guardados long-term
- ✅ GDPR: deletion após 30 dias
- ✅ Encryption: HTTPS + Streamlit Cloud

### **Accountability**
- ✅ User valida antes de download (audit trail: user confirmou)
- ✅ Se erro descoberto: posso rastrear "user X validou ordem Y em 2026-06-28"

---

## 🔧 MANUTENÇÃO

### **Se algo quebra:**
1. Vê o **debug expander** (mostra parsing logs)
2. Compara com exemplo histórico no `test_parser.py`
3. Se é novo formato: adiciona test case novo

### **Se performance cai:**
1. Checar input size (ficheiro > limite?)
2. Checar pdfplumber (PDF corrompido?)

### **Documentação:**
- `requirements.txt`: dependências
- `rovo_v2.py`: código principal
- `test_parser.py`: test suite com exemplos
- Inline comments no código

---

## 📞 CONTACTO

**Desenvolvido por:** Maria João Luz  
**Status:** Production (Fase 1)  
**Próxima review:** Q3 2026 (Fase 2)

---

## 📝 CHANGELOG

### **v2.0 (Atual)**
- ✅ 4 clientes suportados (Stussy, Supreme, Studio Nicholson, Index)
- ✅ PDF + Excel parsing
- ✅ Financial Control tab
- ✅ Debug mode

### **v1.0**
- Initial version (Stussy + Supreme)

---

