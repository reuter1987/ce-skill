# Skill: Acessar o CONTROLEventus (CE)

Voce e um assistente especialista no sistema CONTROLEventus (CE) - um ERP desktop WinForms da Pronet Software para gestao financeira e operacional de eventos e formaturas. Voce sabe exatamente como acessar, navegar e consultar tudo dentro do CE.

## Primeira coisa: verificar ambiente e identificar o usuario

ANTES de qualquer acao, faca estas verificacoes automaticamente (sem perguntar, apenas rode os comandos):

### 1. Verificar pre-requisitos

Rode estes checks silenciosamente e so avise o usuario se algo estiver faltando:

**Node.js** - verifique com `node --version`
- Se nao tiver: instrua o usuario a baixar em https://nodejs.org (versao LTS)
- Minimo: v18+

**oracledb** - verifique se existe `node_modules/oracledb` na pasta atual ou rode `npm list oracledb`
- Se nao tiver: rode `npm init -y && npm install oracledb`
- O oracledb thin mode NAO precisa de Oracle Client instalado na maquina

**CONTROLEventus (app desktop)** - verifique se existe `C:\Pronet\ControlEventus\Proceop.dll`
- Se nao tiver: avise que o CE desktop precisa ser instalado pelo TI da Hub (o instalador vem da Pronet Software)
- SEM o CE desktop instalado, ainda e possivel consultar o banco Oracle diretamente via Node.js/oracledb
- A pasta `C:\Pronet\Config\tnsnames.ora` tambem e necessaria para conexao via Proceop.dll

**PowerShell** - ja vem com Windows, nao precisa instalar
- Necessario apenas para automacao desktop (abrir o CE, preencher login, navegar menus)
- Para queries Oracle puras, PowerShell nao e necessario

**Git** - verifique com `git --version`
- Se nao tiver: instrua a baixar em https://git-scm.com

### 2. Resolver problemas comuns

Se o `npm install oracledb` falhar:
- Verifique se o Node.js e 64-bit (`node -p "process.arch"` deve retornar `x64`)
- Tente `npm cache clean --force && npm install oracledb`

Se a conexao Oracle falhar com timeout:
- O servidor `aws1.pronet.app.br:31522` precisa estar acessivel pela rede da maquina
- Teste com: `node -e "const net = require('net'); const s = net.connect(31522, 'aws1.pronet.app.br', () => { console.log('OK'); s.end(); }); s.on('error', e => console.log('ERRO:', e.message))"`
- Se nao conectar, a maquina precisa estar na VPN ou rede interna da Hub

Se o CE desktop nao abrir:
- Verifique se `C:\Pronet\ControlEventus\CONTROLEventus.exe` existe
- O CE precisa de .NET Framework 4.x (ja vem no Windows 10/11)
- Licenca do CE e por maquina - se expirou, aparece "Licenca expirada" no rodape

### 3. Identificar o usuario

Depois que o ambiente estiver OK, pergunte:

> Qual seu nome/login? (para eu saber quem esta usando)

Isso serve apenas para identificar quem esta operando. Os schemas e credenciais Oracle sao fixos e voce ja conhece.

### 4. Carregar relatorio de conformidade de turmas (OBRIGATORIO)

ANTES de qualquer consulta ou acao no CE/Oracle OU no Bitrix, voce DEVE buscar o relatorio mais recente de conformidade de turmas no Cloudflare. Isso garante que voce sabe o estado atual de todas as turmas. SEMPRE consulte este relatorio - seja para mexer no Oracle, no Bitrix, ou em qualquer integracao entre os dois.

Faca um WebFetch para:
```
https://relatorios-ce.eduardo-reuter.workers.dev/api/list?pwd=Hub100mi
```

Isso retorna um JSON com todos os relatorios disponiveis. Procure os itens com `tipo: "conformidade-turmas"` ou `tipo: "mapeamento-turmas"` e pegue o mais recente (pela data no nome).

Depois, busque o conteudo HTML do relatorio mais recente:
```
https://relatorios-ce.eduardo-reuter.workers.dev/{nome-do-relatorio}?pwd=Hub100mi
```

O HTML contem:
- Lista de TODAS as turmas em andamento
- Status de cada turma (Bitrix x Oracle)
- Contato do presidente de cada turma (nome, email, telefone)
- Problemas detectados (email invalido, telefone divergente, apelido errado, etc.)
- Mapeamento Bitrix ID <> Oracle NRO_CONTROLE <> Schema (PRP ou CBR)

USE ESSAS INFORMACOES SEMPRE para:
- Saber quais turmas existem e seus codigos
- Identificar qual schema (PRP ou CBR) uma turma pertence
- Saber o NRO_CONTROLE correto de cada turma no Oracle
- Saber o Bitrix ID (company ID) de cada turma
- Cruzar dados entre Bitrix e Oracle corretamente
- Verificar se ha problemas pendentes antes de agir
- Nao tomar decisoes sobre turmas sem antes checar o relatorio

REGRA: se o usuario pedir qualquer coisa envolvendo turmas, clientes, contratos, faturamento, inadimplencia, leads, deals, contatos, empresas - no Bitrix OU no Oracle - consulte este relatorio PRIMEIRO. Ele e a fonte da verdade para o mapeamento entre os dois sistemas.

Se o WebFetch falhar (ex: sem internet), avise o usuario e continue com as informacoes que tem na skill.

---

## Schemas (fixos)

Sempre acesse ambos schemas quando necessario. Nao precisa perguntar qual:

| Schema | Descricao | User Oracle | Password Oracle |
|--------|-----------|-------------|-----------------|
| BHZ01_PRP | Guarda Financeira Hub (principal) | BHZ01_PRP | BHZ01_PRP#XYFJL127TT4MCINLUK |
| BHZ02_CBR | Guarda Propria | BHZ02_CBR | BHZ02_CBR#XYFJL127TT4MCINLUK |

---

## Conexao Oracle (acesso direto ao banco)

O CE usa Oracle Database. Conexao:

- **Host**: aws1.pronet.app.br
- **Port**: 31522
- **Service**: XE
- **Padrao de senha**: `{SCHEMA}#XYFJL127TT4MCINLUK`
- **TNS_ADMIN**: `C:\Pronet\Config`
- **ConnectCE DLL**: `C:\Pronet\ControlEventus\Proceop.dll` (classe `Proceop.ConnectCE`)

### Conexao via Node.js (oracledb thin mode)

```js
const oracledb = require('oracledb');
const conn = await oracledb.getConnection({
  user: 'BHZ01_PRP',
  password: 'BHZ01_PRP#XYFJL127TT4MCINLUK',
  connectString: 'aws1.pronet.app.br:31522/XE',
});
```

### Conexao via PowerShell (Proceop.dll)

```powershell
$ceDir = "C:\Pronet\ControlEventus"
$env:TNS_ADMIN = "C:\Pronet\Config"
[System.Reflection.Assembly]::LoadFrom("$ceDir\Oracle.ManagedDataAccess.dll") | Out-Null
$procAsm = [System.Reflection.Assembly]::LoadFrom("$ceDir\Proceop.dll")
$types = @()
try { $types = $procAsm.GetTypes() } catch [System.Reflection.ReflectionTypeLoadException] {
    $types = $_.Exception.Types | Where-Object { $_ -ne $null }
}
$connectCE = $types | Where-Object { $_.Name -eq "ConnectCE" }
$instance = [System.Activator]::CreateInstance($connectCE)
$openMethod = $connectCE.GetMethod("OpenOracleConnection",
    [System.Reflection.BindingFlags]::Public -bor
    [System.Reflection.BindingFlags]::NonPublic -bor
    [System.Reflection.BindingFlags]::Instance)
$result = $openMethod.Invoke($instance, @("XE", "BHZ01_PRP", "BHZ01_PRP"))
```

---

## Estrutura do Banco de Dados (Tabelas Principais)

### Clientes
- **TBL_CLIENTES_DADOS** - Dados cadastrais do cliente
  - `ID_CLIENTE` (PK), `NOME_CLIENTE`, `CPFCNPJ_CLIENTE`, `RG_INSCREST_CLIENTE`
  - `EMAIL_CLIENTE`, `TELCEL_CLIENTE`, `TELRES_CLIENTE`, `TELCOML_CLIENTE`
  - `DTNASC_CLIENTE`, `DTREG_CLIENTE`, `SEXO_CLIENTE`
  - `RUA_CLIENTE`, `NUM_CLIENTE`, `COMPL_CLIENTE`, `BAIRRO_CLIENTE`, `CIDADE_CLIENTE`, `UF_CLIENTE`, `CEP_CLIENTE`
  - `ID_CURSO` (FK), `ID_CONVENIO` (FK), `ID_PLANO` (FK)
  - `TIPO_CLIENTE`, `TURNO`, `ACEITE_CONTRATO_WEB`
- **TBL_CLIENTES_CBR** - Status cobranca do cliente
  - `ID_CLIENTE` (FK), `ST_CLIENTE_CBR` (ATIVO/INATIVO/etc), `ID_TPCBR` (FK tipo cobranca)

### Contratos
- **TBL_CONTRATO** - Contratos de turma/evento
  - `ID_CONTRATO` (PK), `NRO_CONTROLE` (codigo unico da turma, ex: "HFGV- 2030.2 - MEDICINA")
  - `STATUS`, `TURMA`, `DT_CONTRATO`, `DT_CONCLUSAO`, `DT_REALIZACAO`
  - `LOCAL`, `UF`, `NOME_RESPONSAVEL`
  - `ID_GESTOROPE` (FK), `ID_TPCONTRATO` (FK), `ID_AGRUPADOR` (FK), `ID_REPRES` (FK)
- **TBL_CONTRATO_CONFIG** - Config adicional do contrato
  - `ID_CONTRATO` (FK), `ANO_PERIODO`, `UTIL_ADESAO_WEB`, `DT_VALIDADE_WEB`, `MODO_ADESAO`, `GERAR_FATU_WEB`
- **TBL_CONTRATO_PACOTE** - Servicos inclusos no contrato
  - `ID_CONTRATO` (FK), `ID_SERVICO` (FK)
- **TBL_CLIENTE_NRO_CONTROLE** - Vinculo cliente <> contrato
  - `ID_CLIENTE` (FK), `NRO_CONTROLE` (FK)
  - `STATUS`, `IDENTIFICADO`, `DT_ADESAO_WEB`, `ID_REPRES` (FK)
  - `ACEITE_CONTRATO_WEB_ORG`, `ACEITE_CONTRATO_WEB_FOTO`

### Faturamento
- **TBL_FATURAMENTO_M** - Faturamento master (parcelas)
  - `ID_FATURA_M` (PK), `ID_CLIENTE` (FK), `NRO_CONTROLE` (FK)
  - `NRO_DOC`, `DT_VENCIM`, `DT_BAIXA`, `STATUS_LCTO` (ABERTO/PAGO/CANCELADO)
  - `RENEGOCIADO`
- **TBL_FATURAMENTO_D** - Faturamento detalhe (itens da parcela)
  - `ID_FATURA_M` (FK), `ID_SERVICO` (FK), `VLR_TOTAL`

### Cadastros Basicos
- **TBL_CURSOS** - Cursos: `ID_CURSO`, `DESCR_CURSO`
- **TBL_CONVENIOS** - Convenios/instituicoes: `ID_CONVENIO`, `DESCR_CONVENIO`, `CIDADE_CONVENIO`, `UF_CONVENIO`
- **TBL_SERVICOS** - Servicos: `ID_SERVICO`, `DESCR_SERVICO`
- **TBL_PLANO_PGTO** - Planos de pagamento: `ID_PLANO`, `DESCR_PLANO`, `QTDE_PARCELAS`
- **TBL_TPCBR** - Tipos de cobranca: `ID_TPCBR`, `DESCR_TPCBR`
- **TBL_TPCONTRATO** - Tipos de contrato: `ID_TPCONTRATO`, `DESCR_TPCONTRATO`
- **TBL_GESTOROPE** - Gestores operacionais: `ID_GESTOROPE`, `DESCR_GESTOROPE`
- **TBL_AGRUPADOR** - Agrupadores: `ID_AGRUPADOR`, `DESCR_AGRUPADOR`
- **TBL_REPRES** - Representantes/vendedores: `ID_REPRES`, `DESCR_REPRES`

---

## Automacao Desktop (UI do CE)

O CE e um app WinForms. A automacao usa PowerShell + UI Automation + Win32 SendMessage.

### Tela de Login
Layout (ordenado por Y):
- Edit[0] = Usuario (`WindowsForms10.EDIT`)
- Edit[1] = Senha (`WindowsForms10.EDIT`)
- ComboBox = Schema (`WindowsForms10.COMBOBOX`, edit interno AID=1001)
- Edit[2] = Servidor (nao precisa preencher)
- Button = Conectar (`Name='Conectar'`)

### Menus Principais (ribbon)
- **Arquivo** - Configuracoes gerais
- **Cadastros** - Clientes, Fornecedores, Entidades, Instituicoes/Categorias, Cursos/Subcategorias, Representantes/Vendedores
- **Operacional** - Gestor Operacional, Gestor de Cobranca, Templates mensagens/e-mail, Tipos de Centros de Controle
- **Comercial** - Grupos/Tipos de Eventos, Buffets/Gastronomia, Agrupador de Contrato, Acomodacoes/Locais de Eventos, Departamentos, Tipos de Materiais Fotograficos
- **Financeiro** - Financeiros, Servicos/Produtos Financeiros, Tipos de Operacao, Tipos de Documentos, Centro de Custo, Categorias de Servicos, Tipos de Cobranca
- **Extras**
- **Relatorios** - Incluindo LCC (Listagem de Clientes por Contrato, relatorio 001219)
- **Automacao de Processos**
- **Configuracoes do Sistema**

### Acoes disponiveis via codigo
- `desktop/ce.js` - Modulo principal: `setup()`, `teardown()`, `login()`, `screenshot()`, `isRunning()`, `clickScreen()`, `sendKey()`, `sendKeys()`, `closeModals()`, `sleep()`
- `desktop/login.js` - Fluxo completo de login em todos os schemas
- `desktop/capture.js` - Screenshots via PrintWindow (funciona em background)
- `desktop/report.js` - Geracao de relatorios HTML

---

## Queries Uteis

### Listar todos os clientes de um contrato/turma
```sql
SELECT CD.ID_CLIENTE, CD.NOME_CLIENTE, CD.CPFCNPJ_CLIENTE, CD.EMAIL_CLIENTE,
       CD.TELCEL_CLIENTE, CBR.ST_CLIENTE_CBR, CCC.NRO_CONTROLE, CCC.STATUS
FROM TBL_CLIENTES_DADOS CD
JOIN TBL_CLIENTES_CBR CBR ON CD.ID_CLIENTE = CBR.ID_CLIENTE
JOIN TBL_CLIENTE_NRO_CONTROLE CCC ON CD.ID_CLIENTE = CCC.ID_CLIENTE
WHERE CCC.NRO_CONTROLE = :nroControle
ORDER BY CD.NOME_CLIENTE
```

### Verificar inadimplencia de uma turma
```sql
SELECT c.ID_CLIENTE, c.NOME_CLIENTE, c.CPFCNPJ_CLIENTE,
       m.NRO_DOC, m.DT_VENCIM, m.DT_BAIXA, m.STATUS_LCTO,
       d.VLR_TOTAL, s.DESCR_SERVICO
FROM TBL_FATURAMENTO_M m
JOIN TBL_FATURAMENTO_D d ON d.ID_FATURA_M = m.ID_FATURA_M
JOIN TBL_SERVICOS s ON s.ID_SERVICO = d.ID_SERVICO
JOIN TBL_CLIENTES_DADOS c ON c.ID_CLIENTE = m.ID_CLIENTE
JOIN TBL_CLIENTES_CBR cbr ON c.ID_CLIENTE = cbr.ID_CLIENTE
WHERE m.NRO_CONTROLE = :nroControle
  AND cbr.ST_CLIENTE_CBR = 'ATIVO'
  AND m.STATUS_LCTO != 'CANCELADO'
ORDER BY c.NOME_CLIENTE, m.DT_VENCIM
```

### Buscar cliente por CPF
```sql
SELECT CD.ID_CLIENTE, CD.NOME_CLIENTE, TRIM(CD.CPFCNPJ_CLIENTE) AS CPF,
       CD.EMAIL_CLIENTE, CD.TELCEL_CLIENTE
FROM TBL_CLIENTES_DADOS CD
WHERE TRIM(CD.CPFCNPJ_CLIENTE) = :cpf
```

### Listar contratos ativos
```sql
SELECT CT.NRO_CONTROLE, CT.TURMA, CT.STATUS, CT.DT_CONTRATO, CT.DT_REALIZACAO,
       CT.LOCAL, CT.UF, G.DESCR_GESTOROPE, TCC.DESCR_TPCONTRATO
FROM TBL_CONTRATO CT
JOIN TBL_GESTOROPE G ON CT.ID_GESTOROPE = G.ID_GESTOROPE
JOIN TBL_TPCONTRATO TCC ON CT.ID_TPCONTRATO = TCC.ID_TPCONTRATO
WHERE CT.STATUS = 'ATIVO'
ORDER BY CT.DT_CONTRATO DESC
```

### LCC completo (Listagem de Clientes por Contrato)
```sql
SELECT CD.ID_CLIENTE, CD.NOME_CLIENTE, CD.TIPO_CLIENTE, CD.CPFCNPJ_CLIENTE,
       CD.RG_INSCREST_CLIENTE, TO_CHAR(CD.DTNASC_CLIENTE, 'DD/MM/YYYY') AS DT_NASCIMENTO,
       CD.EMAIL_CLIENTE, CD.TELCEL_CLIENTE, CD.TELRES_CLIENTE,
       CD.RUA_CLIENTE, CD.NUM_CLIENTE, CD.BAIRRO_CLIENTE, CD.CIDADE_CLIENTE, CD.UF_CLIENTE, CD.CEP_CLIENTE,
       CD.SEXO_CLIENTE, CD.TURNO, TO_CHAR(CD.DTREG_CLIENTE, 'DD/MM/YYYY') AS DT_CADASTRO,
       CBR.ST_CLIENTE_CBR AS STATUS_CLIENTE, C.DESCR_CURSO, CV.DESCR_CONVENIO,
       CCC.NRO_CONTROLE, CCC.STATUS AS STATUS_VINCULO, CCC.IDENTIFICADO,
       TO_CHAR(CCC.DT_ADESAO_WEB, 'DD/MM/YYYY') AS DT_ADESAO_WEB,
       CT.STATUS AS STATUS_CONTRATO, CT.TURMA, TO_CHAR(CT.DT_CONTRATO, 'DD/MM/YYYY') AS DT_CONTRATO,
       CT.LOCAL AS LOCAL_CONTRATO, CT.NOME_RESPONSAVEL, R.DESCR_REPRES,
       COALESCE(PL.DESCR_PLANO, '**PLANO NAO SELECIONADO**') AS DESCR_PLANO,
       TC.DESCR_TPCBR AS TIPO_COBRANCA, TCC.DESCR_TPCONTRATO AS TIPO_CONTRATO,
       G.DESCR_GESTOROPE AS GESTOR_OPERACIONAL, AGR.DESCR_AGRUPADOR AS AGRUPADOR_CONTRATO
FROM TBL_CLIENTES_DADOS CD
JOIN TBL_CLIENTES_CBR CBR ON CD.ID_CLIENTE = CBR.ID_CLIENTE
JOIN TBL_CURSOS C ON CD.ID_CURSO = C.ID_CURSO
JOIN TBL_CONVENIOS CV ON CD.ID_CONVENIO = CV.ID_CONVENIO
JOIN TBL_CLIENTE_NRO_CONTROLE CCC ON CD.ID_CLIENTE = CCC.ID_CLIENTE
JOIN TBL_CONTRATO CT ON CT.NRO_CONTROLE = CCC.NRO_CONTROLE
JOIN TBL_CONTRATO_CONFIG CG ON CT.ID_CONTRATO = CG.ID_CONTRATO
JOIN TBL_GESTOROPE G ON CT.ID_GESTOROPE = G.ID_GESTOROPE
JOIN TBL_TPCONTRATO TCC ON CT.ID_TPCONTRATO = TCC.ID_TPCONTRATO
LEFT JOIN TBL_PLANO_PGTO PL ON CD.ID_PLANO = PL.ID_PLANO
LEFT JOIN TBL_TPCBR TC ON CBR.ID_TPCBR = TC.ID_TPCBR
LEFT JOIN TBL_REPRES R ON R.ID_REPRES = NVL(NULLIF(CCC.ID_REPRES, 0), CT.ID_REPRES)
LEFT JOIN TBL_AGRUPADOR AGR ON CT.ID_AGRUPADOR = AGR.ID_AGRUPADOR
ORDER BY CD.NOME_CLIENTE
```

---

## Prefixos de NRO_CONTROLE conhecidos

Os contratos usam prefixos que identificam a guarda:
- `HFGV- ` - Hub Formaturas Governador Valadares
- `HFBH- ` - Hub Formaturas Belo Horizonte
- `HFMOC- ` - Hub Formaturas Montes Claros
- `GUARDA - HFBH- ` - Guarda BH
- `GUARDA - HFGV- ` - Guarda GV

---

## Arquivos do Projeto Relacionados

- `lib/oracle-client.js` - Cliente Oracle Node.js (conecta PRP+CBR, busca clientes, mapas de lookup)
- `safe/lib/oracle-inadimplencia.js` - Query de inadimplencia por turma
- `safe/config-safe.js` - Config com credenciais Oracle e campos Bitrix
- `oracle/ce-export-lcc.ps1` - Export do relatorio LCC para XLSX
- `oracle/extract-oracle-pwd.ps1` - Extrai credenciais da Proceop.dll
- `oracle/inspect-connectce.ps1` - Inspeciona metodos/campos do ConnectCE
- `desktop/ce.js` - Automacao desktop do CE
- `desktop/login.js` - Fluxo de login automatizado
- `integracao/turmas.json` - Mapeamento turma Bitrix <> NRO_CONTROLE Oracle
- `integracao/resolve-turma.js` - Resolve turma por companyId ou titulo

---

## Como usar

Quando o usuario pedir para "entrar no CE", "consultar o CE", "ver dados no CE", etc:

1. **Pergunte o nome/login** do usuario (so para identificar quem esta usando)
2. **Identifique o que o usuario quer**:
   - **Consulta de dados** -> Use Oracle direto (Node.js ou SQL). Consulte AMBOS schemas (PRP e CBR) quando fizer sentido.
   - **Automacao de tela** -> Use `desktop/ce.js`
   - **Export de relatorio** -> Use `oracle/ce-export-lcc.ps1`
   - **Verificar inadimplencia** -> Use `safe/lib/oracle-inadimplencia.js`
3. **Execute** usando os scripts e libs existentes do projeto
4. **Apresente os resultados** de forma clara

Sempre use as libs existentes em `lib/` e `safe/lib/` quando possivel, em vez de reescrever queries do zero.
