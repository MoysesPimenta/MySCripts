1  Planilha‑modelo
Aba	Colunas (linha 1)	Observações
Machines	MachineID · Serial · Model · ClientID · StorageRate (R$/mês) · DateIn · DateOut · Status (stored / service / returned) · LastBilledThrough	Rate pode variar por máquina
Services	ServiceID · MachineID · ClientID · ServiceDate · ServiceType · Description · UnitPrice · QtyHours · TotalPrice · Billed (“YES/NO”)	addService() calcula Unit/Total e marca para fatura
BillingConfig	Item · UnitPrice · TaxRate	Tabela de preços‑referência
Clients	ClientID · Name · CNPJ · Address · Email · Export (YES/NO)	Export=YES → NFSe de exportação
Invoices	InvoiceID · ClientID · PeriodStart · PeriodEnd · IssueDate · DueDate · Total · Paid (YES/NO) · PaymentDate · Overdue (fórmula ou script) · PDFLink · NFSeNumber · NFSeType · LineItemsJSON	DueDate = IssueDate + 15 dias
Dashboards	(vazia – o script gera gráficos)	

Dica: Rode ensureSheets() (menú Automation → Setup) para criar todas as abas e cabeçalhos automaticamente.

2  Fluxo automático
Registro de serviço

Técnico lança linha na aba Services.

Trigger: onEdit → addService(e) preenche preços e marca “NO” em Billed.

Faturamento mensal

Trigger: Time‑driven → Monthly (todo dia 1 às 02 h) → runMonthlyBilling()

O script:

Calcula storage do mês anterior (pro rata).

Puxa todos os serviços Billed=NO.

Cria 1 fatura por cliente, define DueDate = +15 d.

Gera PDF a partir do template, grava link, envia e‑mail.

Atualiza LastBilledThrough em Machines e marca serviços como Billed=YES.

Status de atraso

Trigger: Time‑driven → Daily → updateOverdueStatus() marca Overdue=YES quando passou do vencimento e Paid=NO.

Emissão de NFSe

Você seleciona a linha na aba Invoices → menú Automation → “Issue NFSe”.

Diálogo pede CNPJ ou EXPORT; o script chama issueNFSe() na API PlugNotas/eNotas (token em Script Properties), grava número e tipo.

Dashboards

Trigger: manual (botão) ou semanal → refreshDashboards() recria gráficos “Receita Mensal”, “Custos”, “Lucro”, tempos médios, etc.

3  Triggers a configurar
Função	Tipo	Sugestão
onOpen	Simple	(já automático)
addService	onEdit (aba Services)	Instalable específico de planilha
runMonthlyBilling	Time‑driven → Monthly	Dia 1 às 02:00
updateOverdueStatus	Time‑driven → Daily	01:00
refreshDashboards	Time‑driven → Weekly (seg 03:00)	opcional

4  Dashboards prontos
O script cria/atualiza a aba Dashboards com:

Receita Mensal (colunas)

Custos Mensais (colunas)

Lucro Mensal (linha sobreposta)

Visão Caixa × Competência (colunas empilhadas)

Tempo médio de storage (scorecard)

Top 5 máquinas por faturamento (pizza)

Serviços/máquina (média) (barra)

Tempo médio de recebimento (gauge)

Estoque atual vs em serviço (barra empilhada)

Se quiser refinamentos visuais, basta editar a aba; o script respeita IDs de gráfico ou recria quando não existir.

5  NFSe São Paulo
API terceirizada: o script usa exemplo PlugNotas (https://api.plugnotas.com.br/nfse).

Configure NFSE_TOKEN em Script Properties.

Campos chave enviados: cidade 3550308, CNPJ / Export, aliquota ISS, itemListaServico, descrição consolidada.

Ajuste parâmetros (aliquota, IM, itemListaServico) no topo de issueNFSe().

Próximo passo: abra o arquivo Invoice Automation ao lado, ajuste IDs (template Doc, CNPJ, token), depois Configure → Triggers conforme tabela. Qualquer dúvida, estou por aqui!


--------------------------------