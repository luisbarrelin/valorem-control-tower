# Valorem Control Tower

Dashboard operacional Corteva / Vylor. A publicação carrega automaticamente `data/report.json` ao abrir e consulta novamente a cada cinco minutos enquanto a página está visível. O botão **Sincronizar base** consulta imediatamente. Se a rede falhar, a última base válida permanece visível com aviso.

O fluxo no Power Automate usa SharePoint, Excel Online (Business) e GitHub, todos conectores padrão. O Office Script lê somente os campos utilizados da aba Report e envia um evento `sharepoint-report`; o GitHub valida, salva os dados e publica o site. A planilha original não é alterada nem publicada integralmente. Credenciais permanecem nos conectores, fora do site. O arquivo de dados e os campos exibidos pelo dashboard são públicos.

O conector GitHub limita cada evento a 64 KiB. O exportador verifica o tamanho antes de enviar e interrompe a execução acima de 62.000 bytes, sem truncar registros. Falhas e eventos antigos preservam a base anterior. A atualização depende da execução dos serviços Microsoft/GitHub e da publicação do Pages, podendo levar alguns minutos após salvar o Excel.

Importação manual continua disponível em **Informações da base** e pausa a consulta automática nesta sessão. Use **Sincronizar base** para retornar aos dados compartilhados.

Referência de design: UI UX Pro Max, Next Level Builder (MIT). Consulte THIRD_PARTY_NOTICES.md.
