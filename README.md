# Extraindo informações do XML

## 🖥️ Overview
Automatização da extração de dados dos XMLs de notas fiscais para planilha do cliente.

Esta planilha, após preenchimento, é utilizada para input das informações no sistema ERP.

### Acesse a planilha: [Extrator_modelo.xlsm](Extrator_modelo.xlsm)
### Acesse o código VBA: [codigo_vba.txt](codigo_vba.txt)

## 📈 Objetivos
1. **Eficiência e Produtividade:** não será mais preciso acessar o XML e preencher manualmente cada um. Cliente informou que são aproximadamente 300 documentos/mês.

2. **Acurácia e Redução de Erros:** preenchimento manual é passível de erros, neste novo formato conseguimos mitigar este problema.

3. **Acessibilidade e Custo**: este código em VBA permite que o cliente rode em EXCEL o que possui baixo custo e facilidade de implantação/acesso.



## 💡 Aprendizado

### 1. Entendimento da estrutura do XML
- Conhecimento dos nós do XML e a estrutura que compõe este formato de arquivo (por exemplo, ide, emit,ICMSTot, det)
- Aprendizado sobre nomenclaturas, hierarquias e namespaces de XML.

### 2. Validação de entradas e ambiente
- Código verifica se o caminho da pasta onde os arquivos XML se encontram é valido seguindo boas práticas de programação e evitando que o código falhe silenciosamente.

### 3. Uso de local-name() para evitar problemas com namespaces
- Os namespaces podem falhar ao utilizar somente SelectSingleNode("//tag"). Utilizado o local-name() para garantir robustez e compatibilidade com diferentes layouts de NF-e.

### 4. Reaproveitamento de lógica com estrutura de laços
- Eficiência ao processar múltiplos arquivos usando Dir dentro de um loop Do While.

### 5. Além da técnica as Soft Skills
- Atenção aos detalhes, paciência, lógica, pesquisa e pensamento estruturado foram imprescindíveis neste projeto.




