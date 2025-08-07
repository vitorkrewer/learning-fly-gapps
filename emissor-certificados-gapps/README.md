# Emissor e Gerenciador de Certificados com Google Apps Script

![Status](https://img.shields.io/badge/status-funcional-brightgreen)
![Licença](https://img.shields.io/badge/licen%C3%A7a-MIT-blue)

Um sistema completo e de baixo custo para automatizar a criação, emissão e gestão de certificados digitais. A ferramenta utiliza o poder do ecossistema Google, usando o Google Sheets como banco de dados, Google Docs como templates e Google Drive para armazenamento.

## ✨ Funcionalidades

- **Emissão em Lote**: Crie centenas de certificados personalizados a partir de uma lista em uma Planilha Google.
- **Interface Gráfica**: Menus e janelas interativas dentro da Planilha Google para facilitar o uso.
- **Gestão de Instituições**: Cadastre instituições parceiras para agilizar o preenchimento e manter a consistência.
- **Templates Dinâmicos**: Use tags como `{{NOME_PARTICIPANTE}}` e `{{NOME_EVENTO}}` em um template do Google Docs.
- **Suporte a Campos Personalizados**: Inclui campos opcionais como `{{CPF}}` e `{{Data de Emissão}}`.
- **Código de Verificação Único**: Cada certificado recebe um código alfanumérico único para validação.
- **Nomenclatura Semântica**: Os arquivos PDF são salvos com um padrão de nome organizado e único.
- **Reemissão Inteligente**: Emita certificados apenas para participantes pendentes em um lote já configurado.
- **Relatórios**: Visualize estatísticas de emissão e exporte dados detalhados para uma nova Planilha.
- **Página de Ajuda Integrada**: Uma documentação completa acessível diretamente pelo menu da ferramenta.

## 🛠️ Tecnologias Utilizadas

- Google Apps Script
- Google Sheets
- Google Docs
- Google Drive
- HTML / CSS / JavaScript (para as interfaces)
- Bootstrap 4 (para estilização)

## 🚀 Como Instalar e Configurar

Siga estes passos para ter sua própria cópia da ferramenta funcionando:

**1. Crie e Configure a Planilha Principal:**
   - Crie uma **nova Planilha Google** no seu Google Drive. Este arquivo será o cérebro do sistema.
   - Dentro desta planilha, você precisará criar 4 abas (páginas). Renomeie as abas e adicione os seguintes cabeçalhos na primeira linha de cada uma, **exatamente como descrito abaixo**:

   * **Aba 1: `Cadastro.Instituicoes`**
       *(Onde você irá cadastrar os parceiros)*
       - `IDParceiro`
       - `NomeInstituicao`
       - `Responsavel`
       - `EmailContato`
       - `Telefone`
       - `DataCadastro`

   * **Aba 2: `Config.Salvas`**
       *(Esta aba será preenchida automaticamente pelo sistema)*
       - `ConfigID`
       - `idParceiro`
       - `nomeParceiro`
       - `NomeCertificado`
       - `NomeEvento`
       - `DataEvento`
       - `DataEmissaoProgramada`
       - `MensagemCorpo`
       - `idSheetParticipantes`
       - `TemplateDocID`
       - `TargetFolderID`
       - `DataCriacao`

   * **Aba 3: `CertificadosEmitidos`**
       *(Esta aba será preenchida automaticamente pelo sistema)*
       - `ConfigID`
       - `idParceiro`
       - `nomeParceiro`
       - `NomeParticipante`
       - `EmailParticipante`
       - `DataEmissao`
       - `LinkCertificadoPDF`
       - `CodigoVerificador`

   * **Aba 4: `Logs.Atividades`**
       *(Esta aba será preenchida automaticamente pelo sistema)*
       - `Timestamp`
       - `Ação Realizada`
       - `Detalhes`

**2. Instale o Código do Projeto:**
   - Abra a planilha que você acabou de configurar.
   - Vá em `Extensões > Apps Script`.
   - O editor de script será aberto. Você verá alguns arquivos iniciais (como `Code.gs`).
   - Para cada arquivo de código neste repositório GitHub (`.gs` e `.html`), crie um arquivo correspondente no editor do Apps Script (clicando no `+` > `Script` ou `HTML`).
   - Copie o conteúdo de cada arquivo do GitHub e cole no arquivo de mesmo nome dentro do editor.
   - Salve cada arquivo no editor (`Ctrl+S`).

**3. Configure seus Arquivos e Pastas Externas:**
   - **Crie uma Pasta no Google Drive:** Crie a pasta onde os certificados em PDF serão salvos. Copie o ID desta pasta (da URL).
   - **Crie um Template no Google Docs:** Crie um documento com o design do seu certificado e use as tags (`{{NOME_PARTICIPANTE}}`, etc.). Copie o ID deste documento.
   - **Crie sua Planilha de Participantes:** Para cada evento, crie uma nova Planilha Google para listar os participantes. Siga a estrutura de colunas **obrigatória** definida na página de ajuda da ferramenta. Copie o ID desta planilha.

**4. Autorize e Use:**
   - Volte para a sua planilha principal e **recarregue a página** (F5).
   - Um novo menu, **"🛠️ Emissor de Certificados"**, aparecerá.
   - Clique em qualquer opção do menu pela primeira vez. O Google solicitará autorização para o script rodar. Siga os passos e permita o acesso.
   - Pronto! A ferramenta está pronta para ser usada. Consulte o menu `Ajuda e Instruções` para um guia detalhado de uso.

## 📄 Licença

Este projeto está licenciado sob os termos da [Creative Commons Atribuição-NãoComercial 4.0 Internacional (CC BY-NC 4.0)](https://creativecommons.org/licenses/by-nc/4.0/).

Você pode usá-lo, modificá-lo e compartilhá-lo **para fins não comerciais**, desde que com a devida atribuição a **Vitor Krewer**.  
Para qualquer uso comercial, entre em contato diretamente.

---

## 🤝 Contato

[LinkedIn](https://www.linkedin.com/in/vitorkrewer) • [Email](mailto:vitormkrewer@gmail.com)

---
