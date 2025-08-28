📋 Descrição do Projeto

Sistema completo de gestão de ordens de serviço para oficinas mecânicas, desenvolvido em Python com interface gráfica PyQt5. Permite criar, editar, visualizar e imprimir ordens de serviço com integração direta para geração de PDFs profissionais.
✨ Funcionalidades Principais

    Gestão completa de OS: Criação, edição, busca e exclusão de ordens de serviço

    Interface intuitiva: Design moderno e fácil de usar com validação de dados em tempo real

    Geração de PDF: Criação automática de ordens de serviço em formato PDF com layout profissional

    Armazenamento em Excel: Todos os dados são salvos em planilha Excel para fácil backup e portabilidade

    Consulta de CEP: Integração com API ViaCEP para preenchimento automático de endereços

    Cálculos automáticos: Sistema de cálculos de valores totais com descontos e impostos

    Gestão de peças e serviços: Controle detalhado de itens com tipos, quantidades e valores

🛠️ Tecnologias Utilizadas

    Python 3.7+: Linguagem de programação principal

    PyQt5: Framework para interface gráfica

    Pandas: Manipulação e armazenamento de dados em Excel

    WeasyPrint: Geração de PDFs a partir de templates HTML

    Jinja2: Template engine para formatação de documentos

    Requests: Comunicação com APIs externas (ViaCEP)

    Pillow: Processamento de imagens para logos

📦 Instalação e Configuração
Pré-requisitos

    Python 3.7 ou superior instalado

    Pip (gerenciador de pacotes do Python)

Passos de Instalação

    Clone o repositório:

bash

git clone <url-do-repositorio>
cd oficina-os-system

    Crie um ambiente virtual (recomendado):

bash

python -m venv venv
source venv/bin/activate  # Linux/Mac
# ou
venv\Scripts\activate     # Windows

    Instale as dependências:

bash

pip install -r requirements.txt

    Estrutura de pastas:

<img width="754" height="185" alt="image" src="https://github.com/user-attachments/assets/8bb1d7be-bcc7-4a80-877c-ca5507a165d3" />

    Configure o logo da oficina:

        Coloque uma imagem chamada logo.png na pasta resources/

        Dimensões recomendadas: 100x100 pixels

🖥️ Como Utilizar o Sistema
1. Iniciando a Aplicação

Execute o arquivo principal:
bash

python main.py

2. Criando uma Nova Ordem de Serviço

    Dados da OS: O sistema gera automaticamente um número sequencial

    Dados do Cliente: Preencha todas as informações do cliente

        O CEP será automaticamente completado com dados da ViaCEP

        Telefone e CPF/CNPJ são formatados automaticamente

    Dados do Veículo: Informações completas do veículo

        A quilometragem é formatada automaticamente com separadores

    Problemas e Serviços: Descreva em três seções:

        Problema Informado (pelo cliente)

        Problema Constatado (pela oficina)

        Serviço Executado (detalhamento do trabalho)

    Itens e Peças: Adicione peças e serviços com:

        Tipo (Mão de obra, Peça, Serviço)

        Referência/código

        Descrição detalhada

        Valor unitário e quantidade

        Percentual de desconto (se aplicável)

3. Gerenciando OS Existentes

    Buscar OS: Digite o número da OS no campo de busca

    Editar OS: Após buscar, faça as alterações necessárias

    Excluir OS: Use o botão "Deletar OS Atual" (com confirmação)

    Salvar Alterações: Sempre clique em "Salvar/Atualizar OS"

4. Gerando PDF

    Clique em "Gerar e Visualizar OS (PDF)" para criar um documento impresso

    O PDF será aberto automaticamente no visualizador padrão

    Os arquivos são salvos na pasta OS_Clientes/ com numeração automática

📊 Estrutura do Arquivo Excel

O sistema utiliza um arquivo Excel (Ordens_de_Servico.xlsx) com a seguinte estrutura:

<img width="551" height="284" alt="image" src="https://github.com/user-attachments/assets/efbf85d2-8498-4be4-9eec-a8af58b3a3fa" />
<img width="549" height="369" alt="image" src="https://github.com/user-attachments/assets/401c79f2-586b-4a2f-b4af-bff0e8311a1b" />
<img width="549" height="382" alt="image" src="https://github.com/user-attachments/assets/090d3301-8b11-4c2c-902f-e1f970dd9888" />
<img width="548" height="330" alt="image" src="https://github.com/user-attachments/assets/e36e9aeb-3c6c-45d2-bdc5-781d28df4750" />
<img width="547" height="139" alt="image" src="https://github.com/user-attachments/assets/36469709-680e-473a-a5f5-82772b5d6b23" />

🎨 Personalização
Modificando o Template do PDF

Edite o arquivo os_template.html para alterar o layout do PDF gerado. O template usa:

    HTML/CSS padrão

    Variáveis Jinja2 para dados dinâmicos

    Suporte a imagens base64 encoded

Alterando Informações da Oficina

Modifique a constante INFO_OFICINA no código fonte para atualizar:

    Nome da oficina

    Endereço completo

    CNPJ

    Telefones de contato

Customizando Validações

As validações de campos podem ser ajustadas nos métodos:

    _formatar_telefone_cpf_cnpj(): Formatação de documentos

    _formatar_quilometragem(): Formatação de KM

    _validate_user_data(): Validações personalizadas

🔧 Solução de Problemas
Problemas Comuns

    Erro ao gerar PDF:

        Verifique se o WeasyPrint está instalado corretamente

        Confirme que o template HTML existe e está acessível

    Falha na consulta de CEP:

        Verifique a conexão com internet

        Confirme se o serviço ViaCEP está disponível

    Erro ao salvar no Excel:

        Feche o arquivo Excel se estiver aberto em outro programa

        Verifique as permissões de escrita na pasta

    Logo não aparece no PDF:

        Confirme que o arquivo logo.png está na pasta resources/

        Verifique se a imagem está em formato suportado (PNG, JPG)

Logs e Debug

    Os logs detalhados são exibidos no console durante a execução

    Erros são registrados com timestamp para facilitar troubleshooting

📝 Licença

Este projeto é destinado para uso interno de oficinas mecânicas. Consulte os termos de uso para mais informações.
🤝 Suporte e Contribuições

Para reportar bugs ou sugerir melhorias:

    Verifique a documentação existente

    Consulte os logs de erro no console

    Entre em contato com a equipe de desenvolvimento

🔄 Atualizações Futuras

    Integração com sistema de estoque

    Controle de usuários e permissões

    Relatórios gerenciais e analytics

    Backup em nuvem automático

    Versão mobile para consulta rápida

Nota: Este sistema foi desenvolvido para otimizar o fluxo de trabalho em oficinas mecânicas, substituindo processos manuais por uma solução digital integrada.
