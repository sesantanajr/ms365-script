
# GERENCIAMENTO DE LICENÇAS VIA POWERSHELL

## Descrição
Este repositório contém um script PowerShell para automação de tarefas no Microsoft 365. O script facilita a administração e o gerenciamento dos serviços Microsoft 365, oferecendo uma solução eficiente e prática para administradores de TI.

![Logo Jornada 365](https://github.com/sesantanajr/ms365-script/blob/main/Walpapper%204.png)

## Propósito
O script foi criado para atender a diversas necessidades administrativas dentro do ambiente Microsoft 365, incluindo:

- Gerenciamento de Licenças Microsoft 365
- Instalação/Atualização do modulo Microsoft Graph e dependências
- Automatização de tarefas rotineiras, permitindo que os administradores de TI.


## Funcionalidades
- **Automatização de Tarefas Administrativas:** Reduz o tempo e o esforço necessários para realizar tarefas repetitivas.
- **Gerenciamento de licenças:** Facilita o gerenciamento de licenças, quando é necessário atribuir dezenas ou centenas de licenças.
- **Informações em Tela** Durante a remoção ou aplicação das licenças, é sempre informado quais contas foram aplicadas/removidas o licenciamento, se houve falha ou sucesso na aplicação.
- **Instalação de Modulos facilitado:** O Script faz tudo para você, verifica se o modulo esta instalado e também atualiza se houver necessidade.
- **Cria um arquivo csv de exemplo:** Para gerenciar as licenaçs é necessário ter um arquivo csv com 1 coluna Email e as contas de todos que serão atribuidas ou removidas as licenças. O script cria um arquivo de exemplo para facilitar o uso. O arquivo é criado na pasta C:\MS365.

## Requisitos
Para utilizar o script, você precisar dos seguintes requisitos:
- **PowerShell 5.1 ou superior**
- **Permissões de administrador no Microsoft 365**
- **O script cria uma pasta no diretório C:\MS365**

## Como Usar
Siga os passos abaixos para utilizar o script. Abra o Powershell como Adminsitrado:

1. **Clone este repositório:**
   ```sh
   cd  ~\Downloads\
   git clone https://github.com/sesantanajr/ms365-script.git
   ```

2. **Navegue até o diretório do script**
   ```sh
   cd ms365-script
   ```

3. **Execute o script:*
   ```sh
   powershell -ExecutionPolicy Bypass -File .\ms365.ps1
   ```

## Sobre
Este projeto foi desenvolvido por [Sérgio Sant'Ana Júnior](https://www.linkedin.com/in/sergiosantanacloud/) para a Jornada 365. Para mais informações e outros projetos, visite nosso [site](https://jornada365.cloud).

## Contribuições
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues e enviar pull requests. Se você encontrar um problema ou tiver uma sugestão de melhoria, não hesite em nos contatar.

**Jornada 365 - Sua Jornada Começa Aqui**
