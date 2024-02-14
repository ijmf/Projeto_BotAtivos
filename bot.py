"""
WARNING:

Please make sure you install the bot with `pip install -e .` in order to get all the dependencies
on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the bot.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at https://documentation.botcity.dev/
"""


# Import for the Web Bot
from aifc import Error
import shutil
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *
from botcity.web.util import element_as_select
from botcity.web.parsers import table_to_dict
from botcity.plugins.excel import BotExcelPlugin
from botcity.plugins.email import BotEmailPlugin
from pandas import *

# Instanciar o plug -in
email = BotEmailPlugin()

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

excel = BotExcelPlugin()
excel.add_row(["Nome", "Último", "Máxima", "Mínima",
              "Variação", "Var. %", "Vol.", "Hora"])


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    # Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    # Se executar pelo VScode comentar o trecho abaixo, executando pelo maestro necessário descomentar.

    # maestro.login(server="https://developers.botcity.dev",
    #               login="57444048-4a34-432e-985f-88d6252065f1",
    #               key="574_SFXHGJ4TTVUWBDXFN6ES")

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    # Obtendo credenciais do Maestro
    # usuario = maestro.get_credential("dados-login", "usuario")
    # senha = maestro.get_credential("dados-login", "senha")

    # Enviando alerta para o Maestro
    maestro.alert(
        task_id=execution.task_id,
        title="Iniciando processo",
        message=f"O processo de consulta foi iniciado",
        alert_type=AlertType.INFO
    )

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Chrome
    bot.browser = Browser.CHROME

    # Uncomment to set the WebDriver path
    bot.driver_path = r"C:\Treinamento BotCity\chromedriver-win64\chromedriver.exe"

    # Abrimos o site do Investing.
    bot.browse("https://br.investing.com/equities/most-active-stocks")

    bot.wait(3000)

    # Captura da tabela de dados com os nomes das cidades
    table_dados = bot.find_element(
        "/html/body/div[1]/div[2]/div[2]/div[2]/div[1]/div[5]/div[2]/div[1]/table", By.XPATH)
    # Transformação da tabela em um dicionário
    table_dados = table_to_dict(table=table_dados)

    # Supondo que table_dados seja sempre uma lista não vazia
    for linhas in table_dados:

        # Obtém os valores das colunas do dicionário
        # Obtém o nome do ativo dentro do dicionário
        var_Ativo = linhas.get("nome", "")
        # Obtém o valor "último" dentro do dicionário
        var_Ultimo = linhas.get("último", "")
        # Obtém o valor "máxima" dentro do dicionário
        var_Maxima = linhas.get("máxima", "")
        # Obtém o valor "mínima" dentro do dicionário
        var_Minima = linhas.get("mínima", "")
        # Obtém o valor da variação total dentro do dicionário
        var_Tot = linhas.get("variação", "")
        # Obtém o valor da porcentagem da variação dentro do dicionário
        var_Por = linhas.get("var_", "")
        # Obtém o valor do volume dentro do dicionário
        var_Volume = linhas.get("vol", "")
        # Obtém a hora que foi capturado dentro do dicionário
        var_Hora = linhas.get("hora", "")

        print(var_Ativo, var_Ultimo, var_Maxima, var_Minima,
              var_Tot, var_Por, var_Volume, var_Hora)

        maestro.new_log_entry(
            activity_label="ATIVOS",
            values={"NOME": f"{var_Ativo}",
                    "ULTIMO": f"{var_Ultimo}",
                    "MAXIMA": f"{var_Maxima}",
                    "MINIMA": f"{var_Minima}",
                    "VARIACAO": f"{var_Tot}",
                    "VAR_POR": f"{var_Por}",
                    "VOLUME": f"{var_Volume}",
                    "HORA": f"{var_Hora}",
                    })

    # Adiciona a linha ao Excel
        excel.add_row([var_Ativo, var_Ultimo, var_Maxima,
                      var_Minima, var_Tot, var_Por, var_Volume, var_Hora])

    excel.write(
        r"C:\Treinamento BotCity\Projetos\RelatorioAtivos\Infos_Ativos.xlsx")

    # Configure IMAP com o servidor Hotmail
    try:
        email.configure_imap("outlook.office365.com", 993)

    # Configure SMTP com o servidor Hotmail
        # smtp.office365.com ou smtp-mail.outlook.com
        email.configure_smtp("smtp-mail.outlook.com", 587)

    # Faça login com uma conta de email válida
        email.login("junio.str@hotmail.com", "Teste123@")

    except Exception as e:
        print(f"Erro durante a configuração do e-mail: {e}")

    # Definindo os atributos que comporão a mensagem
    para = ["junio.str@hotmail.com"]
    assunto = "Planilha Ativos"
    corpo_email = ""
    arquivos = [
        r"C:\Treinamento BotCity\Projetos\RelatorioAtivos\Infos_Ativos.xlsx"]

    # Enviando a mensagem de e -mail
    email.send_message(assunto, corpo_email, para,
                       attachments=arquivos, use_html=True)

    # Feche a conexão com os servidores IMAP e SMTP
    email.disconnect()
    print("Email enviado com sucesso")

    # Subindo arquivo de resultados
    caminho_arquivo_xlsx = r"C:\Treinamento BotCity\Projetos\RelatorioAtivos\Infos_Ativos.xlsx"
    caminho_pasta_xlsx = r"C:\Treinamento BotCity\Projetos\RelatorioAtivos"
    shutil.make_archive(caminho_arquivo_xlsx, 'zip', caminho_pasta_xlsx)
    maestro.post_artifact(
        task_id=execution.task_id,
        artifact_name="Infos_Ativos",
        filepath=caminho_arquivo_xlsx + ".zip"
    )

    # Alerta de email
    maestro.alert(
        task_id=execution.task_id,
        title="E-mail OK",
        message=f"E-mail enviado com sucesso",
        alert_type=AlertType.INFO
    )

    # Implement here your logic...
    ...

    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Reportando erro ao Maestro
    # maestro.error(
    # task_id=execution.task_id,
    # exception=erro,
    # screenshot="erro.png"
    # )

    # Uncomment to mark this task as finished on BotMaestro
    maestro.finish_task(
        task_id=execution.task_id,
        status=AutomationTaskFinishStatus.SUCCESS,
        message="Task Finalizada"
    )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
