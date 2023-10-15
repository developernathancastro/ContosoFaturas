from botcity.core import DesktopBot
import pandas as pd

def not_found(label):
    print(f"Element not found: {label}")

dados = pd.read_excel('./file/Contoso+Coffee+Shop+Invoices.xlsx')

bot = DesktopBot()
path_app = "C:\Program Files (x86)\Contoso, Inc\Contoso Invoicing\LegacyInvoicingApp.exe"

bot.execute(path_app)
bot.wait(2000)
bot.maximize_window()
bot.wait(1000)

if not bot.find("Invoices", matching=0.97, waiting_time=10000):
    not_found("Invoices")
bot.click()

def cadastrafaturas(data, conta, contato, valor, status):

    if not bot.find( "novo-registro", matching=0.97, waiting_time=10000):
        not_found("novo-registro")
    bot.click()

    if not bot.find( "date", matching=0.97, waiting_time=10000):
        not_found("date")
    bot.click_relative(47, 4)

    bot.type_keys(['shift', 'end'])
    bot.paste(data)

    bot.tab()
    bot.paste(conta)
    bot.tab()
    bot.paste(contato)
    bot.tab()
    bot.paste(valor)

    if not bot.find( "status-inicio", matching=0.97, waiting_time=10000):
        not_found("status-inicio")
    bot.click_relative(61, 4)

    coluna = status

    if coluna == "Uninvoiced":
        if not bot.find( "univoiced", matching=0.97, waiting_time=10000):
            not_found("univoiced")
        bot.click_relative(67, 25)

    elif coluna == "Invoiced":
        if not bot.find( "invoiced", matching=0.97, waiting_time=10000):
            not_found("invoiced")
        bot.click_relative(77, 47)

    else:
        if not bot.find( "paid", matching=0.97, waiting_time=10000):
            not_found("paid")
        bot.click_relative(58, 68)


    if not bot.find( "salvar", matching=0.97, waiting_time=10000):
        not_found("salvar")
    bot.click()


for coluna in dados.itertuples():
    cadastrafaturas(str(coluna[1]), str(coluna[2]), str(coluna[3]), str(coluna[4]), str(coluna[5]))

bot.alt_f4()









