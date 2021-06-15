import pandas  as pd
import numpy   as np
#import seaborn as sns
#import matplotlib.pyplot as plt
import win32com.client as win32

from twilio.rest import Client


def start():
    # Set Up Twilio Account
    while True:
        account = input('Input your Twilio Account SID: ')

        try:
            account = float(account)
            print('Provide Correct SID! \n')
        except ValueError:
            print('\n')
        if type(account) == str:
            break
        
    while True:
        token = input('Input your Twilio Token: ')

        try:
            token = float(token)
            print('Provide Correct Token! \n')
        except ValueError:
            print('\n')
        if type(token) == str:
            break

    ''' # Future CEO sales number Provider
    while True:
        selected_num = input('Input Sales Value to Check: ')

        try:
            selected_num = str(selected_num)
            print('Provide Correct SID! \n')
        except ValueError:
            print('\n')
        if type(selected_num) == float:
            break'''

    account_sid = account
    auth_token  = token
    client      = Client(account_sid, auth_token)

    # Load Dataset
    df_aux = []

    for i in range(6):
        df_raw01 = pd.read_excel(f'data/{i}.xlsx')

        df_aux.append( df_raw01 )

    df_raw02 = pd.concat(df_aux)


    # Statistical Values
    num_attributes = df_raw02.select_dtypes( include=['int64'] )
    max_   = num_attributes.apply( max )
    min_   = num_attributes.apply( min )
    mean   = num_attributes.apply( np.mean )
    median = num_attributes.apply( np.median )

    #sns.displot( df_raw02['Vendas'] ) # Future Plotly.
    #plt.show()

    # Select Expensives
    max_sllr = df_raw02.loc[df_raw02['Vendas'] >= 55000, 'Vendedor'].astype('str')  + '  - R$: ' + df_raw02.loc[df_raw02['Vendas'] >= 55000, 'Vendas'].astype('str')
    min_sllr = df_raw02.loc[df_raw02['Vendas'] <= 15050, 'Vendedor'].astype('str')  + '  - R$: ' + df_raw02.loc[df_raw02['Vendas'] <= 15050, 'Vendas'].astype('str')


    message = client.messages.create(
                                body=f'Max at 55000: {max_sllr.values}\n More Information in Your Email.',
                                from_='+', # Twilio Fone Number
                                to='+'     # Fone Number
                            )


    outlook        = win32.Dispatch('outlook.application')
    mail           = outlook.CreateItem(0)
    mail.To        = 'YOUREMAIL@gmail.com'
    mail.Subject   = 'Relatório de Vendas'
    mail.Body      = 'Message body'
    mail.HTMLBody  = f'''
    <p>Prezados, Segue o Email.</p>

    <p> ________________________ </p>
    <p>Média: {mean.values}</p>
    <p>Mediana: {median.values}</p>
    <p>Máximo: {max_.values}</p>
    <p>Minimo: {min_.values}</p>
    <p> ________________________ </p>

    <p><strong>Maior</strong> numero de Vendas.</p>
    <p>{max_sllr.values}</p>
    <p> ________________________ </p>

    <p><strong>Menores</strong> numeros de Vendas.</p>
    <p>{min_sllr.values}</p>
    <p> ________________________ </p>

    Att.

    '''
    mail.Send()
start()