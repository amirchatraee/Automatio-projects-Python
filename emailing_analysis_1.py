import os
import pySheets as ps
import pyEmail as pe
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

if __name__ == '__main__':
    service_providers = {
        'sp_id': '1kdvZquiSEK_xCujoZKmIOkyqDwkg_Sn4u02A9RBbX2E',
        'tab_name': 'service_provider',
        'range_': 'A:B'}
    mailling_lists = {
        'sp_id': '1kdvZquiSEK_xCujoZKmIOkyqDwkg_Sn4u02A9RBbX2E',
        'tab_name': 'mailing_lists',
        'range_': 'B2:AK'
    }


    service_providers_df = ps.readSheetToDf(**service_providers)
    service_providers_dict = ps.readSheetToDf(**service_providers).to_dict("records")

    mailling_lists_df = ps.readSheetToDf(**mailling_lists)

    query_id = 24582
    df, email_date = pe.fetch_email(subject=f'[Success] Redash query:{query_id}',
                                    From='reporting3_bi_de@auto1.com')
    # Fetching desired columns and sending csv file
    df['final_planned_handover_date'] = pd.to_datetime(df['final_planned_handover_date'], format='%Y-%m-%d')

    # Filter cars to be handed over in the next 3 days
    filtered = df[(df['final_planned_handover_date'] >= datetime.now() + timedelta(days=3))]
    service_provider= service_providers_df['Service provider']
    filtered=filtered.join(service_provider)
    filtered = filtered[
        ['vin', 'stock_number', 'car_make', 'car_model', 'car_handover_on', 'handover_location', 'order_datetime',
         'delivery_date', 'planned_handover_date', 'order_number', 'final_planned_handover',
         'final_planned_handover_date','shipping_provider_id','final_planned_handover_time', 'delivery_hub','delivery_type','workshop_name','Service provider','current_location']]
    filtered['Service provider']= [service for service in service_providers_dict if service["Location"] == filtered['current_location']][0]["Service provider"]
    # Send email with the attached .csv file

    delivery_hubs = df['delivery_hub']
    tab_name=[]
    today=datetime.today()
    for hub in delivery_hubs:
        subject = f'Autohero Reinigungsliste - Showroom Berlin-Spandau - 2022-03-04'
        #subject = tab_name + " - " + [today]
        filename = f'Show room Berlin Spandau.csv'
        # email_body = f'''
		# <p>Hallo,</p>
		# <p>Hereby you can find the list of car(s) whose final_planned_handover_date start from date of tomorrow.
        #
		# '''
        email_body = '<body>'
        '<p>Hallo,</p>'
        '<p>anbei findet ihr die Liste für die Übergaben der nächsten drei Tage.</p>'
        '<p>@Filiale Bitte zeigt dem Mitarbeiter wo er diese Fahrzeuge findet. </p>'

        '<p>Viele Grüße</p>'
        '<p>Autohero-Team</p>'
        # Prepare attached file
        attached_df = filtered[filtered['delivery_hub'] == hub]
        # Select email addresses from the mailing list
        if hub not in mailling_lists_df.columns:
            continue
        emails = mailling_lists_df[hub]

        # to = emails[0].split(',')
        # cc = list(emails[1:])
        to= ['amirhassan.chatraeeazizabadi@auto1.com']
        cc=[] #daniel.kostic@auto1.com
        text = ''
        # Send email with the attached .csv file
        # pe.send_email_with_csv(
        #     subject,
        #     attached_df,
        #     filename,
        #     to,
        #     cc,
        #     text,
        #     email_body,delimiter=',')


        # print(df.columns)
#---------------------------------------------------------------------------------------------------------------------
        # upload all the new cases to the archive sheet
        # filtered = df[(df['final_planned_handover_date'] >= datetime.now() + timedelta(days=3))]
        # filtered = filtered[
        #     ['vin', 'stock_number', 'car_make', 'car_model', 'car_handover_on', 'handover_location', 'order_datetime',
        #      'delivery_date', 'planned_handover_date', 'order_number', 'final_planned_handover',
        #      'final_planned_handover_date', 'final_planned_handover_time', 'delivery_hub']]
        #
        sheet_id = '1kdvZquiSEK_xCujoZKmIOkyqDwkg_Sn4u02A9RBbX2E'
        tab_name = 'copy'
        range_name: str = 'A1:J'
        filtered['final_planned_handover_date']=filtered['final_planned_handover_date'].dt.strftime('%Y-%m-%d')
        # df['final_planned_handover_time'] = df['final_planned_handover_time'].to_string()
        # df['order_datetime'] = df['order_datetime'].to_string()
        #
        # print(type(df['final_planned_handover_date'][0]))
        # print(type(df['final_planned_handover_time'][0]))
        # print(type(df['order_datetime'][0]))
        # print(type(df['stock_number'][0]))





        sheet = ps.PySheets(sheet_id, tab_name, range_name)
        # existing_exit_data = sheet.get()
        # existing_exit_data = ps.PySheets.list_to_pd_df(existing_exit_data['values'])
        cols = ['stock_number', 'vin', 'car_make', 'car_model',
                'final_planned_handover_time','delivery_hub','order_datetime','Service provider','delivery_type','final_planned_handover_date' ]
        # colst = ['stock_number', 'vin', 'car_make', 'car_model',
        #         'delivery_hub', 'delivery_type', 'workshop_name']


        #
        #print(df[cols])

        if filtered.shape[0] > 0:
            #sheet.append(filtered[cols])
            ps.clearUpdate(sheet_id, tab_name, range_name, filtered[cols])
            print(df.shape)
        else:
           print('Nothing to update')


        break

     service_providers_df.loc[service_providers_df['location']==filtered['current_location']] = service_providers_df['Service provider']
