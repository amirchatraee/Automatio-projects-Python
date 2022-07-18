import pandas as pd
import numpy as np
import pyEmail as pe
import pySheets as ps
from datetime import date
import time
import logging

workshop_email_addresses = (
    '19E8rqTRpMrKUYobKSH_YTssvJycTLuXAm41Aa2WJ0TI',
    'prio',
    'A1:G')

# functio to send email in csv and html
def send_email_to_workshops(workshop_name, attached_df,today, to, cc, attach_csv_as_html=False):
    subject = f'quality check oreder - {workshop_name} - {str(today)}'
    filename = f'{workshop_name}.csv'

    email_body = f'''
            <p>Hallo,</p>
            <p>anbei finden Sie eine Liste von Fahrzeugen mit folgendem Zustand:.</p>
            <p> - Per RC, wo jedes Auto im Quality Check mit >2 Tagen im Status aufgelistet ist, bestellt. </p>
            <p>- Wie Autos bekommen Unvollkommenheiten/Fotos in Quality Check,</p>

            <p>Viele Grüße</p>
            <p>Autohero-Team</p>
            '''
    text = ''

    if attach_csv_as_html:
        print(f'Sending email (csv as HTML) to {workshop_name}')
        attached_df.index = list(range(1, len(attached_df) + 1))
        csv_as_html_str = attached_df.to_html() # returns <table> ... </table>
        email_body += csv_as_html_str
    # Send email with the attached .csv file
    print(f'Sending email to {workshop_name}')
    pe.send_email_with_csv(
        subject,
        attached_df,
        filename,
        to,
        cc,
        text,
        html=email_body, delimiter=',')

# reading csv file for testing

if __name__ == '__main__':
    today = date.today()
    df, email_date = pe.fetch_email(subject='[Success] Redash query:42167')
    df.replace(np.nan, '', inplace=True)

    #df = pd.read_csv(r'C:\Users\amirhassan.chatraeea\Desktop\Quality_check_order_mail-AC_2022_04_21.csv')

    df = df[['ref_id', 'stock_number', 'retail_country', 'b2b_deal_datetime',
         'contract_signed_on', 'first_published_date', 'sold_state', 'state',
         'first_retail_ready', 'sold_before_rr', 'pictures_received_on',
         'auto1return_on', 'auto1return_reason', 'car_handover_on', 'vin',
         'current_location', 'workshop_id', 'workshop_name', 'open_date',
         'refurb_ordered_date', 'prep_start_date', 'refurb_feedback_date',
         'refurb_auth_date', 'refurb_start_date', 'refurb_completed_date',
         'refurb_qa_order_date', 'refurb_qa_completed_date', 'completed_date',
         'cancelled_date', 'merchant_name_admin', 'merchant_id',
         'last_status_updated_on', 'completed_reason', 'cancel_reason',
         'estimated_complete_date', 'provider_estimated_complete_date', 'brand',
         'model', 'rework_date', 'last_service_on', 'inspection_due_date',
         'last_service_mileage', 'first_registration_date', 'maximum_budget',
         'current_published_date', 'prepared_for_entry_check', 'car_in_buffer',
         'car_arrived_in_workshop', 'refurb_etas', 'after_ref_auth_state']]

    wsdf = ps.readSheetToDf(*workshop_email_addresses)
    wsdf = wsdf[wsdf['Active']=='Yes']
    for col in ['recipients', 'cc']:
        wsdf[col] = wsdf[col].apply(lambda x: x.split(','))
        wsdf[col] = wsdf[col].apply(lambda x: [s.strip() for s in x])

    for ind, ws_info in wsdf.iterrows():
        ws = ws_info['Workshop Name']
        to = ws_info['recipients']
        cc = ws_info['cc']
        # to = ['amirhassan.chatraeeazizabadi@auto1.com']
        # cc = []
        try:
            print(f'Send email to {ws}')
            attach_csv_as_html = False
        
            send_email_to_workshops( ws,df, today, to=to, cc=cc, attach_csv_as_html = attach_csv_as_html)
            time.sleep(10)
        except Exception as e:
            logging.basicConfig(filename='logging.log', format='%(asctime)s %(message)s', level=logging.INFO)
            logging.exception(e)

