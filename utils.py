import os
import pickle
from datetime import datetime

import pandas as pd
from kiteconnect import KiteConnect

import config


def get_client(api_key):
    kite = KiteConnect(api_key=api_key)
    pickle_path = os.path.join('token.pkl')

    def update_pickle():
        print("Login: https://kite.trade/connect/login?api_key=" + config.API_KEY)
        print('>>> ')
        request_token = input()
        data = kite.generate_session(request_token, api_secret=config.API_SECRET)
        token_data = {
            'date': datetime.now(),
            'access_token': data['access_token']
        }
        with open(pickle_path, 'wb') as handle:
            pickle.dump(token_data, handle, protocol=pickle.HIGHEST_PROTOCOL)

        return token_data

    if os.path.exists(pickle_path):
        try:
            with open(pickle_path, 'rb') as handle:
                token_data = pickle.load(handle)
        except Exception as e:
            token_data = update_pickle()

        if token_data['date'].date() != datetime.now().date():
            token_data = update_pickle()
    else:
        token_data = update_pickle()

    kite.set_access_token(token_data['access_token'])
    print('KITE API CONNECTED')
    return kite


def set_instrument_df(kite):
    instrument_df = pd.DataFrame(kite.instruments())
    instrument_df.to_csv(os.path.join('kite_master.csv'))


