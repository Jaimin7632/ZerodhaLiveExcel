import os
import pickle
from datetime import datetime
from kiteconnect import KiteConnect

def get_client(api_key):
    kite = KiteConnect(api_key=api_key)
    pickle_path = os.path.join('token.pkl')

    def update_pickle():
        print(f"Login for https://kite.trade/connect/login?api_key={api_key}")
        access_token = input('ENTER ACCESS TOKEN >>> ')
        # request_token = get_enctoken(userid=id, password=client_password,
        #                              totp_token=totp_secret, api_key=api_key)
        # data = kite.generate_session(request_token, api_secret=api_secret)
        token_data = {
            'date': datetime.now(),
            'access_token': access_token
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
    return kite
