import os
import uuid
import redis
import json
import requests_oauthlib
import config
import json
import base64
from os import path

class SP_Graph:
    def __init__(self, token_type='text',**kwargs):
        # Enable non-HTTPS redirect URI for development/testing.
        os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'
        # Allow token scope to not match requested scope. (Other auth libraries allow
        # this, but Requests-OAuthlib raises exception on scope mismatch by default.)
        os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
        os.environ['OAUTHLIB_IGNORE_SCOPE_CHANGE'] = '1'
        self.na_token_flag=True
        if ('filename' in kwargs) & ('redis_conn' in kwargs):
            raise TypeError('Cannot define both filename and redis connection')
        if token_type=='text':
            try:
                if path.exists(kwargs['filename']):
                    print('File exists')
                    with open(kwargs['filename']) as json_file:
                        self.token = json.load(json_file)
                    self.na_token_flag=False
                    self.file_storage=kwargs['filename']
                else:
                    print('File doesnt exist')
                    self.create_token_storage('text',kwargs['filename'])
                    self.na_token_flag=True
            except:
                #print('No filename passed')
                raise NameError('No Filename was passed')

        elif token_type=='redis':
            self.get_redis_conn(kwargs['redis_conn'])
            self.redis_params=kwargs['redis_conn']
            try:
                self.gtoken=json.loads(self.token_storage.get(self.redis_params['key']))
                self.na_token_flag=False
            except:
                print('Token does not exist')
                self.create_token_storage(token_type,self.redis_params['key'])
                self.na_token_flag=True

        if self.na_token_flag:
            old_graph=requests_oauthlib.OAuth2Session(config.CLIENT_ID,scope=config.SCOPES,redirect_uri=config.REDIRECT_URI)
            auth_url=self.login(old_graph)
            print('You need to authorize the application, please go to this url and paste back the response')
            print(auth_url)
            auth_response=input('Response URL:')
            try:
                self.gtoken=self.authorized(auth_response, old_graph)
                if token_type=='redis':
                    self.token_saver_redis(self.gtoken)
                elif token_type=='text':
                    self.token_saver_text(self.gtoken)
            except:
                raise ValueError('Token could not be obtained')

        refresh_url=config.AUTHORITY_URL +config.TOKEN_ENDPOINT
        extra = {
        'client_id': config.CLIENT_ID,
        'client_secret': config.CLIENT_SECRET,
        }
        if token_type=='redis':
            self.Graph=requests_oauthlib.OAuth2Session(config.CLIENT_ID,
                                                       token=self.gtoken,
                                                       auto_refresh_url=refresh_url,
                                                       auto_refresh_kwargs=extra,
                                                       token_updater=self.token_saver_redis)
        elif token_type=='text':
            self.Graph=requests_oauthlib.OAuth2Session(config.CLIENT_ID,
                                                       token=self.gtoken,
                                                       auto_refresh_url=refresh_url,
                                                       auto_refresh_kwargs=extra,
                                                       token_updater=self.token_saver_text)

    def set_siteid(self, domain,url):
        siteid_url='https://graph.microsoft.com/v1.0/sites/'+domain+':/sites/'+url+'?$select=id'
        info=self.Graph.get(siteid_url).json()
        self.siteid=info['id']
        self.drives=self.get_drives()
        self.__dnames={x['name']:x['id'] for x in self.drives['value']}

    def get_drives(self):
        check_drives='https://graph.microsoft.com/v1.0/sites/'+self.siteid+'/drives'
        drives=self.Graph.get(check_drives).json()
        return drives

    def set_drive(self,drive_name):
        try:
            self.__curr_drive=self.__dnames[drive_name]
        except:
            raise KeyError('Folder does not exist in drive')

    def get_drive_items(self):
        drive=self.__curr_drive
        try:
            iurl='https://graph.microsoft.com/v1.0/drives/'+drive+'/items/root/children'
            self.__drive_item={x['name']:x['id'] for x in self.Graph.get(iurl).json()['value']}
            return self.__drive_item
        except:
            raise ValueError('Error accessing drive')

    def get_drive_folders(self):
        drive=self.__curr_drive
        try:
            iurl='https://graph.microsoft.com/v1.0/drives/'+drive+'/items/root/children'
            it=self.Graph.get(iurl).json()
            #folders=list(filter(None,[i['name'] if 'folder' in i else None for i in fdjson['value']]))
            fls={}
            for i in it['value']:
                if 'folder' in i:
                    fls[i['name']]=i['id']
            return fls
        except:
            raise ValueError('Error accessing drive')

    def get_drive_folder_items(self,itemid):
        fltemp='https://graph.microsoft.com/v1.0/drives/{}/items/{}/children'
        try:
            fitems=self.Graph.get(fltemp.format(self.__curr_drive,itemid))
            if fitems.status_code==200:
                return fitems
            else:
                print(fitems.status_code)
                #return(fitems)
                raise RuntimeError('MS Graph call was unsuccesful')
        except:
            raise RuntimeError('Something went wrong')

    def get_subfolders(self,itemid):
        it=self.get_drive_folder_items(itemid).json()
        fls={}
        for i in it['value']:
            if 'folder' in i:
                fls[i['name']]=i['id']
        return fls

    def upload_file(self,folder,name, data, headers):
        UP_Path='https://graph.microsoft.com/v1.0/drives/{}/items/{}:/{}:/content'
        UP=UP_Path.format(self.__curr_drive, folder, name)
        try:
            res=self.Graph.put(UP,data=data, headers=headers)
            if res.status_code==201:
                print('File Uploaded')
                return res
            else:
                print(res.status_code)
                return(res)
                #raise RuntimeError('MS Graph call was unsuccesful')
        except:
            raise RuntimeError('Something went wrong')

    def token_saver_text(self,token):
        with open(self.file_storage, 'w') as outfile:
            json.dump(token, outfile)

    def token_saver_redis(self, token):
        print('saving token')
        self.token_storage.set(self.redis_params['key'],json.dumps(token))

    def login(self, graph):
        """Prompt user to authenticate."""
        auth_base = config.AUTHORITY_URL + config.AUTH_ENDPOINT
        authorization_url, state = graph.authorization_url(auth_base)
        graph.auth_state = state
        return authorization_url

    def authorized(self,authorization_response, graph):
        graph.fetch_token(config.AUTHORITY_URL +\
                            config.TOKEN_ENDPOINT, \
                            client_secret=config.CLIENT_SECRET,\
                            authorization_response=authorization_response)
        return graph.token

    def graphcall(self):
        """Confirm user authentication by calling Graph and displaying some data."""
        endpoint = config.RESOURCE + config.API_VERSION + '/me'
        headers = {'SdkVersion': 'sample-python-requests-0.1.0',
                   'x-client-SKU': 'sample-python-requests',
                   'SdkVersion': 'sample-python-requests',
                   'client-request-id': str(uuid.uuid4()),
                   'return-client-request-id': 'true'}
        graphdata = MSGRAPH.get(endpoint, headers=headers).json()
        return {'graphdata': graphdata, 'endpoint': endpoint, 'sample': 'Requests-OAuthlib'}

    def create_token_storage(self,ftype,destination):
        if ftype=='text':
            with open(destination, 'w') as outfile:
                outfile.write('--TOKENFILE---')
            self.file_storage=destination
        elif ftype=='redis':
            if type(destination) is str:
                try:
                    self.token_storage.set(destination,"{}")
                    self.redis_key=destination
                except:
                    print('Cannot store to redis')
            else:
                raise TypeError('Redis storage requires a destination [key]')
        else:
            raise TypeError('Selected token type does not exist, options are redis or text')

    def get_redis_conn(self,redis_conn):
        if type(redis_conn) is dict:
            if all(i in redis_conn for i in ['host','port','db']):
                try:
                    r=redis.Redis(host=redis_conn['host'],
                              port=redis_conn['port'],
                              db=redis_conn['db']
                             )
                    r.ping()
                    self.token_storage=r
                except:
                    raise ValueError('Couldnt connect to redis host, is it up?')

            else:
                raise ValueError('redis_conn does not have the necessary values')
        else:
            raise TypeError('redis_conn is not a dictionary')
