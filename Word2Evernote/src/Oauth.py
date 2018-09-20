# -*- coding: utf-8 -*-
import requests
import urllib
import hashlib
import os

class Oauth(object):
    def __init__(self, consumerKey, consumerSecret, account, passwd, sandbox=True, isInternational = False):
        self.consumerKey = consumerKey
        self.consumerSecret = consumerSecret
        self.accout = account
        self.passwd = passwd
        if sandbox:
            self.host = 'sandbox.evernote.com'
        elif isInternational:
            self.host = 'app.evernote.com'
        else:
            self.host = 'app.yinxiang.com'

        # self.host = 'app.yinxiang.com'

    def __get_tmp_token(self):
        params = {
            'oauth_callback': '127.0.0.1',
            'oauth_consumer_key': self.consumerKey,
            'oauth_signature': self.consumerSecret,
            'oauth_signature_method': 'PLAINTEXT',
        }

        r = requests.get('https://%s/oauth' % self.host, params=params)
        if not 'oauth_token' in r.text:
            self.tmpOauthToken = None
        else:
            self.tmpOauthToken = str(dict(item.split('=', 1) for item in urllib.unquote(r.text).split('&'))['oauth_token'])

    def oauth(self):
        if os.path.isfile('C:\\tmp_wordextract'):
            with open('C:\\tmp_wordextract', 'r') as f:
                apDict = {}
                for line in f.readlines():
                    if line.replace('\n', ''):
                        key, val = line.split("####", 1)
                        apDict[key] = val
            if self.__get_md5() in apDict.keys():
                return apDict[self.__get_md5()]
        self.__get_tmp_token()
        self.__get_ver()
        token_ = self.__get_token()
        if token_:
            with open('C:\\tmp_wordextract', 'a') as f:
                f.write('\n' + self.__get_md5() + "####" + str(token_[0]))
            return token_[0]
        else:
            return None

    def __get_login_info(self):
        account = self.accout
        password = self.passwd
        return account, password

    def __get_md5(self):
        return  hashlib.md5(str(self.accout) + str(self.passwd)).hexdigest()

    def __get_ver(self):
        if not self.tmpOauthToken:
            self.verifier = None
            return
        account, password = self.__get_login_info()
        access = {
            'authorize': 'Authorize',
            'oauth_token': self.tmpOauthToken,
            'username': account,
            'password': password,
        }
        r = requests.post('https://%s/OAuth.action'%self.host, data = access)
        if 'oauth_verifier' in r.url:
                self.verifier = dict(item.split('=', 1) for item in r.url.split('?')[-1].split('&'))['oauth_verifier']
        else:
            self.verifier = None

    def __get_token(self):
        if not self.verifier:
            return None
        payload = {
            'oauth_consumer_key': self.consumerKey,
            'oauth_token': self.tmpOauthToken,
            'oauth_verifier': self.verifier,
            'oauth_signature': self.consumerSecret,
            'oauth_signature_method': 'PLAINTEXT',
        }
        r = requests.get('https://%s/oauth'%self.host, params = payload)

        if not ('oauth_token' in r.text and 'edam_expires' in r.text):
            raise Exception('Token Not Found')
        return (dict(item.split('=',1) for item in urllib.unquote(r.text).split('&'))['oauth_token'],
            dict(item.split('=',1) for item in urllib.unquote(r.text).split('&'))['edam_expires'], self.host)







