# -*- coding: utf-8 -*-

import login
import collectWeiboDataByKeyword
import urllib
import urllib2
import sys

uid = 'zhangml5511@hotmail.com'
psw = '081620zhang'

reload(sys)
sys.setdefaultencoding('utf8')

simLogin = login.WeiboLogin(uid, psw)
simLogin.login()

collectWeiboDataByKeyword.main();
