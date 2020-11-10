# encoding:utf-8
import requests,time
import base64

class Baiduorc():
    def __init__(self,AK,SK):
        self.AK=AK
        self.Sk=SK
        host = f'https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={AK}&client_secret={SK}'
        try:
            response = requests.get(host)
            if response:
                access_token = response.json()['access_token']
                self.access_token = access_token
            else:
                self.access_token = None
        except:
            self.access_token = None
        #获取鉴权接口,就是身份验证...
    def get_text(self,imagePath):
        '''
        识别文字
        :param imagePath: 图片地址
        :return: 返回结果
        '''

        request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic"
        # 二进制方式打开图片文件
        with open(imagePath,'rb')as f:
            img=f.read()
        img = base64.b64encode(img)
        params = {"image": img}
        access_token = self.access_token
        request_url = request_url + "?access_token=" + access_token
        headers = {'content-type': 'application/x-www-form-urlencoded'}
        response = requests.post(request_url, data=params, headers=headers)
        if response:

            return response.json()
        else:
            return None








