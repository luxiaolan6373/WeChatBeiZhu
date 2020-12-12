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
        except Exception as err:
            print(err)
            input('等待结束')
            self.access_token = None

        #获取鉴权接口,就是身份验证...
    def get_text(self,image):
        '''
        识别文字
        :param imageP: 图片数据
        :return: 返回结果
        '''

        request_url = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic"
        # 二进制方式打开图片文件
        img = base64.b64encode(image)
        params = {"image": img}
        access_token = self.access_token
        request_url = request_url + "?access_token=" + access_token
        headers = {'content-type': 'application/x-www-form-urlencoded'}
        response = requests.post(request_url, data=params, headers=headers)
        try:
            if response:
                return response.json()
            else:
                return None
        except Exception as err:
            print('百度识图错误',err)
            return None
if __name__ == '__main__':
    AK = "Weui83nXBM1ox6ozFzPF9bng"
    SK = "4AEmIGiInMv7gzlATntAs3pjHZrlCrsK"
    bd = Baiduorc(AK="Weui83nXBM1ox6ozFzPF9bng", SK="4AEmIGiInMv7gzlATntAs3pjHZrlCrsK")
    imagePath = "dqyh.png"
    with open(imagePath, 'rb')as f:
        img = f.read()

    print( bd.get_text(image=img))
    input('回车退出')











