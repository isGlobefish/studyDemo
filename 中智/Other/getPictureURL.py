# -*- coding: utf-8 -*-
'''
@createTool    : PyCharm-2020.3.2
@projectName   : pythonProjectPy3.9
@originalAuthor: Made in win10.Sys design by deHao.Zou
@createTime    : 2021/1/10 15:26
'''

import os
import json
import time


import xlrd
import requests
from PIL import Image
from selenium import webdriver
from qiniu import Auth, put_file, etag
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains


# 上传本地图片获取网上图片URL
def get_image_url(imagePath):
    if str(imagePath).split('.')[-1] == 'jpg' or str(imagePath).split('.')[-1] == 'JPG':
        key = 'transitPic_abc' + '.' + str(imagePath).split('.')[-1]  # 七牛云网盘文件名
    elif str(imagePath).split('.')[-1] == 'png' or str(imagePath).split('.')[-1] == 'PNG':
        key = 'transitPic_123' + '.' + str(imagePath).split('.')[-1]  # 七牛云网盘文件名
    else:
        print("请检查图片格式！！！")
    # 七牛云密钥管理：https://portal.qiniu.com/user/key
    # 【账号：144714959@qq.com  密码：thebtx1997】
    access_key = "DZnCErimkn2yQrn4aYel3JX7vPXKRonlvDFoVh1e"
    secret_key = "FBEHIFyMG28nWZrn316df-ny5bmIz_LanRWtabCi"
    q = Auth(access_key, secret_key)
    bucket_name = "qiniu730173201"  # 七牛云盘名
    token = q.upload_token(bucket_name, key)  # 删掉旧图片
    ret, info = put_file(token, key, imagePath)  # 上传新图片
    baseURL = "http://zzsy.zeus.cn/"  # 中智二级域名
    subURL = baseURL + '/' + key
    pictureURL = q.private_download_url(subURL)  # 图片URL
    return pictureURL


get_image_url('C:/Users/Zeus/Desktop/123.jpg')


# 上传图片
def get_img_url(imagePath):
    key = 'abc.png'  # 七牛云网盘文件名
    # 七牛云密钥管理：https://portal.qiniu.com/user/key
    # 【账号：144714959@qq.com  密码：thebtx1997】
    access_key = "DZnCErimkn2yQrn4aYel3JX7vPXKRonlvDFoVh1e"
    secret_key = "FBEHIFyMG28nWZrn316df-ny5bmIz_LanRWtabCi"
    q = Auth(access_key, secret_key)
    bucket_name = "qiniu730173201"  # 七牛云盘名
    policy = {
        'callbackUrl': 'https://requestb.in/1c7q2d31',
        'callbackBody': 'filename=$(fname)&filesize=$(fsize)',
        'persistentOps': 'imageView2/1/w/200/h/200'
    }
    # token过期时间设置为3600秒
    token = q.upload_token(bucket_name, key, 3600, policy)  # 删掉旧图片
    ret, info = put_file(token, key, imagePath)  # 上传新图片
    baseURL = "http://zzsy.zeus.cn/"  # 中智二级域名
    subURL = baseURL + '/' + key
    pictureURL = q.private_download_url(subURL)  # 图片URL
    return pictureURL







































