# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 20:35:22 2019

@author: duan
"""
from PIL import Image #图片处理
import datetime
import time
import pytesseract #验证图文字
from selenium import webdriver #浏览器自动化库
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
from tqdm import tqdm #进度条库
import pandas as pd
# import numpy as np
from chinese_calendar import is_workday

fpath='基础数据.xlsx'
fpath2='zgr.xlsx'
io = pd.io.excel.ExcelFile(fpath)#用io可以提升读取多工作表数据的效率！不用也可以，对单工作表没影响
df1 = pd.read_excel(io, sheetname='课表')
df3 = pd.read_excel(io, sheetname='教学内容')
df4 = pd.read_excel(io, sheetname='时段')
df5 = pd.read_excel(io, sheetname='账号')
dftx= pd.read_excel(fpath2, sheetname=0)
io.close

#字符到日期
def str_date(strdate):
    date = datetime.datetime.strptime(strdate,'%Y-%m-%d')
    return date

#将日期转成字符串
def date_str(datestr):
    str=datestr.strftime('%Y-%m-%d')
    return str

def get_week_day(date):
    week_day_dict = {
            0 : '星期一',
            1 : '星期二',
            2 : '星期三',
            3 : '星期四',
            4 : '星期五',
            5 : '星期六',
            6 : '星期日',
            }
    return week_day_dict[date.weekday()]

#生成范围内所有工作日
gzr =pd.DataFrame(columns=('工作日','星期'))
begin=str_date(df4['开始'][0])
end=str_date(df4['结束'][0])

for i in range((end - begin).days+1):
    day = begin + datetime.timedelta(days=i)
    if is_workday(day):
       week=get_week_day(day)
       tmpday=date_str(day)
       gzr=gzr.append(pd.DataFrame({'工作日':[tmpday],'星期':[week],}),ignore_index=True)
print(gzr)
#替换调休日的星期
tempdf = dftx.set_index('工作日')['星期']
gzr['星期'] = gzr['工作日'].map(tempdf).fillna(gzr['星期']).astype(str)
print(gzr)

#取如果存在多个班名，并放在列表中
df1temp=df1.drop_duplicates(['班'])
ix=list(df1temp['班'])
print(ix)
df= pd.DataFrame()
#按班循环联合处理课表、工作日表、教学内容，形成一次性的上传数据
for ix_ban in ix:
    #右联合查询出工作日范围内所有的班课程表，并按日期和周节次排序，以便准确对应教学内容！！！！
    df6=pd.merge(df1[df1['班']==str(ix_ban)],gzr,on='星期',how='right')
    df6=df6.sort_values(by=['工作日','周节次顺序'], ascending=[True,True], inplace=False)
    #删除确定的工作日范围与课表星期不匹配的行
    df6=df6.dropna(axis=0,subset=['学科','老师'])
    #重定索引
    df6=df6.reset_index(drop=True)
    #添加教学内容列
    df6=pd.concat([df6,df3['教学内容']],axis=1)
    print(df6)
    #删除确定的工作日范围课表与教学内容总节次不匹配的行
    df6=df6.dropna(axis=0,subset=['学科','教学内容'])
    df=df.append(df6,ignore_index=True)
    df=df.reset_index(drop=True)
#输出检查要导入的数据表！！！，先检查，后开始选择浏览器方式自动导入数据
df.to_excel('导入数据.xlsx')

########################################################
############以下为浏览器方式导入数据######################
########################################################
ok = input(r"请先检查导入数据.xlsx文件是否异常！！再选择使用无头浏览器输入0，其它任意键使用可视浏览器方式登陆录入数据:")
if ok == "0" :
    #无头浏览器方式，可提高效率.但不能双开已登陆的网站。
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    bs = webdriver.Chrome(chrome_options=chrome_options)
else:
    #可视浏览器，可视化，但效率低些
    bs = webdriver.Chrome()   

user = str(df5['账号'][0])
upwd = str(df5['密码'][0])

bs.implicitly_wait(20)#网络卡顿，暂时选用的是隐式等待。最长延时等待单页面全局5秒

bs.get('http://218.201.82.184:8118')
#先进入网址http://218.201.82.184:8118/，并登陆成功后

account = bs.find_element_by_id('loginname')
account.clear()
account.send_keys(user)

pwd = bs.find_element_by_id('loginpwd')
pwd.clear()
pwd.send_keys(upwd)

yzmcode = bs.find_element_by_id('txt_validcode')

def caitu(yzmid):
    bs.get_screenshot_as_file('a.png')
    location = bs.find_element_by_id(yzmid).location
    size =bs.find_element_by_id(yzmid).size
    left = location['x']
    top =  location['y']
    right = location['x'] + size['width']
    bottom = location['y'] + size['height']
    a = Image.open("a.png")
    im = a.crop((left,top,right,bottom))
    return im

def clear_image(image):
    image = image.convert('RGB')
    width = image.size[0]
    height = image.size[1]
    noise_color = get_noise_color(image)
    
    for x in range(width):
        for y in  range(height):
            #清除边框和干扰色
            rgb = image.getpixel((x, y))
            if (x == 0 or y == 0 or x == width - 1 or y == height - 1 
                or rgb == noise_color or rgb[1]>100):
                image.putpixel((x, y), (255, 255, 255))
    return image

def get_noise_color(image):
	for y in range(1, image.size[1] - 1):
		# 获取第2列非白的颜色
		(r, g, b) = image.getpixel((2, y))
		if r < 255 and g < 255 and b < 255:
			return (r, g, b)

def binarization(image):
    #转成灰度图
    imgry = image.convert('L')
    #二值化，阈值可以根据情况修改
    threshold = 128
    table = []
    for i in range(256):
        if i < threshold:
            table.append(0)
        else:
            table.append(1)
    out = imgry.point(table, '1')
    return out

bs.maximize_window()#全屏处理
im=caitu('yzm')
im=binarization(clear_image(im))
code = pytesseract.image_to_string(im)
im.show()

code2=input("自动识别验证码为:" + code + "。请改为正确值：")
yzmcode.send_keys(code2 + "\n")

bs.get('http://218.201.82.184:8118/Admini/gongnsjk/gnssynewedit.aspx?deptid=d386f149-2259-483b-a87a-b54faedbc410&id=')
print ("请等待，数据处理中...")
#遍历上传数据所有行，含索引。并设置进度条范围
pbar = tqdm(df.itertuples())
ok_n=0#成功计次
for row in pbar:
    #对应的字段列表：学期、老师、上下午、节次、校区、学科、年级、班、功能室、班班通类别、使用器材、星期、周节次顺序、工作日、教学内容
    ok_n+=1
    bs.find_element_by_id('ckbSelect1').click()
    bs.find_element_by_id('ckbSelect3').click()

    s11 = bs.find_element_by_id('DropDownList2')
    s11.send_keys(getattr(row, '功能室'))

    s2 = bs.find_element_by_id('txtTeacher')
    s2.clear()
    s2.send_keys(getattr(row, '老师'))
    
    s3 = bs.find_element_by_id('TextBox2')
    s3.clear()
    s3.send_keys(getattr(row, '工作日'))
    
    s4 = bs.find_element_by_id('ddltime2')
    s4.send_keys(getattr(row, '上下午'))
    s5 = bs.find_element_by_id('ddltim3')
    s5.send_keys(getattr(row, '节次'))
    s6 = bs.find_element_by_id('ddlXiaoQu')
    s6.send_keys(getattr(row, '校区'))   
    s7 = bs.find_element_by_id('TextBox5')
    s7.send_keys(getattr(row, '学科'))

    s13 = bs.find_element_by_id('TextBox12')
    s13.send_keys(getattr(row, '教学内容'))

    #条件有错！！！
    if getattr(row, '使用器材') == getattr(row, '使用器材'):
        s14 = bs.find_element_by_id('txtSyqc')
        s14.send_keys(str(getattr(row, '使用器材')))

    s1 = bs.find_element_by_id('DropDownList1')
    s1.send_keys(getattr(row, '学期'))
    
    selectFruit1 = bs.find_element_by_id('ddlGrades') 
    Select(selectFruit1).select_by_visible_text(getattr(row, '年级'))
    
    s9 = bs.find_element_by_id('selClass')
    Select(s9).select_by_visible_text(getattr(row, '班'))
    
    #使用情况：TextBox11 元素，如果要加的话，则使用与使用器材相似的语句，并在教师填写的原表中加上使用情况一列，放在最后！
    if getattr(row, '班班通类别') == getattr(row, '班班通类别'):
        s12 = bs.find_element_by_id('ddl_kemutype')
        Select(s12).select_by_visible_text(str(getattr(row, '班班通类别')))
    
    btn_add = bs.find_element_by_id('btnadd')
    btn_add.click()
    
    #确认添加后的对话框
    confirm = bs.switch_to_alert()
    confirm.accept()
    #点击继续添加按钮    
    btn_aa = bs.find_element_by_id('btnClick')
    btn_aa.click()
    pbar.set_description("成功录入 %s 老师 %s 的课程，现已录入 %d 条数据."  % (getattr(row, '老师'),getattr(row, '工作日'), ok_n))
    time.sleep(1)
    pbar.update(1)

pbar.close()
bs.close()
print ("成功添加",ok_n,"条数据！")