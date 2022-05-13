# -*- coding: utf-8 -*-
"""
Created on Thu Oct 10 20:35:22 2019

@author: duan
"""
from PIL import Image #图片处理
import datetime
import time
from selenium import webdriver #浏览器自动化库
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.select import Select
import pandas as pd
# from tqdm import tqdm #进度条库
from chinese_calendar import is_workday
# 注意：安装语句中库名有变动：pip install chinesecalendar
import ddddocr
# 注意：现在仅支持 <=python 3.9
import vthread
# 多线程库，比threading好用
import os,sys
from pyautogui import confirm,prompt

# def del_yzmpng():#删除指定或当前路径中指定类型的文件
#   for root , dirs, files in os.walk(sys.path[0]):#sys.path[0]为当前路径
#     for name in files:
#       if name.endswith(".png"):   #指定要删除的格式，这里是jpg 可以换成其他格式
#         os.remove(os.path.join(root, name))
#         print ("Delete File: " + os.path.join(root, name))

fpath='基础数据.xlsx'
fpath2='zgr.xlsx'
fpath3='导入数据.xlsx'
io = pd.io.excel.ExcelFile(fpath)#用io可以提升读取多工作表数据的效率！不用也可以，对单工作表没影响
df1 = pd.read_excel(io, sheet_name='课表')
df3 = pd.read_excel(io, sheet_name='教学内容')
df4 = pd.read_excel(io, sheet_name='时段')
df5 = pd.read_excel(io, sheet_name='账号')
dftx= pd.read_excel(fpath2, sheet_name=0)
io.close

user = str(df5['账号'][0])
upwd = str(df5['密码'][0])

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
# print(gzr)
# print("---------------------------------------------------------------------------------------------------------")
#替换调休日的星期
tempdf = dftx.set_index('工作日')['星期']
gzr['星期'] = gzr['工作日'].map(tempdf).fillna(gzr['星期']).astype(str)
# print(gzr)
# print("---------------------------------------------------------------------------------------------------------")

#取如果存在多个班名，并放在列表中
df1temp=df1.drop_duplicates(['班'])
ix=list(df1temp['班'])
# print(ix)
# print("---------------------------------------------------------------------------------------------------------")
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
    # print(df6)
    #删除确定的工作日范围课表与教学内容总节次不匹配的行
    df6=df6.dropna(axis=0,subset=['学科','教学内容'])
    df=df.append(df6,ignore_index=True)
    df=df.reset_index(drop=True)

# 按日期和节次排序录入的正式数据，
# 便于如果自动录入过程当中网络等其它非正常因素中断程序后，
# 查出相应时段，补录。感觉这是最优化的！！
df.sort_values(by=['工作日','上下午','节次','老师'],ascending=True,inplace=True)

#输出检查要导入的数据表！！！，先检查，后开始选择浏览器方式自动导入数据
df.to_excel(fpath3)

# gogo=confirm(text="打开 导入数据.xlsx ，检查有无错误，最后一定要关闭xlsx文件！！\n点击继续，正常导入数据；\n点击退出，结束程序，检查后再来！",title="检查  导入数据.xlsx",buttons=["继续","退出"])
gogo=prompt(text="1、打开、检查、关闭当前路径下的->导入数据.xlsx \n2、点击Cancel，结束程序，检查后重新再来！\n3、根据自身硬件情况设置同时运行的线程数，点击OK录入数据",title="检查  导入数据.xlsx",default="6")
if gogo == None:
    sys.exit()

# 自定义多线程数，多线程同时完成一项总任务
xianc_i=int(gogo)

# 新建线程函数
@vthread.thread
def xiancheng(df,n):
    jp_png="jie_tu_quanping"+str(n)+".png"
    yzmpng="yzm"+str(n)+".png"
    
    if xianc_i > 10 :
        #无头浏览器方式，可提高效率.但不能双开已登陆的网站。
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        bs = webdriver.Chrome(chrome_options=chrome_options)
    else:
        #可视浏览器，可视化，但效率低些
        bs = webdriver.Chrome() 
    
    bs.implicitly_wait(60)#网络卡顿，暂时选用的是隐式等待。最长延时等待单页面全局60秒
    bs.get('http://219.153.116.114:19001/Function/Login/login.html?appid=006884ba65664b44a66f3428313b773d&ist=0')
    bs.maximize_window()#全屏处理
    
    print ("请等待，数据处理中...")
    time.sleep(1) 
    
    # 此网站的动态化的验证码，用截屏加裁图的方式获得相应的验证码图片可行
    def caitu():
        bs.get_screenshot_as_file(jp_png)
        location = bs.find_element_by_id('Verify_codeImag').location
        size =bs.find_element_by_id('Verify_codeImag').size
        left = location['x']
        top =  location['y']
        right = location['x'] + size['width']
        bottom = location['y'] + size['height']
        a = Image.open(jp_png)
        im = a.crop((left,top,right,bottom))
        im.save(yzmpng,'PNG')
    def login():
        caitu()
        
        # 识别验证码
        ocr = ddddocr.DdddOcr()              # 实例化
        with open(yzmpng, 'rb') as f:     # 打开图片
            img_bytes = f.read()             # 读取图片
        yzmok = ocr.classification(img_bytes)  # 识别
        print(yzmok)
        
        # 填入用户名、密码和验证码
        account = bs.find_element_by_id('account')
        account.clear()
        account.send_keys(user)

        pwd = bs.find_element_by_id('userpwd')
        pwd.clear()
        pwd.send_keys(upwd)
        
        yzmcode = bs.find_element_by_id('txtCode')
        yzmcode.send_keys(yzmok)
        time.sleep(2)

        # 删除当前路径中，处理验证码时临时png图片
        os.remove(jp_png)
        os.remove(yzmpng)
        time.sleep(2)

    # 登陆，并进入新增数据录入界面，注意网络因素，加多页面加载的延时时间
    login()
    time.sleep(4)
    #谷歌浏览器查看网页元素xpath的步骤：F12-左上角箭头-选择查看的元素右键选中-返回Elements代码框-右键copy-选择xpath
    bs.find_element_by_xpath("/html/body/div[1]/aside[1]/div/section/div/ul/li[9]/a/span[1]").click()
    time.sleep(8)
    bs.find_element_by_xpath("/html/body/div[1]/aside[1]/div/section/div/ul/li[9]/ul/li[2]/a").click()
    time.sleep(6)
    iframe = bs.find_elements_by_tag_name("iframe")[1]
    bs.switch_to.frame(iframe)
    time.sleep(6)
    bs.execute_script("javascript:add()")
    time.sleep(6)
    bs.switch_to.default_content()
    time.sleep(6)
    iframe = bs.find_elements_by_tag_name("iframe")[2]
    bs.switch_to.frame(iframe)
    time.sleep(6)
    
    #遍历上传数据所有行，含索引。并设置进度条范围
    pbar = df.itertuples()
    ok_n=0#成功计次
    
    for row in pbar:
        #对应的字段列表：学期、老师、上下午、节次、校区、学科、年级、班、功能室、班班通类别、使用器材、星期、周节次顺序、工作日、教学内容
        ok_n+=1
        
        s3 = bs.find_element_by_id('Shiyongdate')#如果不行再设法处理那个网页中这个txet元素的只读属性为可写
        js='document.getElementById("Shiyongdate").removeAttribute("readonly");'
        bs.execute_script(js)
        time.sleep(1)
        
        #方式一：清除日期后再发送日期
        s3.clear()
        s3.send_keys(getattr(row, '工作日'))
        time.sleep(1)
        # #方式二：用js方法输入日期
        # js_value = 'document.getElementById("Shiyongdate").value=' + getattr(row, '工作日')
        # bs.execute_script(js_value)
        s14 = bs.find_element_by_id('Syqc')
        s14.clear()
        s14.click()
        
        s4 = bs.find_element_by_id('Sytime2')
        s4.send_keys(getattr(row, '上下午'))
        s5 = bs.find_element_by_id('Sytime3')
        s5.send_keys(getattr(row, '节次'))
        s6 = bs.find_element_by_id('Xiaoquid')
        s6.send_keys(getattr(row, '校区'))   
        s7 = bs.find_element_by_id('Xueke')
        s7.send_keys(getattr(row, '学科'))
        
        s2 = bs.find_element_by_id('Syteacher')
        s2.send_keys(getattr(row, '老师'))#平台中老师名字中没有“陈萍”，只有“陈平”记得要改名字
        time.sleep(1)
        
        selectFruit1 = bs.find_element_by_id('Gradeid') 
        Select(selectFruit1).select_by_visible_text(getattr(row, '年级'))
        
        s9 = bs.find_element_by_id('Classid')
        Select(s9).select_by_visible_text(getattr(row, '班'))
        
        s11 = bs.find_element_by_id('Syjk_gnsid')
        s11.send_keys(getattr(row, '功能室'))
        
        s13 = bs.find_element_by_id('Shoukenr')
        s13.clear()
        s13.send_keys(getattr(row, '教学内容'))
        time.sleep(1)
        
        #条件有错！！！
        s14 = bs.find_element_by_id('Syqc')
        s14.clear()
        if getattr(row, '使用器材') == getattr(row, '使用器材'):
            s14.send_keys(str(getattr(row, '使用器材')))
        time.sleep(1)
        
        #使用情况：TextBox11 元素，如果要加的话，则使用与使用器材相似的语句，并在教师填写的原表中加上使用情况一列，放在最后！
        if getattr(row, '班班通类别') == getattr(row, '班班通类别'):
            s12 = bs.find_element_by_id('Kemutype')
            s12.send_keys(getattr(row, '班班通类别'))
        time.sleep(1)
        
        s1 = bs.find_element_by_id('Xueqinum')
        s1.send_keys(getattr(row, '学期'))
        
        time.sleep(1)
        # #这是以前点确认的方式
        # btn_add = bs.find_element_by_id('layui-layer-btn0')#确认按钮元素还重新去查看定位，这个暂时不对
        # btn_add.click()
        
        #这是提交表单的方式------------2021年用的这个方式！2022年测试不成功改回这个方式。
        bs.execute_script("javascript:submitForm()")
        time.sleep(5)
            
        print("成功录入 %s 老师 %s 的课程，现已录入 %d 条数据."  % (getattr(row, '老师'),getattr(row, '工作日'), ok_n))
    
    bs.close()
    print ("共成功添加",ok_n,"条数据！")
    time.sleep(1)


#2022年3月15日新加：直接再从检查后的导入数据.xlsx文件中
# 导入成新的df数据集，以便灵活处理录入的数据
io = pd.io.excel.ExcelFile(fpath3)#用io可以提升读取多工作表数据的效率！不用也可以，对单工作表没影响
dfok = pd.read_excel(io, sheet_name=0)
io.close

# 总任务行数
maxlen=len(dfok)
# 如果总任务数还比线程数少，则线程数设为总任务数
if maxlen < xianc_i:
    xianc_i=maxlen

# 按设置的线程数循环平均分成相应的数据集，并开启所有线程同时自动录入数据
# 这里的i是用于计线程数，也用于按线程数平均划分数据集，也用于各线程的验证码图片文件命名，以区分处理！
for i in range(0,maxlen,int(maxlen/xianc_i)):
    xianc_df = dfok.iloc[i:i+int(maxlen/xianc_i),:]
    xiancheng(xianc_df,i)