from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import sleep
import time
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
import xlrd
import os
# 设置默认字体
def chg_font(obj, fontname='微软雅黑', size=None):
    ## 设置字体函数
    obj.font.name = fontname
    obj._element.rPr.rFonts.set(qn('w:eastAsia'), fontname)
    if size and isinstance(size, Pt):
        obj.font.size = size

def readFile(file = "车牌.xlsx"):
    lastFile = xlrd.open_workbook(file)
    sheet = lastFile.sheet_by_index(0)
    res = []
    for i in range(0,sheet.nrows):
        res.append(sheet.row_values(i))
    return res

# 导出数据
def writeData(name, corName = ""):
    try:
        doc = Document("结果.docx")
    except:
        doc = Document()
    #chg_font(doc.styles['Normal'], fontname='宋体')
    tab = doc.add_table(rows=7, cols=2, style="Table Grid")  # 添加一个4行4列的空表
    tab.style.paragraph_format.space_before = Pt(0)
    tab.style.paragraph_format.space_after = Pt(0)
    tab.cell(0, 0).merge(tab.cell(0, 1))
    tab.cell(1, 0).merge(tab.cell(1, 1))
    tab.cell(2, 0).paragraphs[0].add_run("查验日期")
    
    tab.cell(2, 1).paragraphs[0].add_run(time.strftime("%Y-%m-%d", time.localtime()) )
    tab.cell(3, 0).paragraphs[0].add_run("所属运输公司")
    tab.cell(3, 1).paragraphs[0].add_run(corName)
    tab.cell(4, 0).paragraphs[0].add_run("车牌号")
    tab.cell(4, 1).paragraphs[0].add_run(name)
    tab.cell(5, 0).paragraphs[0].add_run("视频轨迹情况")
    tab.cell(6, 0).paragraphs[0].add_run("审验人")
    tab.cell(0, 0).paragraphs[0].add_run().add_picture("./img/轨迹.png", width=Inches(5.75))
    tab.cell(1, 0).paragraphs[0].add_run().add_picture("./img/实时视频_处理后.png", width=Inches(5.75))
    for i in range(len(tab.rows)):
        for j in range(len(tab.columns)):
            for item in tab.cell(i, j).paragraphs:
                item.paragraph_format.space_before=Pt(0)
                item.paragraph_format.space_after = Pt(0)
    doc.save('结果.docx')


def getByName(driver, name = "粤SGH6593"):
    waitNotEle(driver,".ant-spin-dot.ant-spin-dot-spin")
    inputTxt = driver.find_elements_by_css_selector(".ant-input.ant-input-lg")[0]
    inputTxt.clear()
    inputTxt.send_keys(name)
    sleep(1)
    # 等待搜索完毕
    while len(driver.find_elements_by_css_selector(".ant-tree-treenode-checkbox-checked")) > 20:
        sleep(0.1)
        
    # 查找主列表
    waitEle(driver,".ant-tree-treenode-checkbox-checked", 2)
    selectList = driver.find_elements_by_css_selector(".ant-tree-treenode-checkbox-checked")
    if (len(selectList) < 1):
        return {
            "state": "error",
            "msg": "没有找到该车牌的车"
        }
    eleClass = selectList[1].get_attribute("class")
    if (eleClass.find("ant-tree-treenode-switcher-open") == -1):
        selectList[1].click()
    # 获取公司名
    cor = driver.find_elements_by_css_selector(".ant-tree-title")[1].text
    cor = cor[0:cor.find("(")]
    waitEle(driver,".ant-tree-node-content-wrapper.ant-tree-node-content-wrapper-normal")
    sleep(1)
    # 查找子列表
    selectList = driver.find_elements_by_css_selector(".ant-tree-node-content-wrapper.ant-tree-node-content-wrapper-normal")

    pic = driver.find_elements_by_css_selector(".ant-tree-node-content-wrapper.ant-tree-node-content-wrapper-normal img")[0]
    picSrc = pic.get_attribute("src")
    if (picSrc.find("-online") != -1):
        waitClickEle(driver, ".ant-tree-node-content-wrapper.ant-tree-node-content-wrapper-normal")
        selectList[0].click()
    else:
        return {
            "state": "error",
            "msg": "该车辆没有在运行"
        }
    waitClickEle(driver,".bottom-center.amap-info-contentContainer button")
    button = driver.find_elements_by_css_selector(".bottom-center.amap-info-contentContainer button")
    waitClickEle(driver,".bottom-center.amap-info-contentContainer button")
    button[0].click()
    button[2].click()
    openWebs = {}
    # 判断打开的网页
    for item in driver.window_handles:
        driver.switch_to.window(item)
        if str(driver.current_url).find("monitorMap") != -1:
            openWebs["monitorMap"] = item
        elif str(driver.current_url).find("monitorVideo") != -1:
            openWebs["monitorVideo"] = item
        elif str(driver.current_url).find("trackPlayback") != -1:
            openWebs["trackPlayback"] = item

    print("正在处理轨迹页面")
    # 处理轨迹页面
    driver.switch_to.window(openWebs["trackPlayback"])
    waitEle(driver, ".amap-icon")
    sleep(4)
    driver.save_screenshot("./img/轨迹.png")
    driver.close()

    print("正在处理实时视频")
    #处理实时视频
    jsCode = 'let timeInt, isTimeOut = false, maxTimes = 120\n\
    function fun() {\n\
    console.log("left  " + maxTimes.toString())\n\
    let isBad = false\n\
    let hasSrc = false\n\
    maxTimes--\n\
    if (maxTimes <= 0) {\n\
        let a = document.createElement("div")\n\
        a.className = "python-get-class"\n\
        a.setAttribute("type", "error")\n\
        document.querySelector("body").appendChild(a)\n\
        console.log("超时")\n\
        clearInterval(timeInt)\n\
        return\n\
    }\n\
    videos = document.querySelectorAll("video")\n\
    videos.forEach(item => {\n\
        if (item.src) {\n\
            hasSrc = true\n\
            console.log(item.readyState)\n\
            if (item.readyState === 0) {\n\
                isBad = true\n\
            }\n\
        }\n\
    })\n\
    if (!isBad && hasSrc) {\n\
        let a = document.createElement("div")\n\
        a.className = "python-get-class"\n\
        a.setAttribute("type", "good")\n\
        console.log("所有视频加载完成")\n\
        document.querySelector("body").appendChild(a)\n\
        clearInterval(timeInt)\n\
    }\n\
    }\n\
    timeInt = setInterval(fun, 1000)'
    driver.switch_to.window(openWebs["monitorVideo"])
    waitClickEle(driver, ".ant-col.ant-col-4")
    driver.find_elements_by_css_selector(".ant-col.ant-col-4")[2].click()
    driver.execute_script(jsCode)
    waitEle(driver, ".python-get-class",130, 0.8)
    type = driver.find_elements_by_css_selector(".python-get-class")[0].get_attribute("type")
    if (type == "good"):
        sleep(5)
        driver.save_screenshot("./img/实时视频.png")
        picture = Image.open("./img/实时视频.png")
        picture = picture.crop((617, 115, 617 + 1144, 115 + 620))
        picture.save("./img/实时视频_处理后.png")
        pass
    else:
        driver.close()
        driver.switch_to.window(openWebs["monitorMap"])
        return {
            "state": "error",
            "msg": "读取视频超时"
        }
    driver.close()
    driver.switch_to.window(openWebs["monitorMap"])
    return {
            "state": "good",
            "msg": cor
    }

# 添加cookie
def addCookie(driver, s):
    a = s.split(";")
    print(a)
    t = {}
    for item in a:
        b = item.split("=")
        t[b[0]] = b[1]
        driver.add_cookie({
            "name":b[0],
            "value":b[1]
        })
    print(t)

# 等待浏览器加载元素
def waitEle(driver, ele, times=30, between = 0.3):
    try:
        WebDriverWait(driver, times, between).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ele))
        )
    except:
        print("等待元素出错" + ele)

# 等待元素消失
def waitNotEle(driver, ele, times=30):
    try:
        WebDriverWait(driver, times, 0.2).until_not(
            EC.presence_of_element_located((By.CSS_SELECTOR, ele))
        )
    except:
        print("等待元素出错" + ele)

def waitClickEle(driver, ele, times = 30):
    try:
        WebDriverWait(driver, times, 0.2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, ele))
        )
    except:
        print("等待元素出错")

if __name__ == "__main__":
    choose = input("是否打开浏览器： 1、否  2、是\n")
    x = 0
    if choose == "1":
        x = 1980
    if input("是否删除历史 1、否  2、是\n") == "2":
        try:
            os.remove("结果.docx")
            os.remove("错误.csv")
        except:
            pass
    # 读取文件
    cars = readFile("车牌.xlsx")
    driver = webdriver.Chrome()
    driver.set_window_size(1980,1280)
    driver.set_window_position(x, 0)
    # 打开登陆
    driver.get("http://112.94.64.104:8080/login")
    # 登陆
    waitEle(driver,".ant-input")
    inputBox =  driver.find_elements_by_css_selector(".ant-input")
    inputBox[0].send_keys("zhitou")
    inputBox[1].send_keys("123456")
    waitEle(driver,".ant-btn")
    driver.find_elements_by_class_name("ant-btn")[0].click()
    waitEle(driver,".ant-menu-submenu-title")
    # 跳转到监控页面
    driver.execute_script("window.location.href = 'http://112.94.64.104:8080/monitorMap'")
    for i in range(len(cars)):
        try:
            item = cars[i]
            print("正在开始" + str(item[0]) + "," + str(i + 1) + "/" + str(len(cars)))
            res = getByName(driver, item[0])
            if res["state"] == "good":
                writeData(item[0], res["msg"])
            else: 
                print(res["msg"])
                f = open('错误.csv','a')
                f.write(str(item[0]) + "," + res["msg"]+"\n")
                f.close()
        except:
            f = open('错误.csv','a')
            f.write(str(item[0]) + "," + "未知错误\n")
            f.close()
    print("程序结束")
    driver.quit()



    