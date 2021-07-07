# 套件引入區
import tkinter as tk
import time
from tkinter import Button, Radiobutton, mainloop, messagebox
from tkinter.constants import FALSE
from tkinter.font import Font
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException,ElementClickInterceptedException,TimeoutException,ElementNotInteractableException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import datetime
import os
import threading
import ntplib


# 爬蟲物件
# ===================================================================
class Shopee():
    classes = ''
    context = ''
    num = 0
    sp =''
    sid =''
    Cid = []
    #建構子
    def __init__(self,classes,context,num,sp,sid):
        # 設定變數
        self.produceID ={}
        self.classes = classes
        self.context = context
        self.num = num
        self.sp = sp
        self.sid = sid

        # 讀取分類清單
        try:
            with open('Classes.txt','r',encoding='utf-8') as f:
                lines = f.readlines()
                for line in lines:
                    self.produceID[line.split(':')[0]] = line.split(':')[1].replace('\n','')
        except:
            messagebox.showerror('錯誤','請先建立類別清單')

        # 談窗提示
        try:
            messagebox.showinfo('資料確認區','您輸入的資料如下:\n%s\n%s\n%s\n您設定的範圍為從第%s頁第%s項'%(self.produceID[classes],context + loginWindows.getFileData('defult_text') ,num,sp,sid))
            messagebox.showinfo('即將啟動蝦皮','即將啟動蝦皮')
        except:
            try:
                messagebox.showinfo('資料確認區','您輸入的資料如下:\n%s\n%s\n%s\n您設定的範圍為從第%s頁第%s項'%(classes,context + loginWindows.getFileData('defult_text'),num,sp,sid))
                messagebox.showinfo('即將啟動蝦皮','即將啟動蝦皮')
            except:
                messagebox.showerror('預設文字未設定','您好請先聯絡管理人員，設定聊聊文字預設內容 !!')
                os._exit(0)
        
        #設定瀏覽器驅動
        options = Options()
        options.add_argument("--disable-notifications")
        options.add_argument("--user-data-dir=C:\\Program Files (x86)\\Default")
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        global browser
        try:
            browser = webdriver.Chrome(chrome_options=options)
        except:
            messagebox.showerror('權限不足','請以裝置管理員執行 !!')
            os._exit(0)
        global wait
        wait = WebDriverWait(browser,2)
        browser.get('https://shopee.tw/')
        browser.maximize_window()
        time.sleep(5)
        Nowtext.configure(text='瀏覽器作業中.....')
        global soup
        soup = BeautifulSoup(browser.page_source,"html.parser")
        
        # 主程序分解動作
        # -----------------------------
        # 循環檢測登入狀態
        self.loginStatus()
        # 進入分類選單
        self.enterClasses(self.classes)
        # 進入產品列表
        self.enterProduce(int(self.num))
    
    # 取得登入狀態
    def loginStatus(self):
        while True:
            Nowtext.configure(text='登入狀態檢測中.....')
            time.sleep(5)
            global soup
            status = soup.find_all('a',{'class':'navbar__link navbar__link--account navbar__link--tappable navbar__link--hoverable navbar__link-text navbar__link-text--medium'})
            if len(status) != 0:
                messagebox.showinfo('未登入通知','請將網頁蝦皮登入，完成後回到這裡按下確定')
                soup = BeautifulSoup(browser.page_source,"html.parser")
            else:
                Nowtext.configure(text='已完成登入')
                time.sleep(2)
                break

    # 進入分類
    def enterClasses(self,id):
        errorNum = 0
        try:
            text = self.produceID[id]
        except:
            text = id
        
        time.sleep(5)

        # 檢測彈窗廣告
        try:
            browser.find_element_by_class_name('shopee-popup__close-btn').click()
            Nowtext.configure(text='已關閉彈窗廣告')
            time.sleep(5)
        except:
            Nowtext.configure(text='未檢測到彈窗廣告')
            time.sleep(5)

        # 轉跳已完全加載元素
        Nowtext.configure(text='滾動至分類表')
        x = windows.winfo_screenwidth()
        y = windows.winfo_screenheight()
        Nowtext.configure(text='當前畫面解析度為' + str(x) + '*' + str(y))
        if y < 800 and y > 600:
            browser.execute_script('window.scrollTo(0,window.screen.height+150)')
        elif y >= 800:
            browser.execute_script('window.scrollTo(0,window.screen.height-100)')
        elif y <= 600 and y >= 500:
            browser.execute_script('window.scrollTo(0,window.screen.height+300)')
        elif y < 500:
            messagebox.showerror('解析度過小','很抱歉該軟體不支援此解析度的螢幕 !')
            os._exit(0)
        time.sleep(5)
        
        # 取得資源列表
        Nowtext.configure(text='獲取分類表')
        classes = soup.find_all('a',{'class':'home-category-list__category-grid'})
        for c in classes:
           self.Cid.append(c.text)
        
        # 點擊選擇列表
        Nowtext.configure(text='取得輸入的分類')
        try:
            group = int(self.Cid.index(text)/2)+1
            idx = (self.Cid.index(text)%2)+1
        except:
            messagebox.showerror('錯誤','無法找到您輸入的分類id或是名稱，請檢查輸入分類是否錯誤，請開啟Classes.txt查看，或嘗試重新開啟軟體再次執行 !')
            browser.quit()
            os._exit(0)
        
        # 重複檢測分類可用性
        while True: 
            if(errorNum > 5):
                messagebox.showerror('錯誤','錯誤執行已超過5次，有可能是網路不穩定所致，或是所輸入的分類不存蝦皮，請檢查您輸入的分類於蝦皮上是否存在，並重新啟動軟體 !')
                os._exit(0)
            try:
                Nowtext.configure(text='進入分類中...')
                browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a['+str(idx)+']').click()
                
                """
                # 改成js強制按下按鈕
                data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a['+str(idx)+']')
                browser.execute_script("$(arguments[0]).click()", data)
                """
                time.sleep(5)
                break
            except NoSuchElementException as error:
                if '//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[1]/ul/li[11]/div/a['+str(idx)+']' in str(error.msg):
                    Nowtext.configure(text='進入分類失敗，重新進入中-(1)')
                    browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a').click()
                    
                    """
                    # 改成js強制按下按鈕
                    data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a')
                    browser.execute_script("$(arguments[0]).click()", data)
                    """

                    time.sleep(5)
                    errorNum += 1
                    break
                elif '//*[@id="main"]/div/div[2]/div[2]/div[2]/div[4]/div[1]/div/div/div[2]/div/div[1]/ul/li[11]/div/a['+str(idx)+']' in str(error.msg):
                    Nowtext.configure(text='進入分類失敗，重新進入中-(2)')
                    browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[4]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a').click()
                    
                    """
                    # 改成js強制按下按鈕
                    data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[4]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a')
                    browser.execute_script("$(arguments[0]).click()", data)
                    """

                    time.sleep(5)
                    errorNum += 1
                    break
                elif '//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a['+str(idx)+']' in str(error.msg):
                    Nowtext.configure(text='進入分類失敗，重新進入中-(3)')
                    browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[4]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a['+str(idx)+']').click()
                    
                    """
                    # 改成js強制按下按鈕
                    data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[4]/div[1]/div/div/div[2]/div/div[1]/ul/li['+str(group)+']/div/a['+str(idx)+']')
                    browser.execute_script("$(arguments[0]).click()", data)
                    """

                    time.sleep(5)
                    errorNum += 1
                    break
            except ElementClickInterceptedException :
                Nowtext.configure(text='網頁元素不完全，重新進入中-(4)')
                browser.refresh()
                x = windows.winfo_screenwidth()
                y = windows.winfo_screenheight()
                Nowtext.configure(text='當前畫面解析度為' + str(x) + '*' + str(y))
                if y < 800 and y > 600:
                    browser.execute_script('window.scrollTo(0,window.screen.height+150)')
                elif y >= 800:
                    browser.execute_script('window.scrollTo(0,window.screen.height-100)')
                elif y <= 600 and y >= 500:
                    browser.execute_script('window.scrollTo(0,window.screen.height+300)')
                elif y < 500:
                    messagebox.showerror('解析度過小','很抱歉該軟體不支援此解析度的螢幕 !')
                    os._exit(0)
                time.sleep(5)
                errorNum += 1
                pass
            except ElementNotInteractableException:
                Nowtext.configure(text='發現隱藏元素，處理元素中~')
                time.sleep(5)
                while True:
                    try:
                        x = windows.winfo_screenwidth()
                        y = windows.winfo_screenheight()
                        Nowtext.configure(text='當前畫面解析度為' + str(x) + '*' + str(y))
                        if y < 800 and y > 600:
                            browser.execute_script('window.scrollTo(0,window.screen.height+150)')
                        elif y >= 800:
                            browser.execute_script('window.scrollTo(0,window.screen.height-100)')
                        elif y <= 600 and y >= 500:
                            browser.execute_script('window.scrollTo(0,window.screen.height+300)')
                        elif y < 500:
                            messagebox.showerror('解析度過小','很抱歉該軟體不支援此解析度的螢幕 !')
                            os._exit(0)
                        time.sleep(5)
                        browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[3]').click()
                        
                        # 改成js強制按下按鈕
                        #data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div/div/div[2]/div/div[3]')
                        #browser.execute_script("$(arguments[0]).click()", data)
                        
                        time.sleep(5)
                        break
                    except:
                        pass
                    
                   
                

    # 進入產品列表中之一
    def enterProduce(self,num):
        id = 1
        n = 1
        pagenum = 0
        # 判斷起始頁
        if self.sp != '':
            url = browser.current_url
            pagenum = int(self.sp) -1
            browser.get(str(url).split('?')[0] + '?page=' + str(pagenum))
        
        # 判斷起始id
        if self.sid != '':
            id = int(self.sid)
        print('-------------------------------')
        while n <= num :
                # 嘗試進入商品頁
                try:
                    # 開始爬蟲
                    Nowtext.configure(text='進入商品中...')
                    self.scoll()
                    time.sleep(5)

                    # 變更介面文字
                    Nowtext.config(text="當前位於第"+str(pagenum+1)+"頁，第"+str(id)+"項")
                   
                    self.WaitforElement('//*[@id="main"]/div/div[3]/div/div[4]/div[2]/div/div[2]/div['+str(id)+']')
                    browser.find_element_by_xpath('//*[@id="main"]/div/div[3]/div/div[4]/div[2]/div/div[2]/div['+str(id)+']').click()

                    """
                    # 改成js強制按下按鈕
                    data = browser.find_element_by_xpath('//*[@id="main"]/div/div[3]/div/div[4]/div[2]/div/div[2]/div['+str(id)+']')
                    browser.execute_script("$(arguments[0]).click()", data)
                    """

                    print('Debug1')
                    time.sleep(5)
                except NoSuchElementException as error:
                    if '//*[@id="main"]/div/div[3]/div/div[4]/div[2]/div/div[2]/div['+str(id)+']' in str(error.msg):
                        print('Error1')
                        browser.refresh()
                        time.sleep(5)
                    else:
                        print('Error2')
                        continue
                except TimeoutException:
                    print('Error3')
                    id = 1
                    pagenum+=1
                    url = browser.current_url
                    if 'page' in str(url):
                        browser.get(str(url).split('?')[0] + '?page=' + str(pagenum))
                    else:
                        browser.get(str(url) + '?page=' + str(pagenum))
                    time.sleep(5)
                    continue
                # 按下聊聊按鈕
                try:
                    print('Debug2')
                    # 滾動畫面至聊聊按鈕
                    for kk in range(1,900,10):
                        browser.execute_script('var q=document.documentElement.scrollTop=' + str(kk))
                    print('滾動完成')
                    time.sleep(5)
                    
                    self.WaitforElement('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div[1]/div/div[3]/button')       
                    #browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div[1]/div/div[3]/button').click()
                    
                    
                    # 改成js強制按下按鈕
                    data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[1]/div[1]/div/div[3]/button')
                    browser.execute_script("$(arguments[0]).click()", data)
                    

                    time.sleep(5)
                except TimeoutException:
                    print('Error4')
                    try:
                        self.WaitforElement('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[2]/div[1]/div/div[3]/button')
                        #browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[2]/div[1]/div/div[3]/button').click()  
                        
                        
                        # 改成js強制按下按鈕
                        data = browser.find_element_by_xpath('//*[@id="main"]/div/div[2]/div[2]/div[2]/div[3]/div[2]/div[1]/div/div[3]/button')
                        browser.execute_script("$(arguments[0]).click()", data)                         
                        

                        time.sleep(5)
                    except TimeoutException:
                        print('Error4-1')
                        messagebox.showerror('網頁異常','目前發現網頁異常，請手動回到商品清單頁，再按下確定 !')
                        continue

                #點擊聊聊輸入框
                try:
                    print('Debug3')
                    # 啟動聊聊輸入文字
                    self.conversation(self.context)
                    print('本次已完成商品編號:' + str(id))
                    print('本次已完成控制次數:' + str(n))
                    id+=1
                    n+=1
                    browser.back()
                    time.sleep(5)
                except :
                    print('Error5')
                    browser.back()
                    time.sleep(5)
                    continue
                print('-------------------------------')
        browser.quit()
        messagebox.showinfo('通知','您的程序已完成，系統即將關閉')
        os._exit(0)

    # 將文字貼上聊聊並發送
    def conversation(self,content):
        print('1')
        Nowtext.configure(text='轉貼文字中...')
        print(2)
        # 拆解聊聊輸入文字 
        data1 = self.Split(content)
        for d1 in data1 :
            browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div/textarea').send_keys(d1)
            browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div/textarea').send_keys(Keys.SHIFT + Keys.ENTER)
        time.sleep(2)
        print('3')
        data2 = self.Split(loginWindows.getFileData('defult_text'))
        for d2 in data2:
            browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div/textarea').send_keys(d2)
            browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div/textarea').send_keys(Keys.SHIFT + Keys.ENTER)
        time.sleep(2)
        print(4)

        # 送出訊息
        Nowtext.configure(text='送出文字')
        browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[2]/div[1]/div[2]/div[2]/div/div/div[1]/div/textarea').send_keys('\n')
        time.sleep(2)
        print(5)

        # 縮小聊聊視窗
        self.WaitforElement('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[1]/div[2]/div[2]/div')
        browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[1]/div[2]/div[2]/div').click()
        
        """
        # 改用js強制按下按鈕
        data = browser.find_element_by_xpath('//*[@id="shopee-mini-chat-embedded"]/div[1]/div[1]/div[2]/div[2]/div')
        browser.execute_script("$(arguments[0]).click()", data)
        """

        time.sleep(5)
    
    # 滾動控制
    def scoll(self):
        browser.execute_script("var q=document.documentElement.scrollTop=0")
        for i in range(0,10000,10):
            browser.execute_script('var q=document.documentElement.scrollTop=' + str(i))
    
    # 等待元素
    def WaitforElement(self,element):
        wait.until(EC.presence_of_element_located((By.XPATH,element)))
    
    # 拆解輸入字串
    def Split(self,line):
        data = line.split('\n')
        if '' in data:
            data.remove('')
        return data
# ===================================================================
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# 密碼視窗物件
# ===================================================================
class loginWindows():
    def __init__(self,name):
        global windows1
        windows1 = tk.Tk()
        windows1.title(name)
        windows1.geometry('480x280')


        # 標題
        global title
        title = tk.Label(windows1,text='管理者模式',font=('Arial', 22))
        title.place(relx=0.32,rely=0.01)

        # 說明文字(帳號)
        global text1
        text1 = tk.Label(windows1,text='帳號:',font=('Arial', 18))
        text1.place(relx=0.1,rely=0.3)

        # 說明文字(密碼))
        global text2
        text2 = tk.Label(windows1,text='密碼:',font=('Arial', 18))
        text2.place(relx=0.1,rely=0.53)

        # 建立按紐(確認))
        global btn1
        btn1 = tk.Button(windows1,font=('Arial', 18),text='登入',command=self.loginBtn)
        btn1.place(relx=0.25,rely=0.73)

        # 建立按紐(取消))
        global btn2
        btn2 = tk.Button(windows1,font=('Arial', 18),text='取消',command=self.canelBtn)
        btn2.place(relx=0.62,rely=0.73)

        # 輸入文字(帳號)))
        global input1
        input1 = tk.Entry(windows1,font=('Arial', 18))
        input1.place(relx=0.3,rely=0.3)

        # 輸入文字(密碼))
        global input2
        input2 = tk.Entry(windows1,show='*',font=('Arial', 18))
        input2.place(relx=0.3,rely=0.53)


        windows1.mainloop()

    # 建立按紐觸發事件
    # ---------------------------------------------------------------
    def canelBtn(self):
        windows1.destroy()

    def loginBtn(self):
        user = input1.get()
        passwd = input2.get()
        if user == 'root' and passwd == 'zzzz037921661': # zzzz037921661
            pass
        else:
            messagebox.showerror('錯誤','帳號或密碼錯誤 !')
            return
        windows1.destroy()
        a = Setup('設定')

    def readAll():
        try:
            with open('C:\\Program Files (x86)\\ShopeeConteoller3.json','r') as f:
                data = json.load(f)
                return data
        except:
            pass

    def setFileData(content):
        try:
            with open('C:\\Program Files (x86)\\ShopeeConteoller3.json','w') as f:
                data = json.dumps(content)
                f.write(data)
        except:
            messagebox.showerror('權限不足','請以裝置管理員執行本軟體 !!')
            os._exit(0)
    
    def getFileData(key):
        try:
            with open('C:\\Program Files (x86)\\ShopeeConteoller3.json','r') as f:
                data = json.load(f)
                return data[key]
        except:
            pass

# ===================================================================
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# 設定視窗物件
# ===================================================================
class Setup():
    def __init__(self,name):
        global windows2
        windows2 = tk.Tk()
        windows2.title(name)
        windows2.geometry('450x500')

        # 標題
        global title
        title = tk.Label(windows2,text='設定',font=('Arial', 22))
        title.place(relx=0.45,rely=0.01)

        # 說明文字(版本)
        global text0
        text0 = tk.Label(windows2,text='軟體版本',font=('Arial', 18))
        text0.place(relx=0.1,rely=0.1)

        # 說明文字(說明)
        global text1
        text1 = tk.Label(windows2,text='正式版',font=('Arial', 18))
        text1.place(relx=0.5,rely=0.1)

        # 說明文字(時間)
        global text4
        text4 = tk.Label(windows2,text='時間授權',font=('Arial', 18))
        text4.place(relx=0.1,rely=0.2)

        # 說明文字(說明)
        global text5
        text5 = tk.Label(windows2,text="3天版",font=('Arial', 18))
        text5.place(relx=0.5,rely=0.2)

        # 說明文字(說明)
        global text_Duf
        text_Duf = tk.Label(windows2,text="聊聊預設文字",font=('Arial', 18))
        text_Duf.place(relx=0.4,rely=0.4)

        # 輸入文字
        global input_Duf
        input_Duf = tk.Text(windows2,font=('Arial', 18))
        input_Duf.place(relx=0.25,rely=0.5,width= 265,height=150)
        try:
            input_Duf.insert('0.0',loginWindows.getFileData('defult_text'))
        except:
            pass

        
        # 建立按紐(儲存)
        global btn1
        btn1 = tk.Button(windows2,font=('Arial', 18),text='儲存狀態',command=self.btn1Control)
        btn1.place(relx=0.2,rely=0.85)

        # 建立按紐(撤銷授權)
        global btn_delete
        btn_delete = tk.Button(windows2,font=('Arial', 18),text='撤銷授權',command=self.btn_delete,background="red")
        btn_delete.place(relx=0.6,rely=0.85)
        
        # 建立按紐(時間碼)
        global btn3
        btn3 = tk.Button(windows2,font=('Arial', 12),text='輸入序號',command=self.btn2Control)
        btn3.place(relx=0.7,rely=0.2)


        windows2.mainloop()

    def btn1Control(self):
        row_data['defult_text'] = input_Duf.get('0.0','end')
        loginWindows.setFileData(row_data)
        print(row_data)
        windows2.destroy()

    def btn2Control(self):
        d = dialogWindows('序號輸入框')

    def btn_delete(self):
        ans = messagebox.askokcancel('警告','當您撤銷授權後，必須重新授權才可使用本軟體，是否繼續執行?')
        if ans == True:
            try:
                os.remove('C:\\Program Files (x86)\\ShopeeConteoller3.json')
                messagebox.showinfo('撤銷完成','已將授權徹底移除!!!')
                os._exit(0)
            except PermissionError:
                messagebox.showerror('權限不足','請以裝置管理員執行本軟體 !!')
                os._exit(0)
            except:
                messagebox.showerror('授權檔案無效','本軟體尚未授權，無須撤銷 !!')
        else:
            pass

# ===================================================================
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# 彈跳視窗物件
# ===================================================================

class dialogWindows():
    def __init__(self,name):
        global windows3
        windows3 = tk.Tk()
        windows3.title(name)
        windows3.geometry('300x150')

        # 輸入文字
        global inputa
        inputa = tk.Entry(windows3,font=('Arial', 18),show="*",takefocus=True)
        inputa.place(relx=0.05,rely=0.2)
        
        # 輸入(確認)
        global btna
        btna = tk.Button(windows3,font=('Arial', 18),text='確認',command=self.btnControl)
        btna.place(relx=0.4,rely=0.5)
        windows3.mainloop()

    def btnControl(self):  
        s =  dialogWindows.tkeyCheck(inputa.get())
        if s:
            row_data['tkeyStatus'] = True
            row_data['expiryDate'] = dialogWindows.CalculateDate()
        else:
            row_data['tkeyStatus'] = False
            row_data['expiryDate'] = dialogWindows.getNowTime()
        windows3.destroy()
    
    def tkeyCheck(key):
        nowkey = int(str(datetime.datetime.now())[5:7])
        if key == str((nowkey*123)+1):
            return True
        else:
            return False

    def checktime():
        while True:
            try:
                if (datetime.datetime.strptime(loginWindows.getFileData('expiryDate'),'%Y-%m-%d')-datetime.datetime.strptime(dialogWindows.getNowTime(),'%Y-%m-%d')).days <= 0:
                    row_data['tkeyStatus'] = False
                    loginWindows.setFileData(row_data)
                    break
                else:
                    break
            except:
                row_data['expiryDate'] = dialogWindows.getNowTime()
                loginWindows.setFileData(row_data)

    def getNowTime():
        try:
            ntp1 = ntplib.NTPClient()
            respone1 = ntp1.request('pool.ntp.org')
            date1 = time.strftime("%Y-%m-%d",time.localtime(respone1.tx_time))
            return str(date1)
        except:
            messagebox.showerror('網路時間驗證失敗','網路時間服務器驗證失敗，按下確認號請稍等5秒，再繼續您的動作')
            os._exit(0)
    
    def CalculateDate():
        try:
            ntp = ntplib.NTPClient()
            respone = ntp.request('pool.ntp.org')
            date = time.strftime('%Y-%m-%d',time.localtime(respone.tx_time))
            date = str(datetime.datetime.strptime(date,'%Y-%m-%d') + datetime.timedelta(days=3))
            print(date[:10])
            return date[:10]
        except:
             messagebox.showerror('網路時間驗證失敗','網路時間服務器驗證失敗，按下確認號請稍等5秒，再繼續您的動作')
             os._exit(0)


# ===================================================================
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# ///////////////////////////////////////////////////////////////////
# # 視窗物件
# ===================================================================
class frame():
    num = 0
    # 建構子
    # -------------------------------------------------------------------
    def __init__(self,name):
        # 檢查權限
        try:
            if not loginWindows.getFileData('tkeyStatus') or loginWindows.getFileData('expiryDate') == '':
                with open('C:\\Program Files (x86)\\ShopeeConteoller3.json','w') as f:
                    mdata ={'expiryDate': dialogWindows.getNowTime(),'tkeyStatus':False,'defult_text':'','classes':'','page':'','num':'','content':'','times':''}
                    data = json.dumps(mdata)
                    f.write(data)
        except PermissionError:
            messagebox.showerror('權限不足','請以裝置管理員執行本軟體 !!')
            os._exit(0)
        
        global row_data
        row_data = loginWindows.readAll()

        global windows
        windows = tk.Tk()
        windows.title(name)
        windows.geometry('500x500')

        dialogWindows.checktime()

        # 元件建立區
        # -------------------------------------------------------------------

        # 標題
        self.title = tk.Label(windows,text='蝦皮聊聊控制器',font=('Arial', 22))
        self.title.place(relx=0.32,rely=0.01)

        # 說明文字(分類選擇)
        self.text1 = tk.Label(windows,text='請輸入編號:',font=('Arial', 18))
        self.text1.place(relx=0.01,rely=0.1)

        # 說明文字(範圍)
        self.text2 = tk.Label(windows,text='請輸入範圍:',font=('Arial', 18))
        self.text2.place(relx=0.01,rely=0.2)

        # 說明文字(第)
        self.text3 = tk.Label(windows,text='第',font=('Arial', 18))
        self.text3.place(relx=0.38,rely=0.2)

        # 輸入文字(分類))
        self.input4 = tk.Entry(windows,font=('Arial', 18),width=3)
        self.input4.place(relx=0.43,rely=0.2)
        try:
            self.input4.insert('0',loginWindows.getFileData('page'))
        except:
            pass

        # 說明文字(頁)
        self.text4 = tk.Label(windows,text='頁,第',font=('Arial', 18))
        self.text4.place(relx=0.48,rely=0.2)

        # 輸入文字(分類))
        self.input5 = tk.Entry(windows,font=('Arial', 18),width=3)
        self.input5.place(relx=0.60,rely=0.2)
        try:
            self.input5.insert('0',loginWindows.getFileData('num'))
        except:
            pass

        # 說明文字(項)
        self.text5 = tk.Label(windows,text='項',font=('Arial', 18))
        self.text5.place(relx=0.65,rely=0.2)

        # 說明文字(聊聊視窗))
        self.text6 = tk.Label(windows,text='請輸入聊聊文字:',font=('Arial', 18))
        self.text6.place(relx=0.01,rely=0.3)

        # 說明文字(筆數))
        self.text7 = tk.Label(windows,text='請輸入執行筆數:',font=('Arial', 18))
        self.text7.place(relx=0.01,rely=0.65)

        # 說明文字(當前筆數))
        global Nowtext
        Nowtext = tk.Label(windows,text='',font=('Arial', 14))
        Nowtext.place(relx=0.4,rely=0.8)

        # 輸入文字(分類))
        self.input1 = tk.Entry(windows,font=('Arial', 18))
        self.input1.place(relx=0.4,rely=0.1)
        try:
            self.input1.insert('0',loginWindows.getFileData('classes'))
        except:
            pass

        # 輸入文字(聊聊視窗))
        self.input2 = tk.Text(windows,font=('Arial', 18))
        self.input2.place(relx=0.4,rely=0.3,width= 265,height=150)
        try:
            self.input2.insert('0.0',loginWindows.getFileData('content'))
        except:
            pass

        # 輸入文字(筆數))
        self.input3 = tk.Entry(windows,font=('Arial', 18),width=20)
        self.input3.place(relx=0.4,rely=0.65)
        try:
            self.input3.insert('0',loginWindows.getFileData('times'))
        except:
            pass

        # 說明文字(時間)
        self.text8 = tk.Label(windows,text='有效時間:',font=('Arial', 18))
        self.text8.place(relx=0.01,rely=0.8)

        # 說明文字(說明)
        self.text9 = tk.Label(windows,text= str((datetime.datetime.strptime(loginWindows.getFileData('expiryDate'),'%Y-%m-%d')  - datetime.datetime.strptime(dialogWindows.getNowTime(),'%Y-%m-%d')).days) + ' 天',font=('Arial', 18))  
        self.text9.place(relx=0.25,rely=0.8)

        # 建立按紐(管理人員))
        self.btn1 = tk.Button(windows,font=('Arial', 18),text='管理人員',command=self.btn1Event)
        self.btn1.place(relx=0.11,rely=0.9)

        # 建立按紐(鎖定資料))
        self.btn2 = tk.Button(windows,font=('Arial', 18),text='鎖定資料',command=self.btn2Event)
        self.btn2.place(relx=0.42,rely=0.9)

        # 建立按紐(開始任務))
        self.btn3 = tk.Button(windows,font=('Arial', 18),text='開始任務',command=self.btn3Event,state='disable')
        self.btn3.place(relx=0.72,rely=0.9)

        # 檢查授權
        self.checkLicense()

        windows.protocol("WM_DELETE_WINDOW",self.delteBrowser)
        windows.mainloop()
    
     # 建立按紐觸發事件
    # -------------------------------------------------------------------
    def btn1Event(self):
        f2 = loginWindows('管理者模式驗證')
        
    
    def btn2Event(self):
        if((self.num % 2) == 0 ):
            self.input1.configure(state='disable')
            self.input2.configure(state='disable')
            self.input3.configure(state='disable')
            self.input4.configure(state='disable')
            self.input5.configure(state='disable')
            self.btn3.configure(state='normal')
            self.btn2.configure(text='解除資料')
        else:
            self.input1.configure(state='normal')
            self.input2.configure(state='normal')
            self.input3.configure(state='normal')
            self.input4.configure(state='normal')
            self.input5.configure(state='normal')
            self.btn3.configure(state='disable')
            self.btn2.configure(text='鎖定資料')
        self.num += 1
    
    def btn3Event(self):
        row_data['classes'] = self.input1.get()
        row_data['page'] = self.input4.get()
        row_data['num'] = self.input5.get()
        row_data['content'] = self.input2.get('0.0','end')
        row_data['times'] = self.input3.get()
        loginWindows.setFileData(row_data)
        self.btn1.configure(state='disable')
        self.btn2.configure(state='disable')
        self.btn3.configure(state='disable')
        global td
        td = threading.Thread(target=self.StartBrowser)
        td.start()
    
    def StartBrowser(self):
        s = Shopee(self.input1.get(),self.input2.get('0.0','end'),self.input3.get(),self.input4.get(),self.input5.get())
        

    # 特殊控制函數
    #----------------------------------------------------------------
    def checkLicense(self):
        if loginWindows.getFileData('tkeyStatus'):
            self.input1.configure(state='normal')
            self.input2.configure(state='normal')
            self.input3.configure(state='normal')
            self.input4.configure(state='normal')
            self.input5.configure(state='normal')
            self.btn3.configure(state='disable')
            self.btn2.configure(text='鎖定資料')
        else:
            self.input1.configure(state='disable')
            self.input2.configure(state='disable')
            self.input3.configure(state='disable')
            self.input4.configure(state='disable')
            self.input5.configure(state='disable')
            self.btn3.configure(state='disable')
            self.btn2.configure(state='disable')
    
    def delteBrowser(self):
        windows.destroy()
        try:
            browser.quit()
        except:
            pass
        
# -------------------------------------------------------------------

# 主程式運行區
# --------------------------------------------------------------------------
if __name__ == '__main__':
    f = frame('ShopeeController')