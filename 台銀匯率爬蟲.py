import requests
from bs4 import BeautifulSoup
import pandas as pd
import datetime 
from tkinter import *
from tkinter import ttk 
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties 
from PIL import Image, ImageTk

#讀取歷史匯率的網址
url_m="https://rate.bot.com.tw/xrt/flcsv/0/L3M/"

plt.rc('font', family='Microsoft JhengHei')

def bt_reload():    #按下按鈕的背景變化
    bt_USD.config(bg="white")
    bt_HKD.config(bg="white")
    bt_GBP.config(bg="white")
    bt_JPY.config(bg="white")
    bt_CNY.config(bg="white")
    

def USD():
    bt_reload()
    bt_USD.config(bg="lightblue")
    Plot(url_m+"USD")   #將歷史匯率網址傳至plot圖表函式
    
    
def HKD():
    bt_reload()
    bt_HKD.config(bg="lightblue")
    Plot(url_m+"HKD")
    
def GBP():
    bt_reload()
    bt_GBP.config(bg="lightblue")
    Plot(url_m+"GBP")
    
def JPY():
    bt_reload()
    bt_JPY.config(bg="lightblue")
    Plot(url_m+"JPY")
    
def CNY():
    bt_reload()
    bt_CNY.config(bg="lightblue")
    Plot(url_m+"CNY")

def Plot(url):  # 製作歷史匯率圖表
    url = url   # 讀入歷史匯率資料網址
    rate = requests.get(url)   # 爬取網址內容
    rate.encoding = 'utf-8'    # 調整回應訊息編碼為utf-8，避免編碼不同造成亂碼
    rt = rate.text             # 以文字模式讀取內容
    rts = rt.split('\n')       # 使用「換行」將內容拆分成串列
    for i in range(1):         # 分別讀取資料
        date=[]     # 日期
        plt_b_i=[]  # 現金買入
        plt_b_o=[]  # 現金賣出
        plt_s_i=[]  # 即期買入
        plt_s_o=[]  # 即期賣出
        n=0         # 資料筆數
        for i in rts:
            if (i!="\n")&(n<42):        # 避開最後一行的空白行
                a = i.split(',')        # 每個項目用逗號拆分成子串列
                date.append(a[0])
                plt_b_i.append(a[3])
                plt_b_o.append(a[13])
                plt_s_i.append(a[4])
                plt_s_o.append(a[14])
                n+=1
            else:
                break
    # 隱藏index=0,因為不重要,有中文字'現金''即期'
    del date[0],plt_b_i[0],plt_b_o[0],plt_s_i[0],plt_s_o[0]
    date.reverse()
    plt_b_i.reverse()
    plt_b_o.reverse()
    plt_s_i.reverse()
    plt_s_o.reverse()

    # 將list中的str轉float
    plt_b_i=[float(x) for x in plt_b_i]
    plt_b_o=[float(x) for x in plt_b_o]
    plt_s_i=[float(x) for x in plt_s_i]
    plt_s_o=[float(x) for x in plt_s_o]

    fig=plt.figure()
    ax1=fig.add_subplot(2,1,1)  # 現金買入/賣出匯率
    ax2=fig.add_subplot(2,1,2)  # 即期買入/賣出匯率

    A,=ax1.plot(date,plt_b_i)
    B,=ax1.plot(date,plt_b_o)
    plt.setp(ax1.get_xticklabels(), # 設定x軸刻度文字
              fontsize=8,
              rotation='vertical')  
    ax1.set_title("現金匯率")
    ax1.set_xlabel("日期")
    ax1.set_ylabel("匯率")
    
    C,=ax2.plot(date,plt_s_i)
    D,=ax2.plot(date,plt_s_o)
    plt.setp(ax2.get_xticklabels(),
              fontsize=8,
              rotation='vertical')
    ax2.set_title("即期匯率")
    ax2.set_xlabel("日期")
    ax2.set_ylabel("匯率")
    
    plt.tight_layout()      # 避免子圖間重疊
    plt.savefig("plt.png")  # 存檔
    
    global tk_img #ImageTk.PhotoImage()放函式, img會被回收, 所以使用global
    tk_img =ImageTk.PhotoImage(file="plt.png")  # 轉檔以顯示在Tkinter視窗中

    # 開啟圖並縮小到合適大小
    img = Image.open("plt.png")
    img = img.resize((432, 288))  # 自訂大小
    tk_img = ImageTk.PhotoImage(img)
    
    label = Label(window, image=tk_img) # 顯示圖表
    label.pack()
    label.place(x=250,y=45)

df=pd.DataFrame()

# 台灣銀行即時匯率網址
url="https://rate.bot.com.tw/xrt?Lang=zh-TW"

# 使用requests物件的get方法把網頁抓下來
response=requests.get(url)

# 把HTML原始碼存入data
html_doc=response.text

# 使用lxml解析data
soup=BeautifulSoup(html_doc,"lxml")

# 找到匯率表格
rate_table=soup.find("table").find("tbody")

# 記錄匯率表格的每一行(整行)
rate_table_rows=rate_table.find_all("tr")

for index,row in enumerate(rate_table_rows,1):
    # 設定用來存放資料
    a=[]
    buyin=[]
    buyout=[]
    spotbuy=[]
    spotsold=[]
    
    # 把每行資料拆開來紀錄
    columns=row.find_all("td")
    
    for c in columns:
        # 讀取幣別
        if c.attrs["data-table"]=="幣別":
            # 初始化
            last_div=None
        
            # 在if條件下, 把所有tag "div"記錄下來
            divs=c.find_all("div")
            
            # 用來找最後一個div
            for last_div in divs:pass
            
            # 指定條件下才將資料存入dataframe
            # last_div.string.strip() ->最後一個div中的字串去除前後空格（strip）
            if "美金"  in last_div.string.strip():
                # 將資料先儲存在變數
                a=last_div.string.strip()
            elif "日圓"  in last_div.string.strip():
                a=last_div.string.strip()
            elif "港幣"  in last_div.string.strip():
                a=last_div.string.strip()
            elif "英鎊"  in last_div.string.strip():
                a=last_div.string.strip()
            elif "人民幣"  in last_div.string.strip():
                a=last_div.string.strip()
            else: a=0
            
        # 讀取即時匯率資料
        if c.getText().find("查詢")!=0 and str(c).find("print_width")>0 and a!=0:
            if c.attrs["data-table"]=="本行現金買入" :
                # 存入匯率資訊
                if c.getText().strip()!='-':
                    buyin=c.getText().strip()
                    buyin= float(buyin)
            if c.attrs["data-table"]=="本行現金賣出" :
                if c.getText().strip()!='-':
                    buyout=c.getText().strip()
                    buyout= float(buyout)
            if c.attrs["data-table"]=="本行即期買入" :
                if c.getText().strip()!='-':
                    spotbuy=c.getText().strip()
                    spotbuy= float(spotbuy)
            if c.attrs["data-table"]=="本行即期賣出" :
                if c.getText().strip()!='-':
                    spotsold=c.getText().strip()
                    spotsold= float(spotsold)
                    # 將資料加入dataframe中
                    df1 = pd.DataFrame({'幣別':a, '本行現金買入':buyin, '本行現金賣出':buyout, '本行即期買入':spotbuy, '本行即期賣出':spotsold}, index=[len(df)])
                    df = pd.concat([df, df1], ignore_index=True)
    
# 檔名加入時間
now = datetime.datetime.now()
currentDateTime = now.strftime("%Y-%m-%d %H %M %S")

# 轉換為csv檔
df.to_csv(f"test {currentDateTime} .csv",encoding="utf_8_sig",index=False)


# 視窗建立
window=Tk()

window.title("台灣銀行五種貨幣匯率")
window.geometry("800x600")
window.config(bg="gray")

# 以tkinter Treeview讓dataframe能夠在視窗顯示
# 表格設定
tree=ttk.Treeview(window,show="headings",columns=("coin","b_i","b_o","s_i","s_o"))
tree.column("coin",width=80,anchor="center")
tree.column("b_i",width=80,anchor="center")
tree.column("b_o",width=80,anchor="center")
tree.column("s_i",width=80,anchor="center")
tree.column("s_o",width=80,anchor="center")

# 表格標題列
tree.heading("coin",text="幣別")
tree.heading("b_i",text="現金買入")
tree.heading("b_o",text="現金賣出")
tree.heading("s_i",text="即期買入")
tree.heading("s_o",text="即期賣出")

# 資料轉換
val_coin=df["幣別"]
val_b_i=df["本行現金買入"]
val_b_o=df["本行現金賣出"]
val_s_i=df["本行即期買入"]
val_s_o=df["本行即期賣出"]

# 資料輸入
tree.insert("",0,text="line1",values=(val_coin[0],val_b_i[0],val_b_o[0],val_s_i[0],val_s_o[0])) #插入数据
tree.insert("",1,text="line1",values=(val_coin[1],val_b_i[1],val_b_o[1],val_s_i[1],val_s_o[1])) 
tree.insert("",2,text="line1",values=(val_coin[2],val_b_i[2],val_b_o[2],val_s_i[2],val_s_o[2])) 
tree.insert("",3,text="line1",values=(val_coin[3],val_b_i[3],val_b_o[3],val_s_i[3],val_s_o[3])) 
tree.insert("",4,text="line1",values=(val_coin[4],val_b_i[4],val_b_o[4],val_s_i[4],val_s_o[4])) 
tree.pack()

# 按鈕設定
bt_USD=Button(window,text="美金(USD)",width=10,height=2,font=5,command=USD)
bt_HKD=Button(window,text="港幣(HKD)",width=10,height=2,font=5,command=HKD)
bt_GBP=Button(window,text="英鎊(GBP)",width=10,height=2,font=5,command=GBP)
bt_JPY=Button(window,text="日圓(JPY)",width=10,height=2,font=5,command=JPY)
bt_CNY=Button(window,text="人民幣(CNY)",width=10,height=2,font=5,command=CNY)

# 排版
bt_USD.place(x=25,y=45)
bt_HKD.place(x=25,y=105)
bt_GBP.place(x=25,y=165)
bt_JPY.place(x=25,y=225)
bt_CNY.place(x=25,y=285)
tree.place(x=250,y=45)

window.mainloop()
