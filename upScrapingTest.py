# -*- coding: utf-8 -*-
import re
import time
import xlwt
import xlrd
import mechanize
from bs4 import BeautifulSoup

#設定
urllogin = 'https://fesns.com/?m=pc&a=page_o_login'
urlformer = 'http://fesns.com/?m=fe&a=page_war_record&target_unit_id=3&target_seq=583'
urltarget = 'http://fesns.com/?m=fe&a=page_war_record&target_unit_id=3&target_seq=583447'
username = '****'
password = '****'
#0,1,2のどれか
charactor_no = 0

#ログイン
br = mechanize.Browser()
br.open(urllogin)
br.select_form(name='login')
br['username'] = username
br['password'] = password
br.submit()
br.select_form(nr=charactor_no)
br.submit()

#Excelの準備
book = xlwt.Workbook()
sheet = book.add_sheet("Sheets1")


#------------メイン処理----------------
for urlnum in range(1000):

    #連番でURLを拾ってくる
    urltarget = urlformer + str(urlnum).rjust(3,"0")
    print urltarget
    br.open(urltarget)
    soup = BeautifulSoup(br.response().read())

    #初期化
    output_list = []
    output_dict = {}

    #タイトル
    wartitlediv = soup.find('div', attrs={'class':'WarTitle'})
    #例外処理
    if wartitlediv is None:
        print 'no such warnumber'
        continue
    wartitle = wartitlediv.get_text().strip()
    print wartitle
    output_list.append(wartitle)

    #参加人数
    partinum_div = soup.find('td', text='参加人数')
    partinum = partinum_div.find_next_sibling('td').get_text()
    print partinum
    output_list.append(partinum)

    #参加者少数ならスキップ
    partinum_flag = False
    for i in re.findall(r'[0-9]+', partinum):
        if int(i)<48:
            print 'dameeee'
            partinum_flag = True
    if partinum_flag:
        continue

    #ゲージ
    gaugediv = soup.find_all('div', attrs={'class':'gauge'})
    for tag in gaugediv:
        if len(tag.get_text().strip())==0:
            continue
        gauge = tag.get_text().strip()
        m = re.search(r'([0-9]+)', gauge).group(0)
        output_list.append(int(m))

    #職別数
    syokudiv = soup.find_all('div', attrs={'class':'WarMember Heading partsHeading'})
    for tag in syokudiv:
        syoku = tag.get_text().strip()
        if syoku.isalpha():
            continue
        templist = syoku.rstrip(r')').split(r' (')
        if output_dict.has_key(templist[0]):
            output_dict[templist[0] + 'a'] = int(templist[1])
        else:
            output_dict[templist[0]] = int(templist[1])

    #Excelに出力
    rownum = int(urlnum)
    title_row = sheet.row(rownum*4+1)
    defend_row = sheet.row(rownum*4+2)
    attack_row = sheet.row(rownum*4+3)

    for i in range(4):
        title_row.write(i,output_list[i])
        
    defend_row.write(0,output_dict.get('Warrior', 0))
    defend_row.write(1,output_dict.get('Scout', 0))
    defend_row.write(2,output_dict.get('Sorcerer', 0))
    defend_row.write(3,output_dict.get('Fencer', 0)) 
    defend_row.write(4,output_dict.get('Cestus', 0))
                     
    attack_row.write(0,output_dict.get('Warriora', 0))
    attack_row.write(1,output_dict.get('Scouta', 0))
    attack_row.write(2,output_dict.get('Sorcerera', 0))
    attack_row.write(3,output_dict.get('Fencera', 0)) 
    attack_row.write(4,output_dict.get('Cestusa', 0))
  
    #過負荷配慮    
    time.sleep(1)

#Excelに保存して終わり
book.save('sample.xls')
