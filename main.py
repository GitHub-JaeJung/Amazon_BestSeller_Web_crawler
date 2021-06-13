import os
import sys
import time
import urllib.request
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver

query_txt = '아마존'
query_url = 'https://www.amazon.com/bestsellers?ld=NSGoogle'

selc = input('''
1. Amazon Devices & Accessories     2. Amazon Launchpad                3. Appliances
4. Apps & Games                     5. Arts, Crafts & Sewing           6. Audible Books & Originals
7. Automotive                       8. Baby                            9. Beauty & Personal Care      
10. Books                           11. CDs & Vinyl                    12. Camera & Photo Products            
13. Cell Phones & Accessories       14. Clothing, Shoes & Jewelry      15. Collectible Currencies       
16. Computers & Accessories         17. Digital Educational Resources  18. Digital Music                
19. Electronics                     20. Entertainment Collectibles     21. Gift Cards          
22. Grocery & Gourmet Food          23. Handmade Products              24. Health & Household             
25. Home & Kitchen                  26. Industrial & Scientific        27. Kindle Store           
28. Kitchen & Dining                29. Magazine Subscriptions         30. Movies & TV        
31. Musical Instruments             32. Office Products                33. Patio, Lawn & Garden               
34. Pet Supplies                    35. Software                       36. Sports & Outdoors                   
37. Sports Collectibles               38. Tools & Home Improvement            39. Toys & Games   
40. Video Games                   

1.위 분야 중에서 자료를 수집할 분야의 번호를  선택하세요: ''')

cnt = int(input('2.해당 분야에서 크롤링 할 건수는 몇건입니까?(1-100 건 사이 입력): '))

f_dir = input("3.파일을 저장할 폴더명만 쓰세요(예:c:\\temp\\):")
print("\n")

if selc == '1':
    selc_name = 'Amazon Devices & Accessories'
elif selc == '2':
    selc_name = 'Amazon Launchpad'
elif selc == '3':
    selc_name = 'Appliances'
elif selc == '4':
    selc_name = 'Apps & Games'
elif selc == '5':
    selc_name = 'Arts, Crafts & Sewing'
elif selc == '6':
    selc_name = 'Audible Books & Originals'
elif selc == '7':
    selc_name = 'Automotive'
elif selc == '8':
    selc_name = 'Baby'
elif selc == '9':
    selc_name = 'Beauty & Personal Care'
elif selc == '10':
    selc_name = 'Books'
elif selc == '11':
    selc_name = 'CDs & Vinyl'
elif selc == '12':
    selc_name = 'Camera & Photo Products'
elif selc == '13':
    selc_name = 'Cell Phones & Accessories'
elif selc == '14':
    selc_name = 'Clothing, Shoes & Jewelry '
elif selc == '15':
    selc_name = 'Collectible Currencies'
elif selc == '16':
    selc_name = 'Computers & Accessories'
elif selc == '17':
    selc_name = 'Digital Educational Resources'
elif selc == '18':
    selc_name = 'Digital Music'
elif selc == '19':
    selc_name = 'Electronics'
elif selc == '20':
    selc_name = 'Entertainment Collectibles'
elif selc == '21':
    selc_name = 'Gift Cards'
elif selc == '22':
    selc_name = 'Grocery & Gourmet Food'
elif selc == '23':
    selc_name = 'Handmade Products'
elif selc == '24':
    selc_name = 'Health & Household'
elif selc == '25':
    selc_name = 'Home & Kitchen'
elif selc == '26':
    selc_name = 'Industrial & Scientific'
elif selc == '27':
    selc_name = 'Kindle Store'
elif selc == '28':
    selc_name = 'Kitchen & Dining'
elif selc == '29':
    selc_name = 'Magazine Subscriptions'
elif selc == '30':
    selc_name = 'Movies & TV'
elif selc == '31':
    selc_name = 'Musical Instruments'
elif selc == '32':
    selc_name = 'Office Products'
elif selc == '33':
    selc_name = 'Patio, Lawn & Garden'
elif selc == '34':
    selc_name = 'Pet Supplies'
elif selc == '35':
    selc_name = 'Software'
elif selc == '36':
    selc_name = 'Sports & Outdoors'
elif selc == '37':
    selc_name = 'Sports Collectibles'
elif selc == '38':
    selc_name = 'Tools & Home Improvement'
elif selc == '39':
    selc_name = 'Toys & Games'
elif selc == '40':
    selc_name = 'Video Games'

now = time.localtime()
s = '%04d-%02d-%02d-%02d-%02d-%02d' % (now.tm_year, now.tm_mon, now.tm_mday, now.tm_hour, now.tm_min, now.tm_sec)

os.makedirs(f_dir + s +'-'+query_txt+'-'+selc_name)
os.chdir(f_dir + s + '-' + query_txt + '-' + selc_name)

ff_dir = f_dir + s + '-' + query_txt + '-' + selc_name
f_name = f_dir + s + '-' + query_txt + '-' + selc_name + '\\' + s + '-' + query_txt + '-' + selc_name + '.txt'
fc_name = f_dir + s + '-' + query_txt + '-' + selc_name + '\\' + s + '-' + query_txt + '-' + selc_name + '.csv'
fx_name = f_dir + s + '-' + query_txt + '-' + selc_name + '\\' + s + '-' + query_txt + '-' + selc_name + '.xls'

s_time = time.time()

path = "c:/temp/chromedriver_240/chromedriver.exe"
driver = webdriver.Chrome(path)

driver.get(query_url)
time.sleep(5)

# 분야별 더보기 버튼을 눌러 페이지를 엽니다
if selc == '1':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[1]/a""").click()
elif selc == '2':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[2]/a""").click()
elif selc == '3':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[3]/a""").click()
elif selc == '4':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[4]/a""").click()
elif selc == '5':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[5]/a""").click()
elif selc == '6':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[6]/a""").click()
elif selc == '7':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[7]/a""").click()
elif selc == '8':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[8]/a""").click()
elif selc == '9':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[9]/a""").click()
elif selc == '10':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[10]/a""").click()
elif selc == '11':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[11]/a""").click()
elif selc == '12':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[12]/a""").click()
elif selc == '13':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[13]/a""").click()
elif selc == '14':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[14]/a""").click()
elif selc == '15':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[15]/a""").click()
elif selc == '16':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[16]/a""").click()
elif selc == '17':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[17]/a""").click()
elif selc == '18':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[18]/a""").click()
elif selc == '19':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[19]/a""").click()
elif selc == '20':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[20]/a""").click()
elif selc == '21':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[21]/a""").click()
elif selc == '22':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[22]/a""").click()
elif selc == '23':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[23]/a""").click()
elif selc == '24':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[24]/a""").click()
elif selc == '25':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[25]/a""").click()
elif selc == '26':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[26]/a""").click()
elif selc == '27':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[27]/a""").click()
elif selc == '28':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[28]/a""").click()
elif selc == '29':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[29]/a""").click()
elif selc == '30':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[30]/a""").click()
elif selc == '31':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[31]/a""").click()
elif selc == '32':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[32]/a""").click()
elif selc == '33':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[33]/a""").click()
elif selc == '34':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[34]/a""").click()
elif selc == '35':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[35]/a""").click()
elif selc == '36':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[36]/a""").click()
elif selc == '37':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[37]/a""").click()
elif selc == '38':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[38]/a""").click()
elif selc == '39':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[39]/a""").click()
elif selc == '40':
    driver.find_element_by_xpath("""//*[@id="zg_browseRoot"]/ul/li[40]/a""").click()

time.sleep(1)


def scroll_down(driver):
    driver.execute_script("window.scrollBy(0,9300);")
    time.sleep(1)


scroll_down(driver)

bmp_map = dict.fromkeys(range(0x10000, sys.maxunicode + 1), 0xfffd)

img_src2 = []
file_no = 0

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

reple_result = soup.select('#zg-center-div > #zg-ordered-list')
slist = reple_result[0].find_all('li')

if cnt < 51:
    rank2 = []
    title3 = []
    price2 = []
    score2 = []
    sat_count2 = []
    store2 = []
    count = 0
    img_dir = ff_dir + "\\images"
    os.makedirs(img_dir)
    os.chdir(img_dir)
    for li in slist:
        try:
            photo = li.find('div', 'a-section a-spacing-small').find('img')['src']
        except AttributeError:
            continue
        file_no += 1

        urllib.request.urlretrieve(photo, str(file_no) + '.jpg')
        time.sleep(1)

        if cnt == file_no:
            break

        f = open(f_name, 'a', encoding='UTF-8')
        f.write("-----------------------------------------------------" + "\n")

        print("-" * 70)
        try:
            rank = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except AttributeError:
            rank = ''
            print(rank.replace("#", ""))
        else:
            print("1. 판매순위:", rank)

            f.write('1. 판매순위:' + rank + "\n")

            # 제품 설명
        try:
            title1 = li.find('div', class_='p13n-sc-truncated').get_text().replace("\n", "")
        except AttributeError:
            title1 = ''
            print(title1.replace("\n", ""))
            f.write('2. 제품소개:' + title1 + "\n")
        else:
            title2 = title1.translate(bmp_map).replace("\n", "")
            print("2. 제품소개:", title2.replace("\n", ""))

            count += 1

            f.write('2. 제품소개:' + title2 + "\n")

            # 가격
            try:
                price = li.find('span', 'p13n-sc-price').get_text().replace("\n", "")
            except AttributeError:
                price = ''

            print("3. 가격:", price.replace("\n", ""))
            f.write('3. 가격:' + price + "\n")

            try:
                sat_count = li.find('a', 'a-size-small a-link-normal').get_text().replace(",", "")
            except (IndexError, AttributeError):
                sat_count = '0'
                print('4. 상품평 수: ', sat_count)
                f.write('4. 상품평 수:' + sat_count + "\n")
            else:
                print('4. 상품평 수:', sat_count)
                f.write('4. 상품평 수:' + sat_count + "\n")

            # 상품 별점 구하기
            try:
                score = li.find('span', 'a-icon-alt').get_text()
            except AttributeError:
                score = ' '

            print('5. 평점:', score)
            f.write('5. 평점:' + score + "\n")

            print("-" * 70)

            f.close()

            time.sleep(0.3)

            rank2.append(rank)
            title3.append(title2.replace("\n", ""))
            price2.append(price.replace("\n", ""))

            try:
                sat_count2.append(sat_count)
            except IndexError:
                sat_count2.append(0)

            score2.append(score)

            if count == cnt:
                break

elif cnt >= 51:

    img_src2 = []  # 이미지 URL 저장변수
    file_no = 0

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    count = 0

    reple_result = soup.select('#zg-center-div > #zg-ordered-list')
    slist = reple_result[0].find_all('li')

    # 이미지 내용 추출하기

    img_dir = ff_dir + "\\images"
    os.makedirs(img_dir)
    os.chdir(img_dir)

    for li in slist:

        try:
            photo = li.find('div', 'a-section a-spacing-small').find('img')['src']
        except AttributeError:
            continue
        file_no += 1

        urllib.request.urlretrieve(photo, str(file_no) + '.jpg')
        time.sleep(1)

        if cnt == file_no:
            break

    rank2 = []
    title3 = []
    price2 = []
    score2 = []
    sat_count2 = []
    store2 = []

    for li in slist:
        f = open(f_name, 'a', encoding='UTF-8')
        f.write("-----------------------------------------------------" + "\n")

        # 판매순위
        print("-" * 70)
        try:
            rank = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except AttributeError:
            rank = ''
            print(rank.replace("#", ""))
        else:
            print("1. 판매순위:", rank)

            f.write('1. 판매순위:' + rank + "\n")

    # 제품 설명
        try:
            title1 = li.find('div', class_='p13n-sc-truncated').get_text().replace("\n", "")
        except AttributeError:
            title1 = ''
            print(title1.replace("\n", ""))
            f.write('2. 제품소개:' + title1 + "\n")
        else:
            title2 = title1.translate(bmp_map).replace("\n", "")
            print("2. 제품소개:", title2.replace("\n", ""))

            count += 1

            f.write('2. 제품소개:' + title2 + "\n")

            # 가격
        try:
            price = li.find('span', 'p13n-sc-price').get_text().replace("\n", "")
        except AttributeError:
            price = ''

        print("3. 가격:", price.replace("\n", ""))
        f.write('3. 가격:' + price + "\n")

        try:
            sat_count = li.find('a', 'a-size-small a-link-normal').get_text().replace(",", "")
        except (IndexError, AttributeError):
            sat_count = '0'
            print('4. 상품평 수: ', sat_count)
            f.write('4. 상품평 수:' + sat_count + "\n")
        else:
            print('4. 상품평 수:', sat_count)
            f.write('4. 상품평 수:' + sat_count + "\n")

        # 상품 별점 구하기
        try:
            score = li.find('span', 'a-icon-alt').get_text()
        except AttributeError:
            score = ' '

        print('5. 평점:', score)
        f.write('5. 평점:' + score + "\n")

        print("-" * 70)

        f.close()

        time.sleep(0.5)

        rank2.append(rank)
        title3.append(title2.replace("\n", ""))
        price2.append(price.replace("\n", ""))

        try:
            sat_count2.append(sat_count)
        except IndexError:
            sat_count2.append(0)

        score2.append(score)

    driver.find_element_by_xpath("""//*[@id="zg-center-div"]/div[2]/div/ul/li[3]/a""").click()

    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    reple_result = soup.select('#zg-center-div > #zg-ordered-list')
    slist = reple_result[0].find_all('li')

    for li in slist:
        try:
            photo = li.find('div', 'a-section a-spacing-small').find('img')['src']
        except AttributeError:
            continue
        file_no += 1

        urllib.request.urlretrieve(photo, str(file_no) + '.jpg')
        time.sleep(1)

        if cnt == file_no:
            break

    for li in slist:

        f = open(f_name, 'a', encoding='UTF-8')
        f.write("-----------------------------------------------------" + "\n")

        print("-" * 70)
        try:
            rank = li.find('span', class_='zg-badge-text').get_text().replace("#", "")
        except AttributeError:
            rank = ''
            print(rank.replace("#", ""))
        else:
            print("1. 판매순위:", rank)

            f.write('1. 판매순위:' + rank + "\n")

        # 제품 설명
        try:
            title1 = li.find('div', class_='p13n-sc-truncated').get_text().replace("\n", "")
        except AttributeError:
            title1 = ''
            print(title1.replace("\n", ""))
            f.write('2. 제품소개:' + title1 + "\n")
        else:
            title2 = title1.translate(bmp_map).replace("\n", "")
            print("2. 제품소개:", title2.replace("\n", ""))

            count += 1

            f.write('2. 제품소개:' + title2 + "\n")

        # 가격
        try:
            price = li.find('span', 'p13n-sc-price').get_text().replace("\n", "")
        except AttributeError:
            price = ''

        print("3. 가격:", price.replace("\n", ""))
        f.write('3. 가격:' + price + "\n")

        try:
            sat_count = li.find('a', 'a-size-small a-link-normal').get_text().replace(",", "")
        except IndexError:
            sat_count = '0'
            print('4. 상품평 수: ', sat_count)
            f.write('4. 상품평 수:' + sat_count + "\n")
        except AttributeError:
            sat_count = '0'
            print('4. 상품평 수: ', sat_count)
            f.write('4. 상품평 수:' + sat_count + "\n")
        else:
            print('4. 상품평 수:', sat_count)
            f.write('4. 상품평 수:' + sat_count + "\n")

        # 상품 별점 구하기
        try:
            score = li.find('span', 'a-icon-alt').get_text()
        except AttributeError:
            score = ' '

        print('5. 평점:', score)
        f.write('5. 평점:' + score + "\n")

        print("-" * 70)

        f.close()

        time.sleep(0.5)

        rank2.append(rank)
        title3.append(title2.replace("\n", ""))
        price2.append(price.replace("\n", ""))

        try:
            sat_count2.append(sat_count)
        except IndexError:
            sat_count2.append(0)

        score2.append(score)

        if count == cnt:
            break

else :
    print("에러")

amazon_best_seller = pd.DataFrame()
amazon_best_seller['판매순위'] = rank2
amazon_best_seller['제품소개'] = pd.Series(title3)
amazon_best_seller['판매가격'] = pd.Series(price2)
amazon_best_seller['상품평 갯수'] = pd.Series(sat_count2)
amazon_best_seller['상품평점'] = pd.Series(score2)

amazon_best_seller.to_csv(fc_name, encoding="utf-8-sig", index=True)

amazon_best_seller.to_excel(fx_name, index=True)

e_time = time.time()
t_time = e_time - s_time

orig_stdout = sys.stdout
f = open(f_name, 'a', encoding='UTF-8')
sys.stdout = f

import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(fx_name)
sheet = wb.ActiveSheet
sheet.Columns(3).ColumnWidth = 30  # 이미지 가로 사이즈에 맞게 컬럼 크기 조정
row_cnt = cnt + 1
sheet.Rows("2:%s" % row_cnt).RowHeight = 120  # 이미지 세로 사이즈에 맞게 로우 크기 조정

ws = wb.Sheets("Sheet1")
col_name2 = []
file_name2 = []

for a in range(2, cnt + 2):
    col_name = 'C' + str(a)
    col_name2.append(col_name)

for b in range(1, cnt + 1):
    file_name = img_dir + '\\' + str(b) + '.jpg'
    file_name2.append(file_name)

for i in range(0, cnt):
    rng = ws.Range(col_name2[i])
    image = ws.Shapes.AddPicture(file_name2[i], False, True, rng.Left, rng.Top, 130, 100)
    excel.Visible = True
    excel.ActiveWorkbook.Save()
sys.stdout = orig_stdout
f.close()

print("1.파일 저장 완료: txt 파일명 : %s " % f_name)
print("2.파일 저장 완료: csv 파일명 : %s " % fc_name)
print("3.파일 저장 완료: xls 파일명 : %s " % fx_name)