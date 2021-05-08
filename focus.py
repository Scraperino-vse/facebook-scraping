import time
import xlwt
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException

website_url = input("Jmeno FB stranky pro scrappovani (priklad: https://www.facebook.com/Blesk.cz/): ")

# Nastavení chromu pro vypnutí upozornění požadavku na zasílání notifikací
chrome_options = webdriver.ChromeOptions()
prefs = {"profile.default_content_setting_values.notifications" : 2}
chrome_options.add_experimental_option("prefs",prefs)

# Spuštění prohlížeče
browser = webdriver.Chrome("C:/Users/Mirek/Desktop/Homebrew/python/chromedriver.exe",chrome_options=chrome_options)

browser.maximize_window()

print("Spuštění prohlížeče....")

# Otevření login stránky
browser.get('https://www.facebook.com/')
assert 'Facebook' in browser.title

time.sleep(3)

# Přihlášení
accept_button_cookies = browser.find_element_by_xpath('//*[@title="Přijmout vše"]').click()

print("Cookies přijaty...")
print("Přihlašování...")
# Username
userElem = browser.find_element_by_id("email")
userElem.send_keys('')      #přihlašovací mail
print("Uživatelský email.... OK")

# Password
passwordElem = browser.find_element_by_id("pass")
passwordElem.send_keys('')
print("Heslo.... OK")

# Login
loged_in = browser.find_element_by_id("u_0_b").click()

print("Úspěšně přihlášeno.... OK")

time.sleep(2)

# Otevření požadované stránky
browser.get(website_url)

time.sleep(1)

SCROLL_PAUSE_TIME = 2

# Navrátí aktuální výšku okna
last_height = browser.execute_script("return document.body.scrollHeight")


# O kolik obrazovek se má robot posunout na základě požadovaného počtu příspěvků (nastavitelný parametr)
for x in range(10):
    # Scrollnutí obrazovky na spodek s ohledem na aktuální velikost obrazovky
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Zastavení pohybu pro načtení načtení stránky (kvůli lazy-loading technice na novém FB layoutu)
    time.sleep(SCROLL_PAUSE_TIME)

    print("Načítám požadované množství příspěvků....")
    print(x)

    if x > 20:
        SCROLL_PAUSE_TIME = 4

    # Vytvoření nové výšky pro scroll a porovnání se starou výškou
    new_height = browser.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height

# Vytvoření nového excel dokumentu a sheetu uvnitř něj
wb = xlwt.Workbook()
sheet1 = wb.add_sheet('Data')


# Vyscrollovaní zpět na hořejšek stránky (opět kvůli lazy-loadingu)
browser.execute_script("window.scrollTo(0, 0);")

# Rozbalení komentářů a posun okna do pozice komentáře
expand_comments = browser.find_elements_by_xpath("//span[@class='j83agx80 fv0vnmcu hpfvmrgz']")
i = 0
for expand in expand_comments:
    if expand.text == "":
        print("Prázdné")
        browser.execute_script("window.scrollBy(0, 1080);")
        time.sleep(3)
        if i < 3:
            print("Odkaz nenalezen, počet zbývajících obnovení existujících expanzí komentářů:")
            print(3-i)
            expand_comments = browser.find_elements_by_xpath("//span[@class='j83agx80 fv0vnmcu hpfvmrgz']")
            i += 1
    else:
        print("Lokace rozbalovacího tlačítka:")
        print(expand.location)
        print(expand.text)
        browser.execute_script("arguments[0].scrollIntoView();", expand)
        time.sleep(1)
        browser.execute_script("arguments[0].click();", expand)
        time.sleep(1)
        print("Komentáře rozbaleny....")
        time.sleep(1)

print("Zapisuji jména do souboru Data.xlsx")

browser.execute_script("window.scrollTo(0, 0);")
time.sleep(8)

# Zápis vizualně viditelných jmen do excel souboru a ošetření chyb, aby se program nezavřel
visible_comments = browser.find_elements_by_xpath("//span[@class='d2edcug0 hpfvmrgz qv66sw1b c1et5uql rrkovp55 a8c37x1j keod5gw0 nxhoafnm aigsh9s9 d9wwppkn fe6kdd0r mau55g9w c8b282yb mdeji52x e9vueds3 j5wam9gi lrazzd5p oo9gr5id']")

for names in visible_comments:
    if names.text == "":
        print("Nenalezeno jméno, přeskočeno")
    else:
        index = visible_comments.index(names)
        browser.execute_script("arguments[0].scrollIntoView();", names)
        all_people = browser.find_elements_by_xpath("//div[@class='b3i9ofy5 e72ty7fz qlfml3jp inkptoze qmr60zad rq0escxv oo9gr5id q9uorilb kvgmc6g5 cxmmr5t8 oygrvhab hcukyx3x d2edcug0 jm1wdb64 l9j0dhe7 l3itjdph qv66sw1b']")
        time.sleep(0.1)
        print(names.text)
        print(visible_comments.len())
        commentator = all_people[index-1]
        comment = commentator.text

        firstname = names.text.strip().split(' ')[0]
        lastname = ' '.join((names.text + ' ').split(' ')[1:]).strip()

        comment_no_firstname = comment.replace(firstname, '')
        comment_no_lastname = comment_no_firstname = comment.replace(lastname, '')

        if "Přední fanoušek" in comment_no_lastname:
            comment_no_lastname = comment_no_lastname.replace('Přední fanoušek', '')
            sheet1.write(index, 3, "přední fanoušek")

        sheet1.write(index, 0, firstname)
        sheet1.write(index, 1, lastname)
        sheet1.write(index, 4, comment_no_lastname)
        print(comment_no_lastname.len())
        formula=f'IF(OR((RIGHT(TRIM(B{index+1}),1)="a"),(RIGHT(TRIM(B{index+1}),1)="á")),"ž","m")'
        sheet1.write(index, 2, xlwt.Formula(formula))

print("Ukládám výsledky")

# Uložení výsledků a terminace programu   
wb.save('data_new.xls')
time.sleep(5)
print("Uloženo, bye bye")

