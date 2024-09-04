from PyQt6.QtWidgets import QPushButton,QLabel,QWidget,QApplication,QMessageBox,QProgressBar,QComboBox,QDateEdit
from PyQt6.uic import loadUi
from PyQt6.QtGui import QCloseEvent, QIcon
import sys
import sqlite3
import keyboard
import threading
import ctypes
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import shutil
import datetime
import pyautogui
import json
import pycurl
import urllib.parse
from io import BytesIO
from parsel import Selector

path = os.getcwd()
icon = "assests/icon.ico"

title = "برنامج تحديث"

zonesIds = {"منطقة الجوف":112 , "منطقة نجران": 110, "المنطقة الشرقية": 104, "منطقة عسير": 105, "منطقة القصيم":103 , "منطقة حائل":107 ,"منطقة تبوك":106,"منطقة مكة المكرمة":101,"منطقة المدينة المنورة":102,"منطقة الباحة":111,"منطقة الرياض":100,"منطقة جازان":109,"منطقة الحدود الشمالية":108}
centersId = {"الجوف":120 , "القريات":123,"نجران":113,"الخفجي":127,"الدمام":124,"الهفوف":128,"حفر الباطن":126,"الجبيل":274,"ابها":129,"محايل عسير":131,"بيشة":130,"الرس":122,"القصيم":121,"حائل":105,"تبوك":102,"الطائف":115,"جدة عسفان":287,"الخرمة":116,"جدة جنوب":111,"مكة المكرمة":114,"جدة شمال":112,"ينبع":119,"المدينة المنورة":118,"الباحة":117,"الرياض حي المونسية":103,"المجمعة":108,"الرياض حي القيروان":100,"القويعية":125,"جنوب شرق الرياض مخرج سبعة عشر":319,"وادي الدواسر":109,"الرياض حي الشفا طريق ديراب":104,"الخرج":106,"جيزان":110,"عرعر":107}
centersIdZones = {
    "منطقة الجوف":{
    "الجوف":120 , "القريات":123
    },"منطقة نجران":{
    "نجران":113
    },"المنطقة الشرقية":{
        "الخفجي":127,"الدمام":124,"الهفوف":128,"حفر الباطن":126,"الجبيل":274
    },"منطقة عسير":{
        "ابها":129,"محايل عسير":131,"بيشة":130
    },"منطقة القصيم":{
        "الرس":122,"القصيم":121
    },"منطقة حائل":{
        "حائل":105
    },"منطقة تبوك":{
        "تبوك":102
    },"منطقة مكة المكرمة":{
        "الطائف":115,"جدة عسفان":287,"الخرمة":116,"جدة جنوب":111,"مكة المكرمة":114,"جدة شمال":112
    },"منطقة المدينة المنورة":{
        "ينبع":119,"المدينة المنورة":118
    },"منطقة الباحة":{
        "الباحة":117
    },"منطقة الرياض":{
        "الرياض حي المونسية":103,"المجمعة":108,"الرياض حي القيروان":100,"القويعية":125,"جنوب شرق الرياض مخرج سبعة عشر":319,"وادي الدواسر":109,"الرياض حي الشفا طريق ديراب":104,"الخرج":106}
    ,"منطقة جازان":{
        "جيزان":110
    },"منطقة الحدود الشمالية":{
        "عرعر":107
    }}
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False
class shortCuts(QWidget):
    def __init__(self):
        super().__init__()
        loadUi("shortCut.ui",self)
        self.setFixedSize(442,179)
        cr.execute("SELECT refresh,stop FROM shortcut")
        values = cr.fetchall()[0]
        refresh = values[0]
        stop = values[1]
        self.lineEdit.setText(str(refresh))
        self.lineEdit.setDisabled(True)
        self.lineEdit.setReadOnly(True)

        self.lineEdit_2.setText(str(stop))
        self.lineEdit_2.setDisabled(True)
        self.lineEdit_2.setReadOnly(True)

        self.shortCuts_2.clicked.connect(lambda:self.changeShortCut("refresh"))
        self.stopPrograme.clicked.connect(lambda:self.changeShortCut("stop"))
        self.complete.clicked.connect(self.completeChanged)
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(icon))
    def watingToPress(self,event):
            keyboard.unhook_all()
            key = event.name
            if self.fromw == "refresh":
                self.lineEdit.setText(key)
                self.lineEdit.setDisabled(True)
            elif self.fromw == "stop":
                self.lineEdit_2.setText(key)
                self.lineEdit_2.setDisabled(True)
    def changeShortCut(self,fromW):
        self.fromw = fromW
        if fromW == "refresh":
            self.lineEdit_2.setDisabled(True)
            self.lineEdit.setDisabled(False)
        elif fromW == "stop":

            self.lineEdit.setDisabled(True)
            self.lineEdit_2.setDisabled(False)
            
        key = keyboard.on_press(self.watingToPress)

    def completeChanged(self):
        cr.execute("UPDATE shortcut set refresh = ? , stop = ?",(self.lineEdit.text(),self.lineEdit_2.text()))
        con.commit()
        ctypes.windll.user32.MessageBoxW(0, "تم التغيير بنجاح", "نجاح")
        try:
            self.destroy()
            self.hide()
        except:
            pass

class Settings(QWidget):
    def __init__(self):
        super().__init__()
        loadUi("settings.ui",self)
        self.setFixedSize(246,322)

        self.spinBoxX.setMaximum(99999999)
        self.spinBoxY.setMaximum(99999999)
        self.seconds.setMaximum(99999999)
    
        cr.execute("SELECT x,y,s,zoneId,centerId FROM settings")
        x,y,s,zoneId,centerId = cr.fetchall()[0]

        self.spinBoxX.setValue(int(x))
        self.spinBoxY.setValue(int(y))
        self.seconds.setValue(s)
        self.seconds.setSingleStep(0.1)

        zone = list(zonesIds.keys())[list(zonesIds.values()).index(int(zoneId))]
        center = list(centersId.keys())[list(centersId.values()).index(int(centerId))]

        for i in zonesIds.keys():
            self.zone.addItem(i)

        self.zone.setCurrentText(zone)

        self.changeValues()

        self.center.setCurrentText(center)

        self.zone.currentTextChanged.connect(self.changeValues)

        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(icon))

        self.save.clicked.connect(self.saveSettings)

    def saveSettings(self):
        center = centersIdZones.get(self.zone.currentText()).get(self.center.currentText())
        zone = zonesIds.get(self.zone.currentText())

        cr.execute("UPDATE settings set x=?,y=?,s=?,zoneId=?,centerId=?,date=?",(self.spinBoxX.value(),self.spinBoxY.value(),self.seconds.value(),zone,center,self.date.date().toPyDate()))
        con.commit()

        ctypes.windll.user32.MessageBoxW(0, "تم التغيير بنجاح", "نجاح")
        try:
            self.destroy()
            self.hide()
        except:
            pass
    def changeValues(self):
        self.center.clear()
        for i in centersIdZones.get(self.zone.currentText()).keys():
            self.center.addItem(i)
class Window (QWidget):
    def __init__(self):
        super().__init__()
        loadUi("mainWindow.ui",self)
        self.setFixedSize(635,416)
        self.setWindowTitle(title)
        self.setWindowIcon(QIcon(icon))
        self.playButton.clicked.connect(self.working)
        self.shortCuts.clicked.connect(self.shortCutsFun)
        self.settingsbtn.clicked.connect(self.settings)

    def confirmWorking(self):
        try:
            self.confirmWorkingbutton.destroy()
            self.confirmWorkingbutton.hide()
            self.playButton.show()
        except:
            pass       

        thread = threading.Thread(target=lambda x=self.driver:traking(x))
        thread.daemon = True
        thread.start()

    def working(self):
        cr.execute("SELECT Field1 FROM date")
        limitDate = str(cr.fetchone()[0]).split("-")
        if datetime.date.today() < datetime.date(int(limitDate[0]),int(limitDate[1]),int(limitDate[2])):
            try:
                self.playButton.hide()
            except:
                pass
            self.confirmWorkingbutton = QPushButton("تأكيد",self)
            self.confirmWorkingbutton.setStyleSheet("QPushButton {font-size:20px;text-align:center;background-color:green;}")
            self.confirmWorkingbutton.setGeometry(260,160,161,51)
            self.confirmWorkingbutton.clicked.connect(self.confirmWorking)
            self.confirmWorkingbutton.show()
            options = webdriver.ChromeOptions()
            options.add_argument(f"user-data-dir={path}/data/User Data")
            options.add_argument("ignore-certificate-errors")
            options.add_experimental_option("detach", True)
            self.driver = webdriver.Chrome(options=options,service=ChromeService(ChromeDriverManager().install()))
            self.driver.get("https://vi.vsafety.sa/book/edit")
        else:
            d = QMessageBox(parent=self,text="Error Programe cant be run")
            d.setWindowTitle("Failed")
            d.setIcon(QMessageBox.Icon.Critical)
            d.setStyleSheet("background-color:white")
            d.show()
    def shortCutsFun(self):
        self.shortCutWindow = shortCuts()
        self.shortCutWindow.show()
    
    def settings(self):
        self.settingsWindow = Settings()
        self.settingsWindow.show()

    def closeEvent(self,event):
        global keepTraking
        keepTraking = False
        event.accept()
        try:
            self.driver.quit()
        except:
            pass

def checkAppointments(zone,center,appointment_date):
    # ------------- Customer Info -------------
    # Encode the customer name using URL encode, maybe not necessarily
    customer_name = urllib.parse.quote_plus("نعمان الحسين")
    # customer_name = "%D9%86%D8%B9%D9%85%D8%A7%D9%86+%D8%A7%D9%84%D8%AD%D8%B3%D9%8A%D9%86"
    customer_mobile_no = "533948458" # Mobile number, a valid mobile number has 9 digits

    # ------------- Plate Formatting -------------
    # plate = "7653 T N J"
    # plate = "7653 ح ن ط"
    plate_1 = "J"
    plate_2 = "N"
    plate_3 = "T"
    plate_4 = "7653"

    # ------------- Zone & Center -------------

    # ------------- Appointment Date -------------

    # ------------- Getting a list of appointments -------------
    url = "https://vi.vsafety.sa/en/book/apply?ajax_form=1&_wrapper_format=drupal_ajax"

    headers = ['Accept: application/json, text/javascript, */*; q=0.01', 'Accept-Language: en-US,en;q=0.8', 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8', 'Origin: https://vi.vsafety.sa', 'Referer: https://vi.vsafety.sa/en/book/apply', 'User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.0.0 Safari/537.36', 'X-Requested-With: XMLHttpRequest']
    data = f"form_build_id=form-itL0uBlkYKnpmIgGEFMhtH1IbGMIKr8OZ3WPz7TwwPU&form_id=svis_book_book_apply&customer_name={customer_name}&customer_mobile_no={customer_mobile_no}&country=%2B966&customer_email=&delegate_type=RESIDENT&delegate_name=&delegate_mobile_no=&delegate_country=%2B966&delegate_id=&delegate_hijri=&delegate_residency=&delegate_birthday=&delegate_national_id=&delegate_national_birthday=&delegate_gulf_country=&registered_vehicle=1&registration_country_id=96&plate_1={plate_1}&plate_2={plate_2}&plate_3={plate_3}&plate_4={plate_4}&plate_number=&plate_type_id=55&custom_cert_no=&vehicle_class_id=3&service_id=110&location=0&zone_id={zone}&center_id={center}&appointment_date={appointment_date}&timeslot=0&ajax_loader=1&captcha_sid=5021584&captcha_token=4GVhsxnFO6w5cUtasL_BFyOJvytLI0Ztnwr8KpXBFUo&captcha_response=Google+no+captcha&captcha_cacheable=1&_triggering_element_name=appointment_date&ajax_page_state%5BdialogType%5D=ajax&_drupal_ajax=1&ajax_page_state%5Btheme%5D=vehicle&ajax_page_state%5Btheme_token%5D=&ajax_page_state%5Blibraries%5D=bootstrap%2Fpopover%2Cbootstrap%2Ftooltip%2Ccaptcha%2Fbase%2Cclientside_validation_jquery%2Fcv.jquery.ckeditor%2Cclientside_validation_jquery%2Fcv.jquery.validate%2Cclientside_validation_jquery%2Fcv.pattern.method%2Ccore%2Fdrupal.date%2Ccore%2Fdrupal.states%2Ccore%2Finternal.jquery.form%2Cfontawesome%2Ffontawesome.svg%2Cfontawesome%2Ffontawesome.svg.shim%2Cgoogle_tag%2Fgtag%2Cgoogle_tag%2Fgtag.ajax%2Cgoogle_tag%2Fgtm%2Cslick%2Fslick%2Cslick%2Fslick.easing%2Cslick%2Fslick.mousewheel%2Csuperfish%2Fsuperfish%2Csuperfish%2Fsuperfish_hoverintent%2Csuperfish%2Fsuperfish_smallscreen%2Csuperfish%2Fsuperfish_supersubs%2Csuperfish%2Fsuperfish_supposition%2Csvis_book%2Fselecticons%2Csvis_book%2Fsvis_book_form%2Csystem%2Fbase%2Cvehicle%2Fglobal-scripts%2Cwebform_bootstrap%2Fwebform_bootstrap"
    buffer = BytesIO()


    c = pycurl.Curl()
    c.setopt(c.URL, url)
    c.setopt(c.HTTPHEADER, headers)
    c.setopt(c.POSTFIELDS, data)
    c.setopt(c.WRITEDATA, buffer)
    c.perform()
    c.close()

    listOfAppointments = json.loads(buffer.getvalue().decode("utf-8"))

    # ------------- Format list of appointments to get available appointments -------------

    listOptions = [{'value': opt.xpath('@value').get(), 'text': opt.xpath('text()').get()} for opt in Selector(listOfAppointments[1]["data"]).xpath('//option')]


    noAppointments = listOptions[0]["value"]
    if noAppointments == "0":
        print("No appointments available")
        return False
    return True

def traking(driver):
    global keepTraking
    keepTraking = True
    cr.execute("SELECT refresh,stop FROM shortcut")
    values = cr.fetchall()[0]
    refresh = values[0]
    stop = values[1]
    
    cr.execute("SELECT x,y,s,zoneId,centerId,date FROM settings")
    x,y,seconds,zone,center,date = cr.fetchall()[0]
    
    

    while keepTraking:
        # try:
            try:
                s = driver.title
            except:
                break
            if keyboard.is_pressed(stop):
                time.sleep(1)
                while True:
                    if keyboard.is_pressed(stop):
                        time.sleep(1)
                        break

            driver.execute_script("window.scrollTo(0, 3200)")

            if checkAppointments(zone,center,date):
                while keepTraking:
                    try:
                        pyautogui.moveTo(x,y)
                        pyautogui.leftClick()
                        element = WebDriverWait(driver,5).until(EC.element_to_be_clickable((By.XPATH,"//button[@id='edit-submit']")))
                    except:
                        break
                    time.sleep(seconds)
                break
        # except:
        #     pass
def runAll():
    global con,cr,app,window
    app = QApplication(sys.argv)
    con = sqlite3.connect("app.db",check_same_thread=False)
    cr = con.cursor()

    window = Window()
    window.show()
    app.exec()

if __name__ == "__main__":
    if is_admin():
        runAll()
    else:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)