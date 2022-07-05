from distutils.log import debug
from regex import I
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import logging

logging.basicConfig(
    filename="app.log",
    level=logging.INFO,
    format="%(asctime)s:%(levelname)s:%(message)s",
)

# read/write Excel
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import time

""" Download chromedriver """
# chrome://version/
# https://chromedriver.storage.googleapis.com/index.html
# pip install openpyxl


logging.info("...Iniciando...")
logging.info("reading excel..")


wb = load_workbook("fichas/PRODUCCION BERMUDA OVER UNIT - RUSTICO.xlsm", data_only=True)
ws = wb.active

# driver_service = Service(executable_path="./selenium-driver/chromedriver.exe")
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.maximize_window()
# URL
driver.get("http://vpn.grisino.com:8001/maertest")


class Login:
    def __init__(self, user, password):
        self.user = user
        self.password = password

    def login(self):
        logging.info("Iniciando sesion..")
        login_user = WebDriverWait(driver, 10).until(
            expected_conditions.presence_of_element_located((By.ID, "ext-comp-1002"))
        )

        login_user.send_keys(self.user)

        password_user = WebDriverWait(driver, 10).until(
            expected_conditions.presence_of_element_located((By.ID, "ext-comp-1004"))
        )
        password_user.send_keys(self.password)

        login_btn = driver.find_element(By.ID, "ext-gen31")
        login_btn.click()


class LoadFile:
    logging.info("Cargando file...")

    def loop(self, rango_cod_color, lista_cod_color):
        # Iterar por cod de color
        for cod in rango_cod_color:
            for i in cod:
                lista_cod_color.append(i.value)
                print(i.value)

    def loop_cod_color(self, rango_cod_color, lista_cod_color, celda):
        # Iterar por cod de color
        if type(celda).__name__ == "MergedCell":
            print("Combinada, color todos")
            lista_cod_color.append(celda.value)
            return True
        else:
            for cod in rango_cod_color:
                for i in cod:
                    lista_cod_color.append(i.value)
                    print(i.value)
            return False

    def comprobar_y_cargar(
        self,
        actions,
        descripcion_validacion,
        talles,
        lista_cod_color,
        cantidad_insumo_confeccion,
        insumo_confeccion,
    ):
        if "TALLE" in descripcion_validacion:
            logging.info("Cargnado insumo por talle...")
            # Por cada talle cargar...
            for i in talles:
                time.sleep(2)
                if i != None:
                    logging.info("talles disponibles: ", i)
                    print("talles disponibles: ", i)
                    time.sleep(2)
                    for cod_color in lista_cod_color:
                        if cod_color != None:
                            self.load_insumo_por_talle(
                                actions,
                                insumo_confeccion,
                                cod_color,
                                cantidad_insumo_confeccion,
                                i,
                            )
                            print("Carga de insumo por talle finalizada...")
                            logging.info("Carga de insumo por talle finalizada...")
        else:
            for i in lista_cod_color:
                if i != None:
                    self.load_insumo2(
                        actions,
                        insumo_confeccion,
                        i,
                        cantidad_insumo_confeccion,
                    )
                else:
                    pass

    def load_insumo(self, actions, insumo, color_insumo, cantidad):
        if insumo != None:
            logging.info("Cargando insumo")
            time.sleep(2)
            actions.send_keys(insumo + "." + color_insumo)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(2)
            actions.send_keys(cantidad)
            time.sleep(2)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
        else:
            actions.send_keys(Keys.ESCAPE)
            actions.perform()
            actions.send_keys(Keys.ESCAPE)
            actions.perform()

    def load_insumo2(self, actions, insumo, i, cantidad):
        if insumo != None:
            logging.info("Cargando insumo")
            time.sleep(2)
            actions.send_keys(insumo + "." + i)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(2)
            actions.send_keys(cantidad)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
        else:
            pass

    def load_insumo_por_talle(self, actions, insumo, color_insumo, cantidad, i):
        if insumo != None and color_insumo != None:
            logging.info("Cargando insumo")
            time.sleep(1)
            actions.send_keys(insumo + "." + color_insumo)
            actions.perform()
            time.sleep(1)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(2)
            actions.send_keys(cantidad)
            actions.perform()
            time.sleep(3)
            actions.send_keys(Keys.TAB)
            actions.perform()
            time.sleep(3)
            actions.send_keys(color_insumo)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
            print(type(i))
            if i != None:
                if i != "2":
                    actions.send_keys(i)
                    time.sleep(2)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.TAB)
                    actions.perform()
                else:
                    actions.send_keys(i)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.ARROW_DOWN)
                    actions.send_keys(Keys.ARROW_DOWN)
                    actions.send_keys(Keys.ENTER)
                    actions.perform()
                    time.sleep(1)
                    actions.send_keys(Keys.TAB)
                    actions.perform()

    def load_new(self):
        try:
            time.sleep(1)
            logging.info("Cargando nueva ficha...")
            btn_produccion = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div[1]/div[2]/div/div/div/div[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td[2]/table/tbody/tr[2]/td[2]/em/button",
                    )
                )
            )
            btn_produccion.click()

            btn_ficha_tecnica = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (By.ID, "menuPrincipalProducciónFichas T?cnicas")
                )
            )
            btn_ficha_tecnica.click()

            btn_maxim_ft = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div[4]/div[2]/div/div/div/div[1]/div[6]/div[1]/div/div/div/div[3]",
                    )
                )
            )
            btn_maxim_ft.click()

            btn_add_new = WebDriverWait(driver, 35).until(
                expected_conditions.presence_of_element_located(
                    (
                        By.XPATH,
                        "/html/body/div[4]/div[2]/div/div/div/div[1]/div[6]/div[2]/div[1]/div/div/div/div/div/div[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/em/button",
                    )
                )
            )
            btn_add_new.click()

            time.sleep(10)

            input_coleccion = driver.find_element(
                By.XPATH,
                "//*[@id='ext-comp-1251']",
            )
            actions = ActionChains(driver)
            coleccion = ws["G1"].value
            time.sleep(1)
            input_coleccion.send_keys(coleccion)
            time.sleep(3)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(3)
            actions.send_keys(Keys.TAB)
            actions.perform()
            producto = ws["B2"].value
            time.sleep(1)
            actions.send_keys(producto)
            time.sleep(2)
            actions.perform()
            actions.send_keys(Keys.ARROW_DOWN)
            actions.perform()
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(3)
            actions.send_keys(Keys.TAB)
            actions.perform()
            actions.send_keys(Keys.TAB)
            actions.perform()
            molde = ws["T2"].value
            actions.send_keys(molde)
            actions.perform()
            time.sleep(4)

            btn_add_rule = driver.find_element(
                By.XPATH,
                "//*[@id='ext-gen647' or @id='ext-gen649' or @id='ext-gen651' or @id='ext-gen655'or @id='ext-gen657' or @id='ext-gen907' or @id='ext-gen653' or @id='ext-gen905' or @id='ext-gen905' or @id='ext-gen664' or @id='ext-gen672']",
            )

            btn_add_rule.click()
            time.sleep(1)
            proceso_corte = driver.find_element(By.ID, "ext-comp-1263")
            proceso_corte.send_keys("100-CORTE ORIGINAL")
            time.sleep(1)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            actions.send_keys(Keys.ESCAPE)
            actions.perform()
            time.sleep(2)

            nueva_entrada = driver.find_element(
                By.XPATH,
                "/html/body/div[4]/div[2]/div/div/div/div[1]/div[8]/div[2]/div[1]/div/div/div/div/div/div[1]/div[2]/div/div/div/div/div[2]/div/div[1]/div[2]/div[1]/div/table/tbody/tr/td[3]/div/span/table/tbody/tr[2]/td[2]/em/button",
            )
            time.sleep(5)
            nueva_entrada.click()
            time.sleep(2)

            agregar_insumo = driver.find_element(
                By.XPATH,
                "/html/body/div[24]/div[2]/div[1]/div/div/div/div/div/div/div/div/div[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td[1]/table/tbody/tr[2]/td[2]/em/button",
            )
            time.sleep(3)
            agregar_insumo.click()
            time.sleep(1)
            actions.send_keys(Keys.TAB)
            actions.perform()

            insumo_1 = ws["I6"].value
            color_inusmo = ws["L7"].value
            color_insumo2 = ws["L9"].value
            cantidad_insumo_1 = str(ws["J6"].value)
            cantidad_insumo_2 = ws["J8"].value

            time.sleep(2)
            self.load_insumo(actions, insumo_1, color_inusmo, cantidad_insumo_1)
            time.sleep(2)
            COLOR1 = ws["L4"].value
            COLOR2 = ws["N4"].value
            insumo_2 = ws["I8"].value
            insumo_3 = ws["I10"].value
            color_insumo3 = ws["L11"].value
            cantidad_insumo_3 = ws["J10"].value
            insumo_4 = ws["I12"].value
            insumo_5 = ws["I14"].value
            insumo_6 = ws["I16"].value
            color_insumo4 = ws["N5"].value
            # XTA004
            color_insumo5 = ws["N5"].value
            # XTD001
            color_insumo6 = ws["N5"].value
            cantidad_insumo_4 = str(ws["J12"].value)
            cantidad_insumo_5 = str(ws["J14"].value)
            cantidad_insumo_6 = str(ws["J16"].value)

            # Si insumo existe.. agregar otro
            # Se puede hacer una fx decoradora -----------------------------------------------------------
            if insumo_2 != None:
                agregar_insumo.click()
                actions.send_keys(Keys.TAB)
                time.sleep(2)
                actions.perform()
                time.sleep(2)
                self.load_insumo(actions, insumo_2, color_insumo2, cantidad_insumo_2)
            else:
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

            time.sleep(3)

            if insumo_3 != None:
                agregar_insumo.click()
                actions.send_keys(Keys.TAB)
                time.sleep(2)
                actions.perform()
                time.sleep(2)
                self.load_insumo(actions, insumo_3, color_insumo3, cantidad_insumo_3)
            else:
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

            time.sleep(3)

            if insumo_4 != None:
                agregar_insumo.click()
                actions.send_keys(Keys.TAB)
                time.sleep(2)
                actions.perform()
                time.sleep(2)
                self.load_insumo(actions, insumo_4, color_insumo4, cantidad_insumo_4)

            else:
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

            time.sleep(3)

            if insumo_5 != None:
                agregar_insumo.click()
                actions.send_keys(Keys.TAB)
                time.sleep(2)
                actions.perform()
                time.sleep(2)
                self.load_insumo(actions, insumo_5, color_insumo5, cantidad_insumo_5)

            else:
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

            time.sleep(3)

            if insumo_6 != None:
                agregar_insumo.click()
                actions.send_keys(Keys.TAB)
                time.sleep(2)
                actions.perform()
                time.sleep(2)
                self.load_insumo(actions, insumo_6, color_insumo6, cantidad_insumo_6)
            else:
                actions.send_keys(Keys.ESCAPE)
                actions.perform()
                actions.send_keys(Keys.ESCAPE)
                actions.perform()

            # -------------------------------------------------- ---------------------------------------------------------------------------
        except (TimeoutException) as error:
            logging.warning("Error: ", error)
            logging.info("Erorr: ", error)
            print(error)
            driver.close()
        except (Exception) as error_excepction:
            logging.warning("Error: ", error_excepction)
            logging.info("Erorr: ", error_excepction)
            print(error_excepction)


# Loggearse
log = Login("Gfrassetti", "Guido")
log.login()


# Cargar Nueva ficha
load = LoadFile()
load.load_new()
