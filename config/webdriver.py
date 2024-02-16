import os
from selenium import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.opera import OperaDriverManager
from webdriver_manager.firefox import GeckoDriverManager
import psutil

def iniciar_navegador(browser):
    def verde(texto):
        print(f'\x1b[32m{texto}\x1b[0m')

    if browser == 'Edge':
        try:
            options = webdriver.EdgeOptions()
            options.add_argument('--disable-gpu')
            options.use_chromium = True
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_experimental_option('prefs', {'download.default_directory': os.path.join(os.getcwd(), 'geracao')})
            options.add_argument("--window-size=1920x1080")

            # Use o EdgeChromiumDriverManager para obter o executável do driver automaticamente
            driver_manager = EdgeChromiumDriverManager()
            driver = webdriver.Edge(executable_path=driver_manager.install(), options=options)
            driver.maximize_window()

            print('>>> Webdriver Microsoft Edge iniciado...')

            return driver
        except Exception as e:
            print('error ', e)
    elif browser == 'Chrome':
        try:
            # Define o diretório do perfil do Instagram
            dir_path = os.getcwd()
            profile = os.path.join(dir_path, "Chrome - Bot Engenharia")


            options = webdriver.ChromeOptions()
            
            options.add_argument(r"user-data-dir={}".format(profile))
            options.add_argument('--disable-gpu')
            # options.add_argument('--headless')
            options.use_chromium = True
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_experimental_option('prefs', {'download.default_directory': os.path.join(os.getcwd(), 'geracao')})
            options.add_argument("--window-size=1920x1080")
            try:
                driver_manager = ChromeDriverManager()
                driver = webdriver.Chrome(executable_path=driver_manager.install(), options=options)
            except Exception as e: 
                print('........',e)
                driver_manager = ChromeDriverManager().install()
                driver = webdriver.Chrome(options=options)

            verde(f'\n💻  Navegador Chrome - Bot Engenharia iniciado! ✅\n')
            driver.maximize_window()
            return driver
        
        except Exception as erro: print(f'🚨 : {erro}')
    elif browser == 'Opera':
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--disable-gpu')
            options.add_experimental_option('prefs', {'download.default_directory': os.path.join(os.getcwd(), 'geracao')})
            options.add_argument("--window-size=1920x1080")

            # Use o OperaDriverManager para obter o executável do driver automaticamente
            driver_manager = OperaDriverManager()
            driver = webdriver.Chrome(executable_path=driver_manager.install(), options=options)
            driver.maximize_window()

            print('>>> Webdriver Opera iniciado...')

            return driver
        except Exception as e:
            print('error ', e)
    elif browser == 'Firefox':
        try:
            options = webdriver.FirefoxOptions()
            options.add_argument('--disable-gpu')
            options.set_preference('browser.download.dir', os.path.join(os.getcwd(), 'geracao'))
            options.set_preference('browser.download.folderList', 2)
            options.add_argument("--window-size=1920x1080")

            # Use o GeckoDriverManager para obter o executável do driver automaticamente
            driver_manager = GeckoDriverManager()
            driver = webdriver.Firefox(executable_path=driver_manager.install(), options=options)
            driver.maximize_window()

            print('>>> Webdriver Mozilla Firefox iniciado...')

            return driver
        except Exception as e:
            print('error ', e)
    
    else:
        print('Navegador não suportado:', browser)
        return None

