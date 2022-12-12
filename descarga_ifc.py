from selenium import webdriver
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import os



def descarga(proceso,mes,ano,ruta):
    
    procesos=[]
    options = webdriver.ChromeOptions()
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    prefs = {"download.default_directory" : ruta}
    options.add_experimental_option("prefs",prefs)
    options.headless = True
    service = Service('E:\GitHub\Coordinador_Electrico\\2022\DBCD\BADX\webdriver\chromedriver.exe') 
    driver = webdriver.Chrome(service=service, options=options)
    #driver = webdriver.Chrome(options=options)

    
    driver.get('https://web.ifc.coordinador.cl/')

    time.sleep(1)

    driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/app-filtros-despliegue-publico/div[1]/mat-form-field[1]/div/div[1]/div').click()


    

  
    time.sleep(1)
    for i in range(1, 8):
        procesos.append(driver.find_element(By.XPATH,'//*[@id="cdk-overlay-0"]/div/mat-option[' + str(i) + ']').text)


    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="cdk-overlay-0"]/div/mat-option[' + str(procesos.index(proceso)+1) + ']').click()


    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/app-filtros-despliegue-publico/div[1]/mat-form-field[2]/div/div[1]/div').click()



    time.sleep(2)

    anos=[]
    for j in range(1, 5):
        anos.append(driver.find_element(By.XPATH,'//*[@id="cdk-overlay-1"]/div/mat-option[' + str(j) + ']').text)


    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="cdk-overlay-1"]/div/mat-option[' + str(anos.index(ano)+1) + ']').click()


    time.sleep(1)
    driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/app-filtros-despliegue-publico/div[1]/mat-form-field[3]/div/div[1]/div').click()
    meses=[]
    for j in range(1, 13):
        meses.append(driver.find_element(By.XPATH,'//*[@id="cdk-overlay-2"]/div/mat-option[' + str(j) + ']').text)

      
    time.sleep(2)
    if driver.find_element(By.XPATH,'//*[@id="cdk-overlay-2"]/div/mat-option[' + str(meses.index(mes)+1) + ']').get_attribute('aria-disabled')=='true':
        driver.quit()
        print(mes + '-'+ ano + ' del proceso '+ proceso + ' No existe. Proceso finalizado ...')
        time.sleep(1)
        return


    

    driver.find_element(By.XPATH,'//*[@id="cdk-overlay-2"]/div/mat-option[' + str(meses.index(mes)+1) + ']').click()
    time.sleep(2)
    
    try:
            driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[2]/div[2]/mat-paginator/div/div/div[2]/div')
    except:      
            driver.quit()
            print('No existen documentos para búsqueda realizada. Proceso finalizado ...')
            time.sleep(1)
            return

    

    ruta_proceso = ruta+'\\'+ str(proceso)
    ruta_ano = ruta_proceso +'\\'+ str(ano)
    ruta_mes = ruta_ano +'\\'+ str(mes)

    print(ruta_proceso)
    print(ruta_ano)
    print(ruta_mes)
    try:
        os.stat(ruta_proceso)
    except:
        os.mkdir(ruta_proceso)
        print('directorio del proceso ',proceso,' creado ... ', ruta_proceso)
    else:
        print('directorio del proceso ',proceso,' seleccionado...  ', ruta_proceso)
    
    try:
        os.stat(ruta_ano)
    except:
        os.mkdir(ruta_ano)
        print('directorio del año ',ano,' del proceso ',proceso,' creado... ', ruta_ano)
    else:
        print('directorio del año ',ano,' del proceso ',proceso,' seleccionado... ', ruta_ano)
    
    try:
        os.stat(ruta_mes)
    except:
        os.mkdir(ruta_mes)
        print('directorio del mes ',mes ,' del ',ano,' del proceso ',proceso,' creado... ', ruta_mes)


    time.sleep(5)
 

    driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[2]/div[2]/mat-paginator/div/div/div[1]/mat-form-field').click()
    largo=driver.find_elements(By.XPATH,'//*[@id="cdk-overlay-3"]/div/mat-option/span[@class="mat-option-text"]')
  

    elemento=len(largo)
    time.sleep(2)
    driver.find_element(By.XPATH,'//*[@id="cdk-overlay-3"]/div/mat-option['+str(elemento)+']').click()
    time.sleep(5)

    paginas=driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[2]/div[2]/mat-paginator/div/div/div[2]/div').text
    time.sleep(2)
    

    registros=int(paginas[paginas.index('of ')+3:])
    registrospp=int(paginas[paginas.index('- ')+2:paginas.index(' of')])
    registros_extra=registros%registrospp
  
    totalp=registros//registrospp + (1 if registros%registrospp>0 else 0)



    params = { 'behavior': 'allow', 'downloadPath': ruta_mes }
    driver.execute_cdp_cmd('Page.setDownloadBehavior', params)

    if totalp==1:
        registros_extra=registrospp
    time.sleep(3)

    for i in range (1,totalp+1) :
        if i!=totalp:
            for j in range (1,registrospp) :
                version=driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[1]/mat-expansion-panel['+str(j)+']/mat-expansion-panel-header/span/div[3]').text
                archivo_a_revisar=driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[1]/mat-expansion-panel['+str(j)+']/mat-expansion-panel-header/span/div[6]').text
                if version[0:2]=="VE" and not(os.path.isfile(ruta_mes + '\\' + archivo_a_revisar)):
                    driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[1]/mat-expansion-panel['+str(j)+']/mat-expansion-panel-header/span/div[7]/a').click()
                    print(ruta_mes+'\\'+archivo_a_revisar, '...descargado.')
                elif version[0:2]=="VE":
                    print('archivo ', archivo_a_revisar, 'ya existe. No descargado...')
                time.sleep(1)
        else:
            for j in range (1,registros_extra) :
                version=driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[1]/mat-expansion-panel['+str(j)+']/mat-expansion-panel-header/span/div[3]').text
                archivo_a_revisar=driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[1]/mat-expansion-panel['+str(j)+']/mat-expansion-panel-header/span/div[6]').text
                if version[0:2]=="VE" and not(os.path.isfile(ruta_mes + '\\' + archivo_a_revisar)):
                    driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[1]/mat-expansion-panel['+str(j)+']/mat-expansion-panel-header/span/div[7]/a').click()
                    print(ruta_mes+'\\'+archivo_a_revisar, '...descargado.')
                elif version[0:2]=="VE":
                    print('archivo ', archivo_a_revisar, 'ya existe. No descargado...')
                time.sleep(1)
        driver.find_element(By.XPATH,'//*[@id="mat-tab-content-0-0"]/div/app-despliegue-publico/div/mat-card/div[2]/div[2]/mat-paginator/div/div/div[2]/button[2]').click()
        time.sleep(1)

    time.sleep(3)

    time.sleep(3)

    time.sleep(6)
    
    driver.quit()
    print('... archivos descargados...')
