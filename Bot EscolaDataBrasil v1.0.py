# -*- coding: utf-8 -*-
import os.path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from xlwt import Workbook, easyxf
from xlrd import open_workbook,XL_CELL_TEXT
from unicodedata import normalize
from os.path import join, dirname, abspath, isfile
from xlutils.copy import copy
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException
import time 

def main():
    titleStyle = easyxf('alignment: horizontal center;' 'font:bold True;')
    colStyle = easyxf('alignment: horizontal center;')
    if os.path.isfile("Resultado Busca Escolas.xls"):
        rb = open_workbook("Resultado Busca Escolas.xls")
        count = rb.sheets()[0].nrows
        book = copy(rb)
        lista = book.get_sheet(0)
    else:    
        book = Workbook(encoding="UTF-8")
        lista = book.add_sheet('lista')
        lista.write(0,0,'Nome',titleStyle)
        lista.write(0,1,'Email',titleStyle)
        lista.write(0,2,'Dependencia',titleStyle)
        lista.write(0,3,'UF',titleStyle)
        lista.write(0,4,'Cidade',titleStyle)
        lista.write(0,5,'CEP',titleStyle)
        lista.write(0,6,'Endereco',titleStyle)
        lista.write(0,7,'Numero',titleStyle)
        lista.write(0,8,'Complemento',titleStyle)
        lista.write(0,9,'Bairro',titleStyle)
        lista.write(0,10,'DDD',titleStyle)
        lista.write(0,11,'Telefone1',titleStyle)
        lista.write(0,12,'Telefone2',titleStyle)
        lista.write(0,13,'Telefone3',titleStyle)
        lista.col(0).width = 15000
        lista.col(1).width = 15000
        lista.col(2).width = 5000
        lista.col(3).width = 1000
        lista.col(4).width = 5000
        lista.col(5).width = 5000
        lista.col(6).width = 15000
        lista.col(7).width = 3000
        lista.col(8).width = 5000
        lista.col(9).width = 7500
        lista.col(10).width = 1000
        lista.col(11).width = 5000
        lista.col(12).width = 5000
        lista.col(13).width = 5000
        count = 1
    browser = webdriver.Firefox()
    url = "http://www.dataescolabrasil.inep.gov.br/dataEscolaBrasil/home.seam"
    browser.get(url)
    browser.find_element_by_id("depAdmFederal").click()
    browser.find_element_by_id("depAdmEstadual").click()
    browser.find_element_by_id("depAdmMunicipal").click()
    browser.find_element_by_id("depAdmPrivada").click()
    browser.find_element_by_id("situacaoEmAtividade").click()
    browser.find_element_by_id("modalidadeRegular").click()
    browser.find_element_by_id("modalidadeEspecial").click()
    browser.find_element_by_id("etapaEducacaoInfantil").click()
    browser.find_element_by_id("etapaEnsinoFundamental8anos").click()
    browser.find_element_by_id("etapaEnsinoFundamental9anos").click()
    browser.find_element_by_id("etapaEnsinoMedio").click()
    browser.find_element_by_id("pesquisar").click()
    cellnum=(count%10)-1
    page=count/10
    print "Page number: "
    print page
    if cellnum==-1:
        cellnum=9
    print "Page index: "
    print cellnum

    browser.execute_script("document.getElementsByTagName(\"td\")[92].onclick = function onclick(event) { Event.fire(this, 'rich:datascroller:onscroll', {'page': '"+str(page) +"'});}")
    browser.execute_script("document.getElementsByTagName(\"td\")[92].onclick();")
    print "waiting 30 seconds"
    browser.implicitly_wait(1)#seconds 
    time.sleep(1)
    print "time is over"
    while (browser.find_elements(By.XPATH, '//td[@class=" dr-dscr-button rich-datascr-button"]')[-2].text == u'\xbb'):
        for i in range(5,len(browser.find_elements(By.XPATH, '//a[@href="#"]')),7):
            if cellnum == 0:
                print "\n\nESCOLA"
                #print i
                #print browser.find_elements(By.XPATH, '//a[@href="#"]')[i].text
                browser.find_elements(By.XPATH, '//a[@href="#"]')[i].click()
                lista.write(count,0,browser.find_element_by_id("dad_noEntidadeDecorate:dad_noEntidade").text,colStyle)
                lista.write(count,1,browser.find_element_by_id("dad_emailDecorate:dad_email").text,colStyle)
                lista.write(count,2,browser.find_element_by_id("dad_dependenciaAdmDecorate:dad_dependenciaAdm").text,colStyle)
                lista.write(count,3,browser.find_element_by_id("dad_noUfDecorate:dad_noUf").text,colStyle)
                lista.write(count,4,browser.find_element_by_id("dad_noMunicipioDecorate:dad_noMunicipio").text,colStyle)
                lista.write(count,5,browser.find_element_by_id("dad_cepDecorate:dad_cep").text,colStyle)
                lista.write(count,6,browser.find_element_by_id("dad_enderecoDecorate:dad_endereco").text,colStyle)
                lista.write(count,7,browser.find_element_by_id("dad_numeroDecorate:dad_numero").text,colStyle)
                lista.write(count,8,browser.find_element_by_id("dad_complementoDecorate:dad_complemento").text,colStyle)
                lista.write(count,9,browser.find_element_by_id("dad_bairroDecorate:dad_bairro").text,colStyle)
                lista.write(count,10,browser.find_element_by_id("dad_numDddDecorate:dad_numDdd").text,colStyle)
                lista.write(count,11,browser.find_element_by_id("dad_numTelefoneDecorate:dad_numTelefone").text,colStyle)
                lista.write(count,12,browser.find_element_by_id("dad_numTelefonePublico1Decorate:dad_numTelefonePublico1").text,colStyle)
                lista.write(count,13,browser.find_element_by_id("dad_numTelefonePublico2Decorate:dad_numTelefonePublico2").text ,colStyle)  

                print browser.find_element_by_id("dad_noEntidadeDecorate:dad_noEntidade").text
                print browser.find_element_by_id("dad_emailDecorate:dad_email").text
                print browser.find_element_by_id("dad_dependenciaAdmDecorate:dad_dependenciaAdm").text
                print browser.find_element_by_id("dad_noUfDecorate:dad_noUf").text
                print browser.find_element_by_id("dad_noMunicipioDecorate:dad_noMunicipio").text
                #print browser.find_element_by_id("dad_noDistritoDecorate:dad_noDistrito").text
                print browser.find_element_by_id("dad_cepDecorate:dad_cep").text
                print browser.find_element_by_id("dad_enderecoDecorate:dad_endereco").text
                print browser.find_element_by_id("dad_numeroDecorate:dad_numero").text
                print browser.find_element_by_id("dad_complementoDecorate:dad_complemento").text
                print browser.find_element_by_id("dad_bairroDecorate:dad_bairro").text
                print browser.find_element_by_id("dad_numDddDecorate:dad_numDdd").text
                print browser.find_element_by_id("dad_numTelefoneDecorate:dad_numTelefone").text    
                print browser.find_element_by_id("dad_numTelefonePublico1Decorate:dad_numTelefonePublico1").text
                print browser.find_element_by_id("dad_numTelefonePublico2Decorate:dad_numTelefonePublico2").text 
                count = count+1
                book.save('Resultado Busca Escolas.xls')
                browser.find_element_by_id("j_id207_lbl").click()   
            else:
                #print "cellnum =" + str(-cellnum) + "\nposicao de i " + str(i)
                cellnum=cellnum-1
        browser.find_elements(By.XPATH, '//td[@class=" dr-dscr-button rich-datascr-button"]')[-2].click()
        #browser.implicitly_wait(20)#seconds 
        time.sleep(1)



if __name__ == "__main__":
    main()