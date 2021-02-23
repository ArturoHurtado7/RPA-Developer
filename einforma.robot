*** Settings ***
Library     SeleniumLibrary
Library     Collections
Library     OperatingSystem
Library     String
Library     RPA.Tables
Library     RPA.Excel.Files

*** Variables ***
${browser}      chrome
${url}          https://www.einforma.co/buscador-empresas-empresarios
@{LinkItems}

*** Test Cases ***
SearchBusinessInfo
    Initialize Browser
    Search Args
    ${count}=   get element count   xpath://select[@id='PROVINCIA']
    Run Keyword If  (${count}>0)    Select Provincia List
    Run Keyword If  (${count}>0)    Get All Links
    Create Workbook  Info_Empresas.xlsx
    Run Keyword If  (${count}>0)    Get All Empresas
    Save Workbook
    close browser

*** Keywords ***
Initialize Browser
    open browser    ${url}  ${browser}  
    maximize browser window

*** Keywords ***
Search Args
    input text   id:search2     ${busqueda}
    click element   xpath://input[@id='boton_buscador_nacional']

*** Keywords ***
Select Provincia List
    select from list by Value  PROVINCIA   ${departamento}

*** Keywords ***
Get All Empresas
    ${listCount}    Get Length    ${LinkItems}
    FOR   ${i}    IN RANGE    0   ${listCount}
        ${LinkToClic}=  Set Variable  https://www.einforma.co${LinkItems[${i}]}
        log to console      ${LinkToClic}
        Set Global Variable      ${LinkToClic}
        Go To    ${LinkToClic}
        Get All Information     ${i}
        Write File
        sleep   5s
    END

*** Keywords ***
Razon Social
    [Arguments]    ${i}
    ${Razon_Social}=    Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Razon_Social}

*** Keywords ***
Forma Juridica
    [Arguments]    ${i}
    ${Forma_Juridica}=  Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Forma_Juridica}

*** Keywords ***
Departamento
    [Arguments]    ${i}
    ${Departamento}=    Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Departamento}

*** Keywords ***
Actividad CIIU
    [Arguments]    ${i}
    ${Actividad_CIIU}=  Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Actividad_CIIU}

*** Keywords ***
Fecha Constitucion
    [Arguments]    ${i}
    ${Fecha_Constitucion}=  Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Fecha_Constitucion}

*** Keywords ***
Fecha Ultimo Dato
    [Arguments]    ${i}
    ${Fecha_Ultimo_Dato}=   Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Fecha_Ultimo_Dato}

*** Keywords ***
Fecha Camara
    [Arguments]    ${i}
    ${Fecha_Camara}=   Get Text    xpath://table/tbody/tr[${i}]/td[2]
    Set Global Variable      ${Fecha_Camara}

*** Keywords ***
Get All Information
    [Arguments]    ${i}

    ${Razon_Social}=    Set Variable   NA
    ${Forma_Juridica}=  Set Variable   NA
    ${Departamento}=    Set Variable   NA
    ${Actividad_CIIU}=  Set Variable   NA
    ${Fecha_Constitucion}=  Set Variable   NA
    ${Fecha_Ultimo_Dato}=   Set Variable   NA
    ${Fecha_Camara}=    Set Variable   NA
    ${ScreenPath}=  Set Variable  ScreenShots/${busqueda}_screenshot${i}.png
    
    Set Global Variable      ${Razon_Social}
    Set Global Variable      ${Forma_Juridica}
    Set Global Variable      ${Departamento}
    Set Global Variable      ${Actividad_CIIU}
    Set Global Variable      ${Fecha_Constitucion}
    Set Global Variable      ${Fecha_Ultimo_Dato}
    Set Global Variable      ${Fecha_Camara}
    Set Global Variable      ${ScreenPath}

    ${count}=   get element count   xpath://html/body/div[2]/div[2]/div/div/div[2]/div/div[1]/div[2]/table/tbody/tr
    FOR     ${j}    IN RANGE    1   ${count}+1
        ${title}=    Get Text    xpath://table/tbody/tr[${j}]/td[1]
        log to console      ${title}
        Run Keyword If  ('${title}' == 'Razón Social:')     Razon Social        ${j}
        Run Keyword If  ('${title}' == 'Forma Jurídica:')   Forma Juridica      ${j}
        Run Keyword If  ('${title}' == 'Departamento:')     Departamento        ${j}
        Run Keyword If  ('${title}' == 'Actividad CIIU:')   Actividad CIIU      ${j}
        Run Keyword If  ('${title}' == 'Fecha Constitución:')   Fecha Constitucion      ${j}
        Run Keyword If  ('${title}' == 'Fecha Último Dato:')    Fecha Ultimo Dato       ${j}
        Run Keyword If  ('${title}' == 'Fecha Actualización Cámara Comercio:')  Fecha Camara    ${j}
    END

    ${ScreenPath}=  Set Variable  ScreenShots/${busqueda}_${Razon_Social}_${i}.png
    capture element screenshot  xpath://tbody   ${ScreenPath}

*** Keywords ***
Get All Links
    FOR     ${i}    IN RANGE    99999
        Get Page Links
        ${listCount}    Get Length    ${LinkItems}
        Exit For Loop If    ${listCount} == ${n}
        ${count}=   get element count   xpath://a[contains(text(),'siguiente >')]
        Run Keyword If  (${count}>0)    click element   xpath://a[contains(text(),'siguiente >')]
        Exit For Loop If    ${count} == 0
        sleep   5s
    END

*** Keywords ***
Get Page Links
    ${AllLinksCount}=   get element count   xpath://tbody/tr[contains(@class,"hover-change")]

    FOR     ${i}    IN RANGE    1   ${AllLinksCount}+1
        ${LinkText}=    Get Element Attribute   xpath:(//tbody/tr[contains(@class,"hover-change")])[${i}]   url
        ${listCount}    Get Length    ${LinkItems}
        Run Keyword If  (${listCount}<${n})     Append To List  ${LinkItems}    ${LinkText}
        Exit For Loop If    ${listCount} == ${n}
    END

*** Keywords ***
Write File
    &{row}=     Create Dictionary
    ...     Razon_Social    ${Razon_Social}
    ...     Forma_Juridica  ${Forma_Juridica}
    ...     Departamento    ${Departamento}
    ...     Actividad_CIIU  ${Actividad_CIIU}
    ...     Fecha_Constitucion  ${Fecha_Constitucion}
    ...     Fecha_Ultimo_Dato   ${Fecha_Ultimo_Dato}
    ...     Fecha_Camara    ${Fecha_Camara}
    ...     ScreenPath  ${ScreenPath}
    ...     LinkToClic  ${LinkToClic}
    Append Rows to Worksheet  ${row}  header=${TRUE}
