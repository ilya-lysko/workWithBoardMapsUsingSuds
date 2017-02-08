import sys
import re

import suds
from suds import client
from suds.sax.element import Element
from suds.wsse import *

import xlrd
from urllib.error import URLError

import logging

class ClientBM:
    
    serverURL = str()
    client = None
    currentInterfacenumberInArray = int()
    excelFile = None
    login = None
    password = None
    userToCreateAmount = None
    companyWorkWithShortName = None
    defaultEmail = 'demoboardmaps@yandex.ru'
    interfaces = ['CompanyManagementService','UserManagementService','CollegialBodyManagementService',
                 'MeetingManagementService','MeetingMemberManagementService','IssueManagementService'
                 ,'DecisionProjectManagementService','InvitedMember2IssueManagementService',
                  'SpokesPerson2IssueManagementService',
                  'MaterialManagementService','DocumentManagementService','InstructionManagementService']
    
#=========================================================
# МЕТОДЫ ДЛЯ ПОДКЛЮЧЕНИЯ, АВТОРИЗАЦИИ И СТАРТА РАБОТЫ

    def __init__(self, serverURL):
        self.serverURL = serverURL
        self.createLogFile()
        
    def startWorkWithInterface(self, interfaceNumberInArray):
        '''
        Отсчет interfaceNumberInArray начинается с 0
        '''
        try:
            self.currentInterfacenumberInArray = interfaceNumberInArray
            self.client = suds.client.Client(self.serverURL + "/PublicApi/" + self.interfaces[interfaceNumberInArray] + ".svc?wsdl")
            self.addNoteToLogFile('\n\nНачало работы с сервером %s' % self.serverURL)
        except URLError as e:
            self.addNoteToLogFile('\n\nСбой подключения к серверу %s' % self.serverURL, warning=True)
            self.addNoteToLogFile(e.args, warning=True)
            raise e;

    def setLoginAndPassword(self, login, password):
        self.login = login
        self.password = password

    def authorization(self):
        '''
        После запуска метода startWorkWithInterface
        '''
        security = Security()
        token = UsernameToken(self.login, self.password)
        security.tokens.append(token)
        self.client.set_options(wsse=security)
        try:
            self.client.service.Get()
            self.addNoteToLogFile('Успешная авторизация.')
        except WebFault as e:
            self.addNoteToLogFile('Неверный логин/пароль.', warning=True)
            raise(e)
        except Exception as e:
            self.addNoteToLogFile(e.args, warning=True)

#=========================================================
# МЕТОДЫ ДЛЯ ПОЛУЧЕНИЯ ДОПОЛНИТЕЛЬНОЙ ИНФОРМАЦИИ, ТРЕБУЕМОЙ ДЛЯ НЕКОТОРЫХ ДАЛЬНЕЙШИХ СЦЕНАРИЕВ

    def getCompanyIdByItsShortName(self):
        self.startWorkWithInterface(0)
        self.authorization()
        CompanySearchCriteriaDto = self.client.factory.create('ns0:CompanySearchCriteriaDto')
        CompanySearchCriteriaDto.ShortNameToken = self.companyWorkWithShortName
        try:
            companyInfo = self.client.service.Find(CompanySearchCriteriaDto)
            if companyInfo == '':
                raise Exception('Компании с таким именем нет.')
            return companyInfo.CompanyDto[0].Id
        except WebFault as e:
            print(e)

    def getCompanyShortName(self):
        return str(input('Введите короткое имя любой компании, для получения Id Холдинга: '))

    def getHoldingIdByCompanyShortName(self, companyShortName):
        '''
        Метод для вытягивания ID холдинга
        Входные параметры - короткое название компании (любой)
        '''
        self.startWorkWithInterface(0)
        self.authorization()
        CompanySearchCriteriaDto = self.client.factory.create('ns0:CompanySearchCriteriaDto')
        CompanySearchCriteriaDto.ShortNameToken = companyShortName
        try:
            companyInfo = self.client.service.Find(CompanySearchCriteriaDto)
            if companyInfo == '':
                raise Exception('Компании с таким именем нет.')
            return companyInfo.CompanyDto[0].Holding.Id
        except WebFault as e:
            print(e)

    def getUserIdByHisFI(self, userFI):
        '''
        Получение Id пользователя по его фамилии и имени.
        Формат userFI -- String. Преобразования с ней -- тут.
        Нужно для заполнения поля Id Председателя при создании КО.
        '''
        self.startWorkWithInterface(1)
        self.authorization();

        UserSearchCriteriaDto = t.client.factory.create('ns0:UserSearchCriteriaDto')
        UserSearchCriteriaDto.LastNameToken = userFI.split()[0]
        UserSearchCriteriaDto.FirstNameToken = userFI.split()[1]

        try: # Считаем, что результатом будет только один пользователь.
            userInfo = t.client.service.Find(q)
            return userInfo.UserDto[0].Id
        except WebFault as e:
            print(e)
   
#=========================================================
# МЕТОДЫ ДЛЯ ЛОГИРОВАНИЯ

    def createLogFile(self):
        logging.basicConfig(filename='log.log',level=logging.INFO, format='%(asctime)s %(message)s')

    def addNoteToLogFile(self, message, warning=False):
        if warning:
            logging.warning(message)
        else:
            logging.info(message)

#=========================================================
# МЕТОДЫ ДЛЯ РАБОТЫ С EXCEL ФАЙЛОМ

    def openExcelFile(self, filePathPlusName):
        try:
            self.excelFile = xlrd.open_workbook(filePathPlusName)
            self.addNoteToLogFile('Открыт excel файл %s' % filePathPlusName)
        except FileNotFoundError as e:
            self.addNoteToLogFile('Файл %s не найден.' % filePathPlusName, warning=True)
            raise e

    def readList(self, listName):
        '''
        Название листа с пользователями:
        рус - 'ПОЛЬЗОВАТЕЛИ'
        англ - 'USERS'

        Название листа с КО:
        рус - 'КО'
        англ -- надо уточнить

        Название листа с компанией:
        рус - 'О КОМПАНИИ'
        англ -- надо уточнить
        '''
        return self.excelFile.sheet_by_name(listName)
        
    
    def readInfoFromList(self, listFile, startInfoPosition, isCB=False):
        '''
        Создается массив со ВСЕЙ информацией о сущности из таблицы
        Немного нерационально, т.к. можно сразу тут создать словарь с нужными ключами
        Но не хочется так просто выкидывать часть информации -- вдруг пригодится потом

        Для пользователей startInfoPosition = 4, для компании - startInfoPosition = 3, для КО - startInfoPosition = 2.
        '''
        arrayWithInfo = []
        i = startInfoPosition
        i_y = 2
        if isCB:
            i_y = 3
        try:
            while listFile.row_values(i)[i_y] != '':
                arrayWithInfo.append(listFile.row_values(i))
                i += 1
        except IndexError as e:
            pass
        return arrayWithInfo

#=========================================================
# МЕТОДЫ ДЛЯ СОЗДАНИЯ USER

    def createUser(self, userInfo, usersCompanyId):
        '''
        Запуск этого метода только после запуска метода startWorkWithInterface с параметром 1 
        (и соотв., после метода authorization)
        userInfo - словарь с информацией о пользователе
        '''
        ArrayOfUserCreationCommandDto = self.client.factory.create('ns0:ArrayOfUserCreationCommandDto')
        UserCreationCommandDto = self.client.factory.create('ns0:ArrayOfUserCreationCommandDto.UserCreationCommandDto')
        Company = self.client.factory.create('ns0:ArrayOfUserCreationCommandDto.UserCreationCommandDto.Company')

        UserCreationCommandDto.BasicPassword = userInfo['BasicPassword']
        UserCreationCommandDto.BasicUsername = userInfo['BasicUsername']
        UserCreationCommandDto.Blocked = userInfo['Blocked']
        Company.Id = usersCompanyId
        UserCreationCommandDto.Company = Company
        if userInfo['Email'] == '':
            UserCreationCommandDto.Email = self.defaultEmail
        else:
            UserCreationCommandDto.Email = userInfo['Email']
        UserCreationCommandDto.FirstName = userInfo['FirstName']
        UserCreationCommandDto.LastName = userInfo['LastName']
        UserCreationCommandDto.Phone = userInfo['Phone']
        UserCreationCommandDto.Position = userInfo['Position']
        #UserCreationCommandDto.PhotoImageSource = "materials/default_picture.png" # не работает :(

        ArrayOfUserCreationCommandDto.UserCreationCommandDto.append(UserCreationCommandDto)
        
        try:
            getInfo = self.client.service.Create(ArrayOfUserCreationCommandDto)
            self.addNoteToLogFile('Создан пользователь. %s' % getInfo)
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

    def createSeveralUsers(self, arrayOfUserInfoDict, usersCompanyId):
        '''
        arrayOfUserInfoDict - массив словарей с информацией о пользователях для их создания
        '''
        for userInfo in arrayOfUserInfoDict:
            self.createUser(userInfo, usersCompanyId)
        self.userToCreateAmount = len(arrayOfUserInfoDict)

    def createArrayOfDictWithUsersInfo(self, arrayWithUsersInfo, defaultPassword):
        arrayOfDictWithUsersInfo = []
        for x in arrayWithUsersInfo:
            arrayOfDictWithUsersInfo.append({
                    'BasicPassword': defaultPassword, 'BasicUsername': x[14], 'Blocked': False,
                    'Company': None, 'Email': x[11], 'FirstName': x[2].split()[1], 
                    'LastName': x[2].split()[0], 'Phone': x[12], 'Position': x[9]
                 })
        return arrayOfDictWithUsersInfo

    def workWithUsersExcelController(self, excelFilePathPlusName, defaultPassword):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='ПОЛЬЗОВАТЕЛИ')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=4)
        return self.createArrayOfDictWithUsersInfo(arrayWithInfo, defaultPassword)

    def createUsersFromExcelController(self, excelFilePathPlusName, defaultPassword):
        '''
        Создание пользователей по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с пользователями, авторизацию и т.д.
        '''
        usersCompanyId = self.getCompanyIdByItsShortName()
        self.startWorkWithInterface(interfaceNumberInArray=1)
        self.authorization()
        arrayOfDictWithUsersInfo = self.workWithUsersExcelController(excelFilePathPlusName, defaultPassword)
        self.createSeveralUsers(arrayOfDictWithUsersInfo, usersCompanyId)

#=========================================================
# МЕТОДЫ ДЛЯ СОЗДАНИЯ КОМПАНИИ

    def createCompany(self, companyInfo, holdingId):
        '''
        Создании Компании, опираясь информацию из companyInfo
        Так же, требуется holdingId (дл определения этого параметра есть отдельный метод)
        '''
        ArrayOfCompanyCreationCommandDto = self.client.factory.create('ns0:ArrayOfCompanyCreationCommandDto')
        CompanyCreationCommandDto = self.client.factory.create('ns0:CompanyCreationCommandDto')
        IdentityDto = self.client.factory.create('ns0:IdentityDto')

        CompanyCreationCommandDto.AddressBuildingNumber = companyInfo['AddressBuildingNumber']
        CompanyCreationCommandDto.AddressCity = companyInfo['AddressCity']
        CompanyCreationCommandDto.AddressCountry = companyInfo['AddressCountry']
        CompanyCreationCommandDto.AddressIndex = companyInfo['AddressIndex']
        CompanyCreationCommandDto.Email = companyInfo['Email']
        CompanyCreationCommandDto.FullName = companyInfo['FullName']
        IdentityDto.Id = holdingId
        CompanyCreationCommandDto.Holding = IdentityDto
        CompanyCreationCommandDto.Phone = companyInfo['Phone']
        CompanyCreationCommandDto.PostBuildingNumber = companyInfo['PostBuildingNumber']
        CompanyCreationCommandDto.PostCity = companyInfo['PostCity']
        CompanyCreationCommandDto.PostCountry = companyInfo['PostCountry']
        CompanyCreationCommandDto.PostIndex = companyInfo['PostIndex']
        CompanyCreationCommandDto.ShortDescription = companyInfo['ShortDescription']
        CompanyCreationCommandDto.ShortName = companyInfo['ShortName']
        CompanyCreationCommandDto.UrlSite = companyInfo['UrlSite']

        ArrayOfCompanyCreationCommandDto.CompanyCreationCommandDto.append(CompanyCreationCommandDto)

        try:
            getInfo = self.client.service.Create(ArrayOfCompanyCreationCommandDto)
            self.addNoteToLogFile('Создана компания. %s' % getInfo)
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

    def createArrayWithCompanyInfo(self, arrayWithCompanyInfo):
        arrayOfDictWithCompanyInfo = []
        arrayOfDictWithCompanyInfo.append({
                    'AddressBuildingNumber': arrayWithCompanyInfo[10][2], 'AddressCity': arrayWithCompanyInfo[9][2],
                    'AddressCountry': arrayWithCompanyInfo[8][2], 'AddressIndex': int(arrayWithCompanyInfo[7][2]),
                    'Email': arrayWithCompanyInfo[4][2], 'FullName': arrayWithCompanyInfo[0][2],
                    'holdingId': None, 'Phone': arrayWithCompanyInfo[5][2],
                    'PostBuildingNumber': arrayWithCompanyInfo[15][2], 'PostCity': arrayWithCompanyInfo[14][2],
                    'PostCountry': arrayWithCompanyInfo[13][2], 'PostIndex': int(arrayWithCompanyInfo[12][2]),
                    'ShortDescription': arrayWithCompanyInfo[2][2], 'ShortName': arrayWithCompanyInfo[1][2],
                    'UrlSite': arrayWithCompanyInfo[3][2]
                })
        self.companyWorkWithShortName = arrayWithCompanyInfo[1][2]
        return arrayOfDictWithCompanyInfo

    def workWithCompanyExcelController(self, excelFilePathPlusName):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='О КОМПАНИИ')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=3)
        return self.createArrayWithCompanyInfo(arrayWithInfo)

    def createCompanyFromExcelController(self, excelFilePathPlusName):
        '''
        Создание компаниии по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с компаниями, авторизацию и т.д.
        '''
        companyShortName = self.getCompanyShortName()
        holdingId = self.getHoldingIdByCompanyShortName(companyShortName)
        self.startWorkWithInterface(interfaceNumberInArray=0)
        self.authorization()
        arrayOfDictWithCompanyInfo = self.workWithCompanyExcelController(excelFilePathPlusName)
        self.createCompany(arrayOfDictWithCompanyInfo[0], holdingId)

#=========================================================
# МЕТОДЫ ДЛЯ СОЗДАНИЯ КОЛЛЕГИАЛЬНОГО ОРГАНА {НЕ ДОДЕЛАН. НЕСООТВЕТСТВИЕ API И ДОКУМЕНТАЦИИ К НЕМУ}

    def createCollegialBody(self, collegialBodyInfo, holdingId, CBCompanyId, dictWithCBUserRoles):
        '''
        Создание Коллегиального Органа, опираясь на информацию из входного значения - collegialBodyInfo
        После создания пользователей лучше, т.к. тут уже определяется секретарь
        Соответственно, желательно, чтобы пользователь, которому планируется присвоить роль секретаря был создан
        '''
        # Получим сначала ФИО председателя и секретаря данного КО
        CBHeadAndSecretaryInfo = dictWithCBUserRoles[collegialBodyInfo['FullName']]

        for k,v in CBHeadAndSecretaryInfo.items():
            if v == 'СЕК':
                secretaryOfCBName = k
            elif v == 'ПРЕД':
                headOfCBName = k

        # Получим Id председателя и секретаря данного КО, основываясь на их ФИО
        headOfCBId = getUserIdByHisFI(headOfCBName)
        secretaryOfCBId = getUserIdByHisFI(secretaryOfCBName)

        ArrayOfCollegialBodyCreationCommandDto = client.factory.create('ns0:ArrayOfCollegialBodyCreationCommandDto')
        CollegialBodyCreationCommandDto = client.factory.create('ns0:CollegialBodyCreationCommandDto')
        AttendanceTypeEnumDto = client.factory.create('ns0:AttendanceTypeEnumDto')
        CollegialBodyTypeEnumDto = client.factory.create('ns0:CollegialBodyTypeEnumDto')
        IdentityDtoCompany = client.factory.create('ns0:IdentityDto')
        LdapUserIdentityDtoHeadOf = client.factory.create('ns0:LdapUserIdentityDto')
        IdentityDtoParent = client.factory.create('ns0:IdentityDto')
        LdapUserIdentityDtoQM = client.factory.create('ns0:LdapUserIdentityDto')

        IdentityDtoParent.Id = collegialBodyInfo['ParentId']
        CollegialBodyCreationCommandDto.Parent = IdentityDtoParent
        LdapUserIdentityDtoHeadOf.Id = headOfCBId
        #LdapUserIdentityDtoHeadOf.LdapUsername = collegialBodyInfo['HeadOfName']
        CollegialBodyCreationCommandDto.HeadOf = LdapUserIdentityDtoHeadOf
        IdentityDtoCompany.Id = CBCompanyId
        CollegialBodyCreationCommandDto.Company = IdentityDtoCompany
        # поля уже заполнены для полей ниже, которые закоменчены
        # РАЗОБРАТЬСЯ
        #CollegialBodyTypeEnumDto.Executive = collegialBodyInfo['Executive']
        #CollegialBodyTypeEnumDto.ManagementBody = collegialBodyInfo['ManagementBody']
        #CollegialBodyTypeEnumDto.NotCorporate = collegialBodyInfo['NotCorporate']
        #CollegialBodyTypeEnumDto.NotExecutive = collegialBodyInfo['NotExecutive']
        #CollegialBodyTypeEnumDto.State = collegialBodyInfo['State']
        #ollegialBodyCreationCommandDto.CollegialBodyType = CollegialBodyTypeEnumDto
        #AttendanceTypeEnumDto.0 = collegialBodyInfo['0'] 
        #AttendanceTypeEnumDto.1 = collegialBodyInfo['1']
        CollegialBodyCreationCommandDto.Attendance = AttendanceTypeEnumDto
        CollegialBodyCreationCommandDto.FullName = collegialBodyInfo['FullName']
        CollegialBodyCreationCommandDto.Order = collegialBodyInfo['Order']
        CollegialBodyCreationCommandDto.QualifiedMajority = collegialBodyInfo['QualifiedMajority']
        CollegialBodyCreationCommandDto.Secretary = secretaryOfCBId
        CollegialBodyCreationCommandDto.ShortDescription = collegialBodyInfo['ShortDescription']
        CollegialBodyCreationCommandDto.ShortName = collegialBodyInfo['ShortName']

        try:
            getInfo = self.client.service.Create(ArrayOfCollegialBodyCreationCommandDto)
            self.addNoteToLogFile('Создан колегиальный орган. %s' % getInfo)
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

    def getHeadOfAndSecretary(self, CBAmount):
        '''
        Создание структуры данных с ролями всех пользователей во всех КО
        '''
        def createDictWithCBInfoWithStructure(self, arrayWithCBRolesWithNoStructure, CBAmount):
            '''
            Возвращает "словарь словарей" с информацией о ролях всех пользователей во всех КО.
            Конкретно для создания КО есть лишняя информация.
            '''
            dictWithCBUserRoles = {}
            for i in range(CBAmount):
                d_ = {}
                for j in range(self.userToCreateAmount):
                    d_.update({arrayWithCBRolesWithNoStructure[j+1][1]: arrayWithCBRolesWithNoStructure[j+1][2+i]})
                dictWithCBUserRoles.update({arrayWithCBRolesWithNoStructure[0][i+2]: d_})
            return dictWithCBUserRoles

        def getUsefulFormatFromDictWithCBUserRoles(self, dictWithCBUserRoles):
            '''
            "Конвертирование" "словаря словарей" в формат:
                - убрать участников (всевозможных типов)
                - вместо имен пользователей -- Id пользователя
            '''
            for i in dictWithCBUserRoles.keys():
                b = []
                for j in dictWithCBUserRoles[i]:
                    if dictWithCBUserRoles[i][j] not in ['ПРЕД','СЕК']:
                        b.append(j)
                for k in b:
                    dictWithCBUserRoles[i].pop(k)
            return dictWithCBUserRoles

        listWithCBInfo = self.readList('РОЛИ')
        arrayWithCBRolesWithNoStructure = self.readInfoFromList(listWithCBInfo, startInfoPosition=2, isCB=True)
        dictWithCBUserRoles = self.createArrayWithCBInfoWithStructure(arrayWithCBRolesWithNoStructure, CBAmount)
        dictWithCBUserRoles = getUsefulFormatFromDictWithCBUserRoles(dictWithCBUserRoles)
        return dictWithCBUserRoles

    def createSeveralCollegialBodies(arrayOfDictWithCBInfo, CBCompanyId, dictWithCBUserRoles):
        for CBInfo in arrayOfDictWithCBInfo:
            self.createCollegialBody(CBInfo, CBCompanyId, dictWithCBUserRoles)

    def createArrayOfDictWithCBInfo(self, arrayWithCBInfo, qualifiedCBUsersCount=4):
        arrayOfDictWithCBInfo = []
        j = 1
        for x in arrayWithCBInfo:
            arrayOfDictWithCBInfo.append({
                    'FullName': x[2], 'ShortName': x[4], 'ParentId': None, 'HeadOfId': None,
                    'Order': j, 'CompanyId': None, 'QualifiedMajority': qualifiedCBUsersCount,
                    'Secretary': None, 'ShortDescription': x[9]

                })
            j += 1
        return arrayOfDictWithCBInfo, len(arrayOfDictWithCBInfo)

    def workWithCBExcelController(self, excelFilePathPlusName):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='КО')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=4)
        return self.createArrayOfDictWithCBInfo(arrayWithInfo)

    def createCBFromExcelController(self, excelFilePathPlusName):
        '''
        Создание КО по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с КО, авторизацию и т.д.
        '''
        CBCompanyId = self.getCompanyIdByItsShortName(companyShortName)
        self.startWorkWithInterface(interfaceNumberInArray=2)
        self.authorization()
        arrayOfDictWithCBInfo, CBAmount = self.workWithCBExcelController(excelFilePathPlusName)
        dictWithCBUserRoles = getHeadOfAndSecretary(CBAmount)
        self.createSeveralCollegialBodies(arrayOfDictWithCBInfo, CBCompanyId, dictWithCBUserRoles)




'''
ПАРАМЕТРЫ СКРИПТА (ПОРЯДОК ВАЖЕН):
- адрес сервера BM
- путь к excel файлу (с расширением и именем самого файла)
- логин для входа на сервер BM
- пароль для входа на сервер BM
- желаемый пароль для создаваемых пользователей
'''


if __name__ == '__main__':
    try:
        sys.argv = sys.argv[1:]
        url = sys.argv[0]
        clientBM = ClientBM(url)
        excelFilePathPlusName = sys.argv[1]
        login = sys.argv[2]
        password = sys.argv[3]
        clientBM.setLoginAndPassword(login, password)
        defaultPassword = sys.argv[4]

        clientBM.createCompanyFromExcelController(excelFilePathPlusName)
        clientBM.createUsersFromExcelController(excelFilePathPlusName, defaultPassword)

    except Exception as e:
        clientBM.addNoteToLogFile(e.args, warning=True)
