import logging
from urllib.error import URLError

import openpyxl
import suds
import xlrd
from suds import client
from suds.wsse import *


class ClientBM:
    serverURL = str()
    client = None
    currentInterfacenumberInArray = int()
    excelFile = None
    login = None
    excelFileForWriting = None
    guidDictQueue = {}
    password = None
    userToCreateAmount = None
    companyWorkWithId = None
    defaultEmail = 'demoboardmaps@yandex.ru'
    interfaces = ['CompanyManagementService', 'UserManagementService', 'CollegialBodyManagementService',
                  'MeetingManagementService', 'MeetingMemberManagementService', 'IssueManagementService',
                'DecisionProjectManagementService', 'InvitedMember2IssueManagementService',
                  'SpokesPerson2IssueManagementService',
                  'MaterialManagementService', 'DocumentManagementService', 'InstructionManagementService']

    # =========================================================
    # МЕТОДЫ ДЛЯ ПОДКЛЮЧЕНИЯ, АВТОРИЗАЦИИ И СТАРТА РАБОТЫ C ОПРЕДЕЛЕННЫМ СЕРВИСОМ

    def __init__(self, serverURL):
        self.serverURL = serverURL
        self.createLogFile()

    def startWorkWithInterface(self, interfaceNumberInArray):
        '''
        Отсчет interfaceNumberInArray начинается с 0
        '''
        try:
            self.currentInterfacenumberInArray = interfaceNumberInArray
            self.client = suds.client.Client(
                self.serverURL + "/PublicApi/" + self.interfaces[interfaceNumberInArray] + ".svc?wsdl")
            self.addNoteToLogFile('\n\nНачало работы интерфейсом' + self.interfaces[interfaceNumberInArray] + ' сервера %s' % self.serverURL)
        except URLError as e:
            self.addNoteToLogFile('\n\nСбой подключения к серверу %s' % self.serverURL, warning=True)
            self.addNoteToLogFile(e.args, warning=True)
            raise e

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
            raise (e)
        except Exception as e:
            self.addNoteToLogFile(e.args, warning=True)

    # =========================================================
    # МЕТОДЫ ДЛЯ ПОЛУЧЕНИЯ ДОПОЛНИТЕЛЬНОЙ ИНФОРМАЦИИ, ТРЕБУЕМОЙ ДЛЯ НЕКОТОРЫХ ДАЛЬНЕЙШИХ СЦЕНАРИЕВ

    def getCompanyIdByItsShortName(self, companyWorkWithShortName):
        self.startWorkWithInterface(0)
        self.authorization()
        CompanySearchCriteriaDto = self.client.factory.create('ns0:CompanySearchCriteriaDto')
        CompanySearchCriteriaDto.ShortNameToken = companyWorkWithShortName
        try:
            companyInfo = self.client.service.Find(CompanySearchCriteriaDto)
            if companyInfo == '':
                raise Exception('Компании с таким именем нет.')
            return companyInfo.CompanyDto[0].Id
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

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
            self.addNoteToLogFile(e.args, warning=True)

    def getUserIdByHisFI(self, userLastName):
        '''
        Получение Id пользователя по его фамилии и имени.
        Формат userFI -- String. Преобразования с ней -- тут.
        Нужно для заполнения поля Id Председателя при создании КО.
        '''
        self.startWorkWithInterface(1)
        self.authorization()

        UserSearchCriteriaDto = self.client.factory.create('ns0:UserSearchCriteriaDto')
        UserSearchCriteriaDto.LastNameToken = userLastName

        try:  # Считаем, что результатом будет только один пользователь.
            userInfo = self.client.service.Find(UserSearchCriteriaDto)
            return userInfo.UserDto[0].Id
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

    # =========================================================
    # МЕТОДЫ ДЛЯ ЛОГИРОВАНИЯ

    def createLogFile(self):
        logging.basicConfig(filename='log.log', level=logging.INFO, format='%(asctime)s %(message)s')

    def addNoteToLogFile(self, message, warning=False):
        if warning:
            logging.warning(message)
            #print(message)
        else:
            logging.info(message)
            #print("WARNING: " + message)

    # =========================================================
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

    def readInfoFromList(self, listFile, startInfoPosition, isCBRoles=False):
        '''
        Создается массив со ВСЕЙ информацией о сущности из таблицы
        Немного нерационально, т.к. можно сразу тут создать словарь с нужными ключами
        Но не хочется так просто выкидывать часть информации -- вдруг пригодится потом
        '''
        arrayWithInfo = []
        i = startInfoPosition
        i_y = 2 - isCBRoles
        try:
            while listFile.row_values(i)[i_y] != '':
                arrayWithInfo.append(listFile.row_values(i))
                i += 1
        except IndexError as e:
            pass
        return arrayWithInfo

    def writeGuidToExcel(self):
        '''
        Записывает GUID созданного объекта в колонку в соотв. листе excel файла

        :param excelListName:
        :param guid:
        :return:
        '''
        excelFile = openpyxl.load_workbook(excelFilePathPlusName)

        for excelSheetName, guidRowDict in self.guidDictQueue.items():
            excelSheet = excelFile[excelSheetName]
            if excelSheetName == "О КОМПАНИИ":
                columnNumber = 5
            elif excelSheetName == "ПОЛЬЗОВАТЕЛИ":
                columnNumber = 18
            elif excelSheetName == "КО":
                columnNumber = 16
            else:
                raise Exception("Unknown sheet name")
            for k, v in guidRowDict.items():
                excelSheet.cell(row=k, column=columnNumber).value = v

        excelFile.save(excelFilePathPlusName)

    def addGuidDictToQueue(self, excelSheetName, guidRowDict):
        self.guidDictQueue.update({excelSheetName: guidRowDict})

    # =========================================================
    # МЕТОДЫ ДЛЯ СОЗДАНИЯ ПОЛЬЗОВАТЕЛЕЙ

    def createSeveralUsers(self, arrayOfUserInfoDict, usersCompanyId):
        '''
        Запуск этого метода только после запуска метода startWorkWithInterface с параметром 1
        (и соотв., после метода authorization)
        arrayOfUserInfoDict - массив словарей с информацией о пользователях для их создания
        Создание все равно происходит по одному объекту пользователя из-за необходимости валидации
        '''

        guidRowDict = {}
        rowCounter = 7

        self.userToCreateAmount = len(arrayOfUserInfoDict)

        for userInfo in arrayOfUserInfoDict:
            if (userInfo.get('GUID') != None and workStrategyFlag == 'i') or \
                    (userInfo.get('GUID') == None and workStrategyFlag == 'u'):
                continue

            if userInfo.get('GUID') != None:
                actionType = "Update"
            else:
                actionType = "Creation"

            ArrayOfUserActionCommandDto = self.client.factory.create('ns0:ArrayOfUser' + actionType + 'CommandDto')
            UserActionCommandDto = self.client.factory.create(
                'ns0:ArrayOfUser' + actionType + 'CommandDto.User' + actionType + 'CommandDto')
            Company = self.client.factory.create('ns0:ArrayOfUser' + actionType + 'CommandDto.User' + actionType + 'CommandDto.Company')

            if userInfo.get('GUID') != None:
                UserActionCommandDto.Id = userInfo.get('GUID')
            else:
                UserActionCommandDto.BasicPassword = userInfo['BasicPassword']
                UserActionCommandDto.BasicUsername = userInfo['BasicUsername']
                Company.Id = usersCompanyId
                UserActionCommandDto.Company = Company
            UserActionCommandDto.Blocked = False
            if userInfo['Email'] == '':
                UserActionCommandDto.Email = self.defaultEmail
            else:
                UserActionCommandDto.Email = userInfo['Email']
            UserActionCommandDto.FirstName = userInfo['FirstName']
            UserActionCommandDto.LastName = userInfo['LastName']
            UserActionCommandDto.MiddleName = userInfo['MiddleName']
            UserActionCommandDto.Phone = userInfo['Phone']
            UserActionCommandDto.Position = userInfo['Position']
            UserActionCommandDto.Bio = userInfo['Bio']
            # UserCreationCommandDto.PhotoImageSource = "materials/default_picture.png" # не работает :(

            if userInfo.get('GUID') != None:
                ArrayOfUserActionCommandDto.UserUpdateCommandDto.append(UserActionCommandDto)
            else:
                ArrayOfUserActionCommandDto.UserCreationCommandDto.append(UserActionCommandDto)

            try:
                if userInfo.get('GUID') != None:
                    info = self.client.service.Update(ArrayOfUserActionCommandDto)
                    self.addNoteToLogFile('Обновлен пользователь. %s' % info)
                else:
                    info = self.client.service.Create(ArrayOfUserActionCommandDto)
                    guidRowDict.update({rowCounter: info.UserDto[0].Id})
                    self.addNoteToLogFile('Создан пользователь. %s' % info)
                rowCounter += 1
            except WebFault as e:
                self.addNoteToLogFile(e.args, warning=True)

        self.addGuidDictToQueue("ПОЛЬЗОВАТЕЛИ", guidRowDict)

    def createArrayOfDictWithUsersInfo(self, arrayWithUsersInfo, defaultPassword):
        arrayOfDictWithUsersInfo = []
        for x in arrayWithUsersInfo:
            if x[17] != '':
                arrayOfDictWithUsersInfo.append({
                    'BasicPassword': defaultPassword, 'BasicUsername': x[12], 'Bio': x[15],
                    'Email': x[8], 'FirstName': x[3], 'GUID': x[17],
                    'LastName': x[2], 'MiddleName': x[4], 'Phone': x[9], 'Position': x[10]
                })
            else:
                arrayOfDictWithUsersInfo.append({
                    'BasicPassword': defaultPassword, 'BasicUsername': x[12], 'Bio': x[15],
                    'Email': x[8], 'FirstName': x[3],
                    'LastName': x[2], 'MiddleName': x[4], 'Phone': x[9], 'Position': x[10]
                })
        return arrayOfDictWithUsersInfo

    def workWithUsersExcelController(self, defaultPassword):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='ПОЛЬЗОВАТЕЛИ')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=6)
        return self.createArrayOfDictWithUsersInfo(arrayWithInfo, defaultPassword)

    def createUsersFromExcelController(self, defaultPassword):
        '''
        Создание пользователей по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с пользователями, авторизацию и т.д.
        '''
        usersCompanyId = self.companyWorkWithId
        self.startWorkWithInterface(interfaceNumberInArray=1)
        self.authorization()
        arrayOfDictWithUsersInfo = self.workWithUsersExcelController(defaultPassword)
        self.createSeveralUsers(arrayOfDictWithUsersInfo, usersCompanyId)

    # =========================================================
    # МЕТОДЫ ДЛЯ СОЗДАНИЯ КОМПАНИИ

    def createCompany(self, companyInfo, holdingId):
        '''
        Создании Компании, опираясь информацию из companyInfo
        Так же, требуется holdingId (для определения этого параметра есть отдельный метод)
        '''
        if (companyInfo.get('GUID') != None and workStrategyFlag == 'i') or \
            (companyInfo.get('GUID') == None and workStrategyFlag == 'u'):
            return

        if companyInfo.get('GUID') != None:
            actionType = "Update"
        else:
            actionType = "Creation"
        ArrayOfCompanyActionCommandDto = self.client.factory.create('ns0:ArrayOfCompany' + actionType +'CommandDto')
        CompanyActionCommandDto = self.client.factory.create('ns0:Company' + actionType + 'CommandDto')
        IdentityDto = self.client.factory.create('ns0:IdentityDto')

        if companyInfo.get('GUID') != None:
            CompanyActionCommandDto.Id = companyInfo.get('GUID')
        CompanyActionCommandDto.AddressBuildingNumber = companyInfo['AddressBuildingNumber']
        CompanyActionCommandDto.AddressCity = companyInfo['AddressCity']
        CompanyActionCommandDto.AddressCountry = companyInfo['AddressCountry']
        CompanyActionCommandDto.AddressIndex = companyInfo['AddressIndex']
        CompanyActionCommandDto.Email = companyInfo['Email']
        CompanyActionCommandDto.FullName = companyInfo['FullName']
        IdentityDto.Id = holdingId
        CompanyActionCommandDto.Holding = IdentityDto
        CompanyActionCommandDto.Phone = companyInfo['Phone']
        CompanyActionCommandDto.PostBuildingNumber = companyInfo['PostBuildingNumber']
        CompanyActionCommandDto.PostCity = companyInfo['PostCity']
        CompanyActionCommandDto.PostCountry = companyInfo['PostCountry']
        CompanyActionCommandDto.PostIndex = companyInfo['PostIndex']
        CompanyActionCommandDto.ShortDescription = companyInfo['ShortDescription']
        CompanyActionCommandDto.ShortName = companyInfo['ShortName']
        CompanyActionCommandDto.UrlSite = companyInfo['UrlSite']

        if companyInfo.get('GUID') != None:
            ArrayOfCompanyActionCommandDto.CompanyUpdateCommandDto.append(CompanyActionCommandDto)
        else:
            ArrayOfCompanyActionCommandDto.CompanyCreationCommandDto.append(CompanyActionCommandDto)

        try:
            if companyInfo.get('GUID') != None:
                info = self.client.service.Update(ArrayOfCompanyActionCommandDto)
                self.addNoteToLogFile('Обновлена компания. %s' % info)
            else:
                info = self.client.service.Create(ArrayOfCompanyActionCommandDto)
                guidRowDict = {6: info.CompanyDto[0].Id}
                self.addGuidDictToQueue("О КОМПАНИИ", guidRowDict)
                self.addNoteToLogFile('Создана компания. %s' % info)
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

    def createArrayWithCompanyInfo(self, arrayWithCompanyInfo):
        arrayOfDictWithCompanyInfo = []
        if arrayWithCompanyInfo[0][4] != "":
            arrayOfDictWithCompanyInfo.append({
                'AddressBuildingNumber': arrayWithCompanyInfo[9][2], 'AddressCity': arrayWithCompanyInfo[8][2],
                'AddressCountry': arrayWithCompanyInfo[7][2], 'AddressIndex': int(arrayWithCompanyInfo[6][2]),
                'Email': arrayWithCompanyInfo[4][2], 'FullName': arrayWithCompanyInfo[1][2],
                'Phone': arrayWithCompanyInfo[5][2], 'GUID': arrayWithCompanyInfo[0][4],
                'PostBuildingNumber': arrayWithCompanyInfo[13][2], 'PostCity': arrayWithCompanyInfo[12][2],
                'PostCountry': arrayWithCompanyInfo[11][2], 'PostIndex': int(arrayWithCompanyInfo[10][2]),
                'ShortDescription': arrayWithCompanyInfo[2][2], 'ShortName': arrayWithCompanyInfo[1][2],
                'UrlSite': arrayWithCompanyInfo[3][2]
            })
            self.companyWorkWithId = arrayWithCompanyInfo[0][4]
        else:
            arrayOfDictWithCompanyInfo.append({
                'AddressBuildingNumber': arrayWithCompanyInfo[9][2], 'AddressCity': arrayWithCompanyInfo[8][2],
                'AddressCountry': arrayWithCompanyInfo[7][2], 'AddressIndex': int(arrayWithCompanyInfo[6][2]),
                'Email': arrayWithCompanyInfo[4][2], 'FullName': arrayWithCompanyInfo[1][2],
                'Phone': arrayWithCompanyInfo[5][2],
                'PostBuildingNumber': arrayWithCompanyInfo[13][2], 'PostCity': arrayWithCompanyInfo[12][2],
                'PostCountry': arrayWithCompanyInfo[11][2], 'PostIndex': int(arrayWithCompanyInfo[10][2]),
                'ShortDescription': arrayWithCompanyInfo[2][2], 'ShortName': arrayWithCompanyInfo[1][2],
                'UrlSite': arrayWithCompanyInfo[3][2]
            })
            self.companyWorkWithName = arrayWithCompanyInfo[1][2]

        return arrayOfDictWithCompanyInfo

    def workWithCompanyExcelController(self):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='О КОМПАНИИ')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=5)
        return self.createArrayWithCompanyInfo(arrayWithInfo)

    def createCompanyFromExcelController(self, defaultCompanyShortName):
        '''
        Создание компаниии по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с компаниями, авторизацию и т.д.
        '''
        holdingId = self.getHoldingIdByCompanyShortName(defaultCompanyShortName)
        self.startWorkWithInterface(interfaceNumberInArray=0)
        self.authorization()
        arrayOfDictWithCompanyInfo = self.workWithCompanyExcelController()
        self.createCompany(arrayOfDictWithCompanyInfo[0], holdingId)

    # =========================================================
    # МЕТОДЫ ДЛЯ СОЗДАНИЯ КОЛЛЕГИАЛЬНОГО ОРГАНА

    def getHeadOfAndSecretary(self, CBAmount):
        '''
        Создание структуры данных с ролями всех пользователей во всех КО
        '''

        def createDictWithCBInfoWithStructure(self, arrayWithCBRolesWithNoStructure, CBAmount):
            '''
            Возвращает "словарь словарей" с информацией о ролях всех пользователей во всех КО.
            Конкретно для создания КО есть лишняя информация.
            '''
            arrayOfDictsWithCBUserRoles = []
            for i in range(CBAmount):
                d_ = {}
                for j in range(self.userToCreateAmount):
                    d_.update(
                        {arrayWithCBRolesWithNoStructure[j][i + 2]: arrayWithCBRolesWithNoStructure[j][1]})
                arrayOfDictsWithCBUserRoles.append(d_)
            return arrayOfDictsWithCBUserRoles

        def getUsefulFormatFromDictWithCBUserRoles(arrayOfDictWithCBUserRoles):
            '''
            "Конвертирование" "словаря словарей" в формат:
                - убрать участников (всевозможных типов)
                - вместо имен пользователей -- Id пользователя
            '''
            heads_and_secretaries = []
            for i in range(len(arrayOfDictWithCBUserRoles)):
                heads_and_secretaries.append({})
                for k in arrayOfDictWithCBUserRoles[i].keys():
                    if k == 'ПРЕД':
                        heads_and_secretaries[i].update({'ПРЕД': arrayOfDictWithCBUserRoles[i].get(k)})
                    if k == 'СЕК':
                        heads_and_secretaries[i].update({'СЕК': arrayOfDictWithCBUserRoles[i].get(k)})
            return heads_and_secretaries

        listWithCBInfo = self.readList('РОЛИ')
        arrayWithCBRolesWithNoStructure = self.readInfoFromList(listWithCBInfo, startInfoPosition=11, isCBRoles=True)
        arrayOfDictsWithCBUserRoles = createDictWithCBInfoWithStructure(self, arrayWithCBRolesWithNoStructure, CBAmount)
        arrayOfDictsWithCBUserRoles= getUsefulFormatFromDictWithCBUserRoles(arrayOfDictsWithCBUserRoles)
        return arrayOfDictsWithCBUserRoles

    def getCBIdByItsShortName(self, CBCompanyId, CBName):

        CBSearchCriteriaDto = self.client.factory.create('ns0:CollegialBodySearchCriteriaDto')
        CompanyIdentityDto = self.client.factory.create('ns0:IdentityDto')

        del CBSearchCriteriaDto.Attendance
        del CBSearchCriteriaDto.CollegialBodyType

        CBSearchCriteriaDto.ShortNameToken = CBName
        CompanyIdentityDto.Id = CBCompanyId
        CBSearchCriteriaDto.Company = CompanyIdentityDto
        try:
            cbInfo = self.client.service.Find(CBSearchCriteriaDto)
            if cbInfo == '':
                raise Exception('КО с таким именем нет.')
            return cbInfo.CollegialBodyDto[0].Id
        except WebFault as e:
            self.addNoteToLogFile(e.args, warning=True)

    def createSeveralCollegialBodies(self, arrayOfDictWithCBInfo, CBCompanyId, arrayOfDictsWithCBUserRoles):
        '''
        Создание Коллегиальных Органов, опираясь на информацию из входного значения - arrayOfcollegial
        После создания пользователей лучше, т.к. тут уже определяется секретарь
        Соответственно, желательно, чтобы пользователь, которому планируется присвоить роль секретаря был создан
        Создание все равно присходит "по одному" объекту КО, из-за атрибута "Родительский КО"
                                                             из-за необходимости валидации
        '''

        guidRowDict = {}
        rowCounter = 6

        for collegialBodyInfo in arrayOfDictWithCBInfo:
            if (collegialBodyInfo.get('GUID') != None and workStrategyFlag == 'i') or \
                    (collegialBodyInfo.get('GUID') == None and workStrategyFlag == 'u'):
                continue

            if collegialBodyInfo.get('GUID') != None:
                actionType = "Update"
            else:
                actionType = "Creation"

            ArrayOfCollegialBodyActionCommandDto = self.client.factory.create(
                'ns0:ArrayOfCollegialBody' + actionType + 'CommandDto')

            # Получим сначала ФИО председателя и секретаря данного КО
            CBHeadAndSecretaryInfo = arrayOfDictsWithCBUserRoles[collegialBodyInfo['Order'] - 1]

            secretaryOfCBLastName = CBHeadAndSecretaryInfo['СЕК']
            headOfCBLastName = CBHeadAndSecretaryInfo['ПРЕД']

            # Получим Id председателя и секретаря данного КО, основываясь на их ФИО
            headOfCBId = self.getUserIdByHisFI(headOfCBLastName)
            secretaryOfCBId = self.getUserIdByHisFI(secretaryOfCBLastName)

            # Нужно снова начать работу с нужным интерфейсом, т.к. для получения
            #       Id секретаря и председателя было переклчения на UserManagementService
            self.startWorkWithInterface(2)
            self.authorization()

            CollegialBodyActionCommandDto = self.client.factory.create('ns0:CollegialBody' + actionType + 'CommandDto')
            AttendanceTypeEnumDto = self.client.factory.create('ns0:AttendanceTypeEnumDto')
            CollegialBodyTypeEnumDto = self.client.factory.create('ns0:CollegialBodyTypeEnumDto')
            IdentityDtoCompany = self.client.factory.create('ns0:IdentityDto')
            LdapUserIdentityDtoHeadOf = self.client.factory.create('ns0:LdapUserIdentityDto')
            IdentityDtoParent = self.client.factory.create('ns0:IdentityDto')
            LdapUserIdentityDtoSecretaryOf = self.client.factory.create('ns0:LdapUserIdentityDto')

            if collegialBodyInfo.get('GUID') != None:
                CollegialBodyActionCommandDto.Id = collegialBodyInfo.get('GUID')
            else:
                IdentityDtoCompany.Id = CBCompanyId
                CollegialBodyActionCommandDto.Company = IdentityDtoCompany
            del CollegialBodyActionCommandDto.Id
            if collegialBodyInfo['ParentCBName'] != '':
                IdentityDtoParent.Id = self.getCBIdByItsShortName(CBCompanyId, collegialBodyInfo['ParentCBName'])
                CollegialBodyActionCommandDto.Parent = IdentityDtoParent
            LdapUserIdentityDtoHeadOf.Id = headOfCBId
            del LdapUserIdentityDtoHeadOf.LdapUsername
            CollegialBodyActionCommandDto.HeadOf = LdapUserIdentityDtoHeadOf
            if collegialBodyInfo['CBType'] == 'ИСПОЛНИТЕЛЬНЫЙ':
                CollegialBodyActionCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.Executive)
            elif collegialBodyInfo['CBType'] == 'НЕ ИСПОЛНИТЕЛЬНЫЙ':
                CollegialBodyActionCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.NotExecutive)
            elif collegialBodyInfo['CBType'] == 'НЕ КОРПОРАТИВНЫЙ':
                CollegialBodyActionCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.NotCorporate)
            elif collegialBodyInfo['CBType'] == 'ГОСУДАРСТВЕННЫЙ':
                CollegialBodyActionCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.State)
            elif collegialBodyInfo['CBType'] == 'ОРГАН УПРАВЛЕНИЯ':
                CollegialBodyActionCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.ManagementBody)
            else:
                pass
                # можно добавить выброс исключения

            if collegialBodyInfo['AttendanceType'] == "ЗАОЧНОЕ":
                CollegialBodyActionCommandDto.Attendance.set(AttendanceTypeEnumDto.__keylist__[0])
            else:
                CollegialBodyActionCommandDto.Attendance.set(AttendanceTypeEnumDto.__keylist__[1])

            CollegialBodyActionCommandDto.FullName = collegialBodyInfo['FullName']
            CollegialBodyActionCommandDto.Order = collegialBodyInfo['Order']
            CollegialBodyActionCommandDto.QualifiedMajority = collegialBodyInfo['QualifiedMajority']
            LdapUserIdentityDtoSecretaryOf.Id = secretaryOfCBId
            del LdapUserIdentityDtoSecretaryOf.LdapUsername
            CollegialBodyActionCommandDto.Secretary = LdapUserIdentityDtoSecretaryOf
            CollegialBodyActionCommandDto.ShortDescription = collegialBodyInfo['ShortDescription']
            CollegialBodyActionCommandDto.ShortName = collegialBodyInfo['ShortName']

            if collegialBodyInfo.get('GUID') != None:
                ArrayOfCollegialBodyActionCommandDto.CollegialBodyUpdateCommandDto.append(
                    CollegialBodyActionCommandDto)
            else:
                ArrayOfCollegialBodyActionCommandDto.CollegialBodyCreationCommandDto.append(
                    CollegialBodyActionCommandDto)

            try:
                if collegialBodyInfo.get('GUID') != None:
                    info = self.client.service.Update(ArrayOfCollegialBodyActionCommandDto)
                    self.addNoteToLogFile('Обновлен колегиальный орган. %s' % info)
                else:
                    info = self.client.service.Create(ArrayOfCollegialBodyActionCommandDto)
                    self.addNoteToLogFile('Создан колегиальный орган. %s' % info)
                    guidRowDict.update({rowCounter: info.CollegialBodyDto[0].Id})
                rowCounter += 1
            except WebFault as e:
                self.addNoteToLogFile(e.args, warning=True)

        self.addGuidDictToQueue("КО", guidRowDict)

    def createArrayOfDictWithCBInfo(self, arrayWithCBInfo, qualifiedCBUsersCount=4):
        arrayOfDictWithCBInfo = []
        j = 1
        for x in arrayWithCBInfo:
            if len(x) == 16:
                arrayOfDictWithCBInfo.append({
                    'FullName': x[2], 'ShortName': x[4], 'GUID': x[-1],
                    'Order': j, 'QualifiedMajority': qualifiedCBUsersCount,
                    'ShortDescription': x[6], 'ParentCBName': x[14],
                    'AttendanceType': x[12], 'CBType': x[13]
                })
            else:
                arrayOfDictWithCBInfo.append({
                    'FullName': x[2], 'ShortName': x[4],
                    'Order': j, 'QualifiedMajority': qualifiedCBUsersCount,
                    'ShortDescription': x[6], 'ParentCBName': x[14],
                    'AttendanceType': x[12], 'CBType': x[13]
                })
            j += 1
        return arrayOfDictWithCBInfo

    def workWithCBExcelController(self):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='КО')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=5)
        return self.createArrayOfDictWithCBInfo(arrayWithInfo)

    def createCBFromExcelController(self):
        '''
        Создание КО по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с КО, авторизацию и т.д.
        '''
        if self.companyWorkWithId == None:
            self.companyWorkWithId = self.getCompanyIdByItsShortName(self.companyWorkWithName)
        CBCompanyId = self.companyWorkWithId
        self.startWorkWithInterface(interfaceNumberInArray=2)
        self.authorization()
        arrayOfDictWithCBInfo = self.workWithCBExcelController()
        arrayOfDictsWithCBUserRoles = self.getHeadOfAndSecretary(len(arrayOfDictWithCBInfo))
        self.createSeveralCollegialBodies(arrayOfDictWithCBInfo, CBCompanyId, arrayOfDictsWithCBUserRoles)


if __name__ == '__main__':
    try:
        sys.argv = sys.argv[1:]
        if sys.argv[0] != "help":
            url = sys.argv[0]
            clientBM = ClientBM(url)
            global excelFilePathPlusName
            excelFilePathPlusName = sys.argv[1]
            login = sys.argv[2]
            password = sys.argv[3]
            clientBM.setLoginAndPassword(login, password)
            defaultCompanyShortName = sys.argv[4]
            defaultPassword = sys.argv[5]
            if sys.argv[6] not in ["ui", "i", "u"]:
                raise Exception("Unknown last flag name. Please, choose flag from {ui, i, u}.")
            global workStrategyFlag
            workStrategyFlag = sys.argv[6]
            debugMode = sys.argv[7] != None

            clientBM.createCompanyFromExcelController(defaultCompanyShortName)
            clientBM.createUsersFromExcelController(defaultPassword)
            clientBM.createCBFromExcelController()

            clientBM.writeGuidToExcel()
        else:
            print("""\nПАРАМЕТРЫ СКРИПТА (ПОРЯДОК ВАЖЕН):
                    - адрес сервера BM
                    - путь к excel файлу (с расширением и именем самого файла)
                    - логин для входа на сервер BM
                    - пароль для входа на сервер BM
                    - короткое имя любой существующей компании стенда для поиска Id Холдинга
                    - желаемый пароль для создаваемых пользователей
                    - 'i'/'iu'/'u' -- создание / создание и обновление / обновление
                    - 'debug' (по желанию) для вывода стектрейса в терминал""")

    except Exception as e:
        if debugMode:
            raise e
        clientBM.addNoteToLogFile(e.args, warning=True)
