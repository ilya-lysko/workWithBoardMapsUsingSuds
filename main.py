import logging
from urllib.error import URLError

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
    password = None
    userToCreateAmount = None
    companyWorkWithShortName = None
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
            self.addNoteToLogFile('\n\nНачало работы с сервером %s' % self.serverURL)
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

    # =========================================================
    # МЕТОДЫ ДЛЯ СОЗДАНИЯ ПОЛЬЗОВАТЕЛЕЙ

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
        UserCreationCommandDto.Blocked = False
        Company.Id = usersCompanyId
        UserCreationCommandDto.Company = Company
        if userInfo['Email'] == '':
            UserCreationCommandDto.Email = self.defaultEmail
        else:
            UserCreationCommandDto.Email = userInfo['Email']
        UserCreationCommandDto.FirstName = userInfo['FirstName']
        UserCreationCommandDto.LastName = userInfo['LastName']
        UserCreationCommandDto.MiddleName = userInfo['MiddleName']
        UserCreationCommandDto.Phone = userInfo['Phone']
        UserCreationCommandDto.Position = userInfo['Position']
        UserCreationCommandDto.Bio = userInfo['Bio']
        # UserCreationCommandDto.PhotoImageSource = "materials/default_picture.png" # не работает :(

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
                'BasicPassword': defaultPassword, 'BasicUsername': x[12], 'Bio': x[15],
                'Email': x[8], 'FirstName': x[3],
                'LastName': x[2], 'MiddleName': x[4], 'Phone': x[9], 'Position': x[10]
            })
        return arrayOfDictWithUsersInfo

    def workWithUsersExcelController(self, excelFilePathPlusName, defaultPassword):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='ПОЛЬЗОВАТЕЛИ')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=6)
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

    # =========================================================
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
            'AddressBuildingNumber': arrayWithCompanyInfo[9][2], 'AddressCity': arrayWithCompanyInfo[8][2],
            'AddressCountry': arrayWithCompanyInfo[7][2], 'AddressIndex': int(arrayWithCompanyInfo[6][2]),
            'Email': arrayWithCompanyInfo[4][2], 'FullName': arrayWithCompanyInfo[1][2],
            'Phone': arrayWithCompanyInfo[5][2],
            'PostBuildingNumber': arrayWithCompanyInfo[13][2], 'PostCity': arrayWithCompanyInfo[12][2],
            'PostCountry': arrayWithCompanyInfo[11][2], 'PostIndex': int(arrayWithCompanyInfo[10][2]),
            'ShortDescription': arrayWithCompanyInfo[2][2], 'ShortName': arrayWithCompanyInfo[1][2],
            'UrlSite': arrayWithCompanyInfo[3][2]
        })
        self.companyWorkWithShortName = arrayWithCompanyInfo[1][2]
        return arrayOfDictWithCompanyInfo

    def workWithCompanyExcelController(self, excelFilePathPlusName):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='О КОМПАНИИ')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=5)
        return self.createArrayWithCompanyInfo(arrayWithInfo)

    def createCompanyFromExcelController(self, excelFilePathPlusName, defaultCompanyShortName):
        '''
        Создание компаниии по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с компаниями, авторизацию и т.д.
        '''
        companyShortName = defaultCompanyShortName
        holdingId = self.getHoldingIdByCompanyShortName(companyShortName)
        self.startWorkWithInterface(interfaceNumberInArray=0)
        self.authorization()
        arrayOfDictWithCompanyInfo = self.workWithCompanyExcelController(excelFilePathPlusName)
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
        '''

        for collegialBodyInfo in arrayOfDictWithCBInfo:

            ArrayOfCollegialBodyCreationCommandDto = self.client.factory.create(
                'ns0:ArrayOfCollegialBodyCreationCommandDto')

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

            CollegialBodyCreationCommandDto = self.client.factory.create('ns0:CollegialBodyCreationCommandDto')
            AttendanceTypeEnumDto = self.client.factory.create('ns0:AttendanceTypeEnumDto')
            CollegialBodyTypeEnumDto = self.client.factory.create('ns0:CollegialBodyTypeEnumDto')
            IdentityDtoCompany = self.client.factory.create('ns0:IdentityDto')
            LdapUserIdentityDtoHeadOf = self.client.factory.create('ns0:LdapUserIdentityDto')
            IdentityDtoParent = self.client.factory.create('ns0:IdentityDto')
            LdapUserIdentityDtoSecretaryOf = self.client.factory.create('ns0:LdapUserIdentityDto')

            del CollegialBodyCreationCommandDto.Id
            if collegialBodyInfo['ParentCBName'] != '':
                IdentityDtoParent.Id = self.getCBIdByItsShortName(CBCompanyId, collegialBodyInfo['ParentCBName'])
                CollegialBodyCreationCommandDto.Parent = IdentityDtoParent
            LdapUserIdentityDtoHeadOf.Id = headOfCBId
            del LdapUserIdentityDtoHeadOf.LdapUsername
            CollegialBodyCreationCommandDto.HeadOf = LdapUserIdentityDtoHeadOf
            IdentityDtoCompany.Id = CBCompanyId
            CollegialBodyCreationCommandDto.Company = IdentityDtoCompany
            if collegialBodyInfo['CBType'] == 'ИСПОЛНИТЕЛЬНЫЙ':
                CollegialBodyCreationCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.Executive)
            elif collegialBodyInfo['CBType'] == 'НЕ ИСПОЛНИТЕЛЬНЫЙ':
                CollegialBodyCreationCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.NotExecutive)
            elif collegialBodyInfo['CBType'] == 'НЕ КОРПОРАТИВНЫЙ':
                CollegialBodyCreationCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.NotCorporate)
            elif collegialBodyInfo['CBType'] == 'ГОСУДАРСТВЕННЫЙ':
                CollegialBodyCreationCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.State)
            elif collegialBodyInfo['CBType'] == 'ОРГАН УПРАВЛЕНИЯ':
                CollegialBodyCreationCommandDto.CollegialBodyType.set(CollegialBodyTypeEnumDto.ManagementBody)
            else:
                pass
                # можно добавить выброс исключения

            if collegialBodyInfo['AttendanceType'] == "ЗАОЧНОЕ":
                CollegialBodyCreationCommandDto.Attendance.set(AttendanceTypeEnumDto.__keylist__[0])
            else:
                CollegialBodyCreationCommandDto.Attendance.set(AttendanceTypeEnumDto.__keylist__[1])

            CollegialBodyCreationCommandDto.FullName = collegialBodyInfo['FullName']
            CollegialBodyCreationCommandDto.Order = collegialBodyInfo['Order']
            CollegialBodyCreationCommandDto.QualifiedMajority = collegialBodyInfo['QualifiedMajority']
            LdapUserIdentityDtoSecretaryOf.Id = secretaryOfCBId
            del LdapUserIdentityDtoSecretaryOf.LdapUsername
            CollegialBodyCreationCommandDto.Secretary = LdapUserIdentityDtoSecretaryOf
            CollegialBodyCreationCommandDto.ShortDescription = collegialBodyInfo['ShortDescription']
            CollegialBodyCreationCommandDto.ShortName = collegialBodyInfo['ShortName']

            ArrayOfCollegialBodyCreationCommandDto.CollegialBodyCreationCommandDto.append(
                CollegialBodyCreationCommandDto)

            try:
                getInfo = self.client.service.Create(ArrayOfCollegialBodyCreationCommandDto)
                self.addNoteToLogFile('Создан колегиальный орган. %s' % getInfo)
            except WebFault as e:
                self.addNoteToLogFile(e.args, warning=True)

    def createArrayOfDictWithCBInfo(self, arrayWithCBInfo, qualifiedCBUsersCount=4):
        arrayOfDictWithCBInfo = []
        j = 1
        for x in arrayWithCBInfo:
            arrayOfDictWithCBInfo.append({
                'FullName': x[2], 'ShortName': x[4],
                'Order': j, 'QualifiedMajority': qualifiedCBUsersCount,
                'ShortDescription': x[7], 'ParentCBName': x[14],
                'AttendanceType': x[12], 'CBType': x[13]
            })
            j += 1
        return arrayOfDictWithCBInfo

    def workWithCBExcelController(self, excelFilePathPlusName):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readList(listName='КО')
        arrayWithInfo = self.readInfoFromList(excelList, startInfoPosition=5)
        return self.createArrayOfDictWithCBInfo(arrayWithInfo)

    def createCBFromExcelController(self, excelFilePathPlusName):
        '''
        Создание КО по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с КО, авторизацию и т.д.
        '''
        CBCompanyId = self.getCompanyIdByItsShortName()
        self.startWorkWithInterface(interfaceNumberInArray=2)
        self.authorization()
        arrayOfDictWithCBInfo = self.workWithCBExcelController(excelFilePathPlusName)
        arrayOfDictsWithCBUserRoles = self.getHeadOfAndSecretary(len(arrayOfDictWithCBInfo))
        self.createSeveralCollegialBodies(arrayOfDictWithCBInfo, CBCompanyId, arrayOfDictsWithCBUserRoles)


if __name__ == '__main__':
    try:
        sys.argv = sys.argv[1:]
        if sys.argv[0] != "help":
            url = sys.argv[0]
            clientBM = ClientBM(url)
            excelFilePathPlusName = sys.argv[1]
            login = sys.argv[2]
            password = sys.argv[3]
            clientBM.setLoginAndPassword(login, password)
            defaultCompanyShortName = sys.argv[4]
            defaultPassword = sys.argv[5]

            clientBM.createCompanyFromExcelController(excelFilePathPlusName, defaultCompanyShortName)
            clientBM.createUsersFromExcelController(excelFilePathPlusName, defaultPassword)
            clientBM.createCBFromExcelController(excelFilePathPlusName)
        else:
            print("""\nПАРАМЕТРЫ СКРИПТА (ПОРЯДОК ВАЖЕН):\n
                    - адрес сервера BM\n
                    - путь к excel файлу (с расширением и именем самого файла)\n
                    - логин для входа на сервер BM\n
                    - пароль для входа на сервер BM\n
                    - короткое имя любой существующей компании стенда для поиска Id Холдинга\n
                    - желаемый пароль для создаваемых пользователей\n""")

    except Exception as e:
        clientBM.addNoteToLogFile(e.args, warning=True)
