import suds
from suds import client
from suds.sax.element import Element
from suds.wsse import *
import xlrd

class ClientBM:
    
    serverURL = str()
    client = None
    currentInterfacenumberInArray = int()
    excelFile = None
    defaultEmail = 'demoboardmaps@yandex.ru'
    ArrayOfDictWithUsersInfo = []
    interfaces = ['CompanyManagementService','UserManagementService','CollegialBodyManagementService',
                 'MeetingManagementService','MeetingMemberManagementService','IssueManagementService'
                 ,'DecisionProjectManagementService','InvitedMember2IssueManagementService',
                  'SpokesPerson2IssueManagementService',
                  'MaterialManagementService','DocumentManagementService','InstructionManagementService']
    
    def __init__(self, serverURL):
        self.serverURL = serverURL
        
    def startWorkWithInterface(self, interfaceNumberInArray):
        '''
        Отсчет interfaceNumberInArray начинается с 0
        '''
        self.currentInterfacenumberInArray = interfaceNumberInArray
        self.client = suds.client.Client(self.serverURL + "/PublicApi/" + self.interfaces[interfaceNumberInArray] + ".svc?wsdl")

    def authorization(self, login, password):
        '''
        После запуска метода startWorkWithInterface
        '''
        security = Security()
        token = UsernameToken(login, password)
        security.tokens.append(token)
        self.client.set_options(wsse=security)
        
    def createUser(self, userInfo, usersCompanyId, useDefaultEmail):
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
        if useDefaultEmail:
            UserCreationCommandDto.Email = self.defaultEmail
        else:
            UserCreationCommandDto.Email = userInfo['Email']
        UserCreationCommandDto.FirstName = userInfo['FirstName']
        UserCreationCommandDto.LastName = userInfo['LastName']
        UserCreationCommandDto.Phone = userInfo['Phone']
        UserCreationCommandDto.Position = userInfo['Position']
        
        ArrayOfUserCreationCommandDto.UserCreationCommandDto.append(UserCreationCommandDto)
        
        try:
            getInfo = self.client.service.Create(ArrayOfUserCreationCommandDto)
            print(getInfo)
        except WebFault as e:
            print(e)
            
    def createSeveralUsers(self, arrayOfUserInfoDict, usersCompanyId, useDefaultEmail):
        '''
        arrayOfUserInfoDict - массив словарей с информацией о пользователях для их создания
        '''
        for userInfo in arrayOfUserInfoDict:
            self.createUser(userInfo, usersCompanyId, useDefaultEmail)
            
    def getCompanyShortName(self):
        return str(input('Введите короткое имя компании, в которой требуется создать пользователей: '))
    
    def getDecisionAboutDefaultEmail(self):
        return bool(input('Использовать стандартный email для всех пользователей - ' + self.defaultEmail + 
                                     '?.\n Да - 1, нет - 0. : '))

    def createUsersFromExcelController(self, excelFilePathPlusName, login, password, defaultPassword):
        '''
        Создание пользователей по информации из excel.
        Включает начало работы с интерфейсом по работе (бог тафтологии) с пользователями, авторизацию и т.д.
        '''
        companyShortName = self.getCompanyShortName()
        print('---------------------')
        usersCompanyId = self.getCompanyIdByItsShortName(companyShortName, login, password)
        self.startWorkWithInterface(interfaceNumberInArray=1)
        self.authorization(login, password)
        useDefaultEmail = self.getDecisionAboutDefaultEmail()
        print('---------------------')
        self.workWithExcelController(excelFilePathPlusName, defaultPassword)
        self.createSeveralUsers(self.ArrayOfDictWithUsersInfo, usersCompanyId, useDefaultEmail)
            
    def openExcelFile(self, filePathPlusName):
        self.excelFile = xlrd.open_workbook(filePathPlusName)
        
    def readUsersList(self, usersListName='ПОЛЬЗОВАТЕЛИ'):
        '''
        Название листа с пользователями:
        рус - 'ПОЛЬЗОВАТЕЛИ'
        англ - 'USERS'
        '''
        return self.excelFile.sheet_by_name(usersListName)
    
    def readUsersInfoFromList(self, listFile, startInfoPosition=4):
        '''
        Создается массив со ВСЕЙ информацией о пользователях из таблицы
        Немного нерационально, т.к. можно сразу тут создать словарь с нужными ключами
        Но не хочется так просто выкидывать часть информации -- вдруг пригодится потом
        '''
        arrayWithUsersInfo = []
        i = startInfoPosition
        try:
            while listFile.row_values(i)[2] != '':
                arrayWithUsersInfo.append(listFile.row_values(i))
                i += 1
        except IndexError as e:
            pass
        return arrayWithUsersInfo
    
    def createArrayOfDictWithUsersInfo(self, arrayWithUsersInfo, defaultPassword):
        for x in arrayWithUsersInfo:
            self.ArrayOfDictWithUsersInfo.append({
                    'BasicPassword': defaultPassword, 'BasicUsername': x[14], 'Blocked': False,
                    'Company': None, 'Email': x[11], 'FirstName': x[2].split()[1], 
                    'LastName': x[2].split()[0], 'Phone': x[12], 'Position': x[9]
                 })
            
    def workWithExcelController(self, excelFilePathPlusName, defaultPassword):
        self.openExcelFile(excelFilePathPlusName)
        excelList = self.readUsersList()
        arrayWithUsersInfo = self.readUsersInfoFromList(excelList)
        self.createArrayOfDictWithUsersInfo(arrayWithUsersInfo, defaultPassword)
        
    def getCompanyIdByItsShortName(self, companyShortName, login, password):
        self.startWorkWithInterface(0)
        self.authorization(login, password)
        CompanySearchCriteriaDto = self.client.factory.create('ns0:CompanySearchCriteriaDto')
        CompanySearchCriteriaDto.ShortNameToken = companyShortName
        try:
            companyInfo = self.client.service.Find(CompanySearchCriteriaDto)
            if companyInfo == '':
                raise Exception('Компании с таким именем нет.')
            return companyInfo.CompanyDto[0].Id
        except WebFault as e:
            print(e)
			
if __name__ == '__main__':
    print('Старт работы скрипта.')
    print('---------------------')
    url = str(input('Введите URL сервера boardmaps (with http or https): '))
    print('---------------------')
    clientBM = ClientBM(url)
    excelFilePathPlusName = str(input('''Введите путь к excel файлу и его имя 
                                    (лучше положить excel файл в ту же папку,
                                    что и скрипт). Не забудьте про расширение файла \
                                    - укажите его после имени файла через точку.: '''))
    print('---------------------')
    login = str(input('Введите логин для входа на сервер boardmaps: '))
    print('---------------------')
    password = str(input('Введите пароль для входа на сервер boardmaps: '))
    print('---------------------')
    defaultPassword = str(input('Введите желаемый пароль для всех пользователей, которых вы хотите создать: '))
    print('---------------------')
    clientBM.createUsersFromExcelController(excelFilePathPlusName, login, password, defaultPassword)
    print('---------------------')
    print('Конец работы скрипта.')
    #raw_input()
