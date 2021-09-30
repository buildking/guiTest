import configparser


class ConfigUtil:
    def __init__(self):
        try:
            self.config = configparser.ConfigParser()
            self.readConfig()
        except:
            self.config = None

    def readConfig(self):
        self.config.read('./config.ini', encoding='utf-8')
        self.config.optionxform = str

    def getConfig(self):
        return self.config

    def getExcelSetting(self):
        if self.config is None or len(self.config.sections()) == 0:
            raise Exception("config 파일을 읽지 못했습니다.")
        else :
            excelSection = self.config['excel']
            return dict(excelSection)

    def getOracleSetting(self):
        if self.config is None or len(self.config.sections()) == 0:
            raise Exception("config 파일을 읽지 못했습니다.")
        else :
            oracleSection = self.config['oracle']
            return dict(oracleSection)

#단위 테스트
testConfig = ConfigUtil()
try:
    print(testConfig.getOracleSetting())
except Exception as e:
    print(e)

