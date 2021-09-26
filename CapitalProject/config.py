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

    def getConfig(self):
        return self.config

    def getOracleSetting(self):
        if self.config is None or len(self.config.sections()) == 0:
            raise Exception("config 파일을 읽지 못했습니다.")
        else :
            oracleSection = self.config['oracle']
            oracleSetting = {
                "ip": oracleSection['ip'],
                "port": oracleSection['port'],
                "service": oracleSection['service'],
                "userName": oracleSection['userName'],
                "userPassword": oracleSection['userPassword']
            }
            return oracleSetting

#단위 테스트
testConfig = ConfigUtil()
try:
    print(testConfig.getOracleSetting())
except Exception as e:
    print(e)

