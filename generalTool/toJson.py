import json
import os,sys,xlrd

class workRound:
    def __init__(self, **kwargs):
        os.chdir(os.path.abspath(os.path.dirname(__file__)))
        self.__dict__.update(kwargs)

    def set_spec_dir(self):
        for i in self.__dict__:
            self.__dict__[i] = os.path.split(os.getcwd())[0] + os.sep + self.__dict__[i]

    def __str__(self):
        mesg = ""
        for i in self.__dict__:
           mesg += f"{i}:\t{self.__dict__[i]}\n"
        return mesg

    @staticmethod
    def get_data_from_excel(e_path, sheet_num):
        new_title = ["ns", "key", "ns_p", "sid", "l_name", "length"]  # rename title which given from excels
        data = xlrd.open_workbook(e_path)
        table = data.sheets()[sheet_num]
        nor = table.nrows
        nol = table.ncols
        dict = {}
        for i in range(2, nor):
            for j in range(nol):
                # title = table.cell_value(0, j)
                title = new_title[j]
                value = table.cell_value(i, j)
                dict[title] = value
            yield dict


class Ns_Key(workRound):
    jsonDict = {}
    jsonList = []
    def __init__(self, ns, key, ns_p, sid = "n/a", l_name= "n/a", length= "n/a"):
        """

        :param ns: Namespace_hex
        :param key: Key
        :param ns_p: Namespace_purpose
        :param sid: Specific ID
        :param lName: Long name
        :param length: Length
        """
        self.ns = ns
        self.key = key
        self.ns_p = ns_p
        self.sid = sid
        self.l_name = l_name
        self.length = length
        self.jsonList.append(self.__dict__)

    def setDict(self):
        return self.__dict__

    def merge_json_list(self,extra_json_list):
        for i in extra_json_list:
            self.jsonList.append(i)

    def add_root_name(self, root_name):
        self.jsonDict[root_name] = self.jsonList

    def __repr__(self):
        return f"ns:\t{self.ns}\n" \
               f"key:\t{self.key}\n" \
               f"ns_p:\t{self.ns_p}\n" \
               f"sid:\t{self.sid}\n" \
               f"l_name:\t{self.l_name}\n" \
               f"length:\t{self.length}\n"





if __name__ == '__main__':
    env = workRound(e_dir = 'excels', j_dir = 'jsons', g_dir = "generalTool")
    env.set_spec_dir()
    e_name = "MIB3_DIAG_Key_Value_Pairs_CNS3.0.xls"
    env.e_path = env.e_dir + os.sep + e_name
    env.nsKeyList = []
    for i in env.get_data_from_excel(env.e_path,3):
        env.nsKeyList.append(i.copy())
    ns_key= Ns_Key("0x3000000","0xF18C","Identification")
    ns_key.merge_json_list(env.nsKeyList)
    ns_key.add_root_name("ns_key")
    j_name = "ns_key.json"
    j_path = env.j_dir + os.sep + j_name
    def saveAsFile(j_path,json_data):
        json.dump(json_data, open(j_path, 'w'), ensure_ascii=False,
                  indent=4, separators=(", ", " : "))
    saveAsFile(j_path,ns_key.jsonDict)

