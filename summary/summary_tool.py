if not __package__:
    import sys
    import os
    path = os.path.join(os.path.dirname(__file__), os.pardir)
    sys.path.insert(0, path)

import os, time, random, sys, enum
import ujson as json
import glob
import pandas as pd
import tkinter as tk
from tkinter import filedialog as fd
from tkinter import *
from tkinter import messagebox
import distutils
from distutils import util
import timeit
from pprint import pprint
from collections import OrderedDict
from datetime import datetime, timedelta
import traceback
import utils
import yaml
import threading
import re
from attachment.attachment_tool import RUN_ATTACHMENT
from win32com.client import Dispatch

class Summary_WorkStatus(enum.Enum):
    PARSING = "PARSING"
    GENERATE = "GENERATING"
    COLORING = "COLORING"
    FINISH = "FINISH"

class RUN_SUMMARY(threading.Thread):
    
    def __init__(self, log_path = '', logfilter = False):
        threading.Thread.__init__(self)
        self.stop_event = threading.Event()
        self.filter_switch = logfilter
        print("===== LOG FILTER IS ", str(self.filter_switch).upper(), " =====")
        time.sleep(1.5)
        self.input_logpath = log_path
        self.load_config()
        self.init_para()

    def load_config(self):
        self.config_file = "./config.yaml"
        yamlfile = utils.read_yaml(self.config_file)
        self.config_data = yaml.load(yamlfile, Loader=yaml.FullLoader)

    def init_para(self):
        self.finish = False
        # self.file_list = {}	
        # self.count = 0
        # self.weight = [0.33, 0.33, 0.33]
        self.boolean = ['TRUE','True','FALSE','False']   
        # self.header_list = ["SN", "Test Mode", "Test Version", "BFT", "Config", "Start time", "End time", "slot", "Outcome", 
        #     "Failure Code", "Failure Items"]
        #----------------------------------------------------
        if not self.input_logpath:
            self.logs_path = utils.read_config(self.config_data, "Output", "LocalOutPutPath")
        else:
            self.logs_path = self.input_logpath

        
        #------------------------------------------------------
        self.datebegin = utils.read_config(self.config_data, "FTPFilter", "DateBegin")
        self.dateend = utils.read_config(self.config_data, "FTPFilter", "DateEnd")
        self.timebegin = utils.read_config(self.config_data, "FTPFilter", "TimeBegin")
        self.timeend = utils.read_config(self.config_data, "FTPFilter", "TimeEnd")
        self.station = utils.read_config(self.config_data, "FTPFilter", "Station")
        self.serialnumber = utils.read_config(self.config_data, "FTPFilter", "SN")
        self.testmode = utils.read_config(self.config_data, "FTPFilter", "Test_mode")
        self.outcome = utils.read_config(self.config_data, "FTPFilter", "Outcome")
        self.slot = utils.read_config(self.config_data, "FTPFilter", "Slot")
        self.attachment_switch = utils.read_config(self.config_data, "FTPFilter", "Attachments")

    def set_output_summary_filename(self, sheetnamekeys):
        testname = ""
        for key in sheetnamekeys:
            testname += str(key)
            testname += "_"

        if self.filter_switch:
            self.summary_outputname = 'summary_output_filter_' + testname + datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + '.xlsx'
        else:
            self.summary_outputname = 'summary_output_filternone_' + datetime.now().strftime("%Y-%m-%d-%H-%M-%S") + '.xlsx'

    def check_folder(self):
        self.ALLfiles = list()
        self.ALLfileNAMEs = list()
        start = timeit.default_timer()
        
        files = glob.glob(self.logs_path + "/**/*.json", recursive=True)
        self.filterout_file = list()
        for file in files:
            filename = os.path.basename(file)

            if self.filter_switch:
                datetimepos, slotpos = utils.check_pos(filename)
                date, datebool = utils.date_filter(file = file, datetimepos = datetimepos, datebegin = self.datebegin, dateend = self.dateend)
                filetime, timebool = utils.time_filter(file = file, datetimepos = datetimepos, timebegin = self.timebegin, timeend = self.timeend, datebegin = self.datebegin, dateend = self.dateend)
                st, stbool = utils.station_filter(file = file, station = self.station)
                sn, snbool = utils.Any_filter(file = file, filtertext = self.serialnumber)
                tm, tmbool = utils.Any_filter(file = file, filtertext = self.testmode)
                sl, slbool = utils.Slot_filter(file = file, slot_pos = slotpos, slot = self.slot)
                oc, ocbool = utils.Outcome_filter(file = file, outcome = self.outcome)
                #print(datebool, timebool, stbool, snbool, tmbool, slbool, ocbool)

                if not datebool or not timebool or not stbool or not snbool or not tmbool or not slbool or not ocbool:
                    self.filterout_file.append((date,file))
                    # 
                    # print(date, file)
                if datebool & timebool & stbool & snbool & tmbool & slbool & ocbool:
                    match_file_txt = "Match Station: {}, SN: {}, Test Mode: {}, Date: {}, Time: {}, Outcome: {}, Slot: {}, \nFILE = {}\n".format(st, sn, tm, date, filetime, oc, sl, filename)
                    # print(match_file_txt)
                    if filename in self.ALLfileNAMEs:
                        print("SAME FILE: ", file)
                        continue
                    self.ALLfileNAMEs.append(filename)
                    self.ALLfiles.append(file)
                    
            else:
                if filename in self.ALLfileNAMEs:
                    print("SAME FILE: ", file)
                    continue
                # print("File is found: ", file)
                self.ALLfileNAMEs.append(filename)
                self.ALLfiles.append(file)
            
        print("TOTAL FILE COUNTS: ", len(self.ALLfiles))

        if self.filterout_file:
            print("***THESE FILE MIGHT NOT IN CORRECT DATE RANGE***")
            for item in self.filterout_file:
                print(item)


        #self.ALLfiles = files

        stop = timeit.default_timer()
        self.previous_time = stop - start


    def __finish__(self):
        self.work_status = Summary_WorkStatus.FINISH.value
        self.stop()

    def stop(self):
        self.stop_event.set()

    def stopped(self):
        return self.stop_event.is_set()       

    def run(self):
        try:
            self.check_folder()
            self.sheet_name = {}
            self.sheet_color = {}
            start = timeit.default_timer()  	
            # for (foldername, self.files) in self.file_list.items():
            #     #print(self.file)	
            #     print("\nReading files under " + foldername + "'s folder...")
            #     self.read_jsonfile_data()
            
            self.read_jsonfile_data()
            
                
            self.write_summary()
            stop = timeit.default_timer()
            print("Time: {:.2f}".format(self.previous_time + (stop - start)))
            utils.write_config(self.config_file, "Output", "SummaryFileName", self.summary_outputname)

            if self.filter_switch and self.attachment_switch:
                self.do_attachment = RUN_ATTACHMENT(self.ALLfiles)
                self.do_attachment.start()
                self.do_attachment.join()

            self.finish = True
            self.__finish__()

        except:
            traceback.print_exc()

    def check_key_ifexist(self, index, key, keyname):
        try:
            if self.jsondata["phases"][index]["measurements"][key][keyname]:
                #print('True')
                return True
        except:
            #print('False')
            return False

    def bool_float_check(self, checkitem):
        if checkitem in self.boolean:
            checkitem = bool(distutils.util.strtobool(checkitem))   #轉boolean型態
        else:
            try:
                checkitem = float(checkitem)   #轉float 型態
                #self.dtype[key] = 'float64'       
            except:			
                pass   #String型態
        #print(checkitem, type(checkitem))
        return checkitem

    def read_jsonfile_data(self):	
 
        cnt_3_times_fail = 0	
        cnt_3_times_fail_sn = dict()
        self.work_status = Summary_WorkStatus.PARSING.value	
        for file in self.ALLfiles:
            if self.stopped():
                sys.exit()
            
            #print(file)
            with open(file , 'r') as reader:
                try:
                    self.jsondata = json.loads(reader.read())
                except Exception as name:
                    print(name)
                    continue
                
                list_of_key = OrderedDict()
                color_dict = OrderedDict()
                
                list_of_key["SN"] = self.jsondata["dut_id"]
                list_of_key["Test Mode"] = self.jsondata["test_mode"]
                list_of_key["Test Version"] = self.jsondata["metadata"]["test_version"]	
                list_of_key["BFT"] = self.jsondata["station_id"]

                try:
                    list_of_key["Config"] = self.jsondata["metadata"]["device_config"]	
                    color_dict["Config"] = "W"
                except:
                    pass

                list_of_key["Start time"] = time.strftime("%Y-%m-%d %H:%M:%S",
                                                         time.localtime(int(self.jsondata["start_time_millis"]) / 1000))	
                list_of_key["End time"] = time.strftime("%Y-%m-%d %H:%M:%S",
                                                         time.localtime(int(self.jsondata["end_time_millis"]) / 1000))
                try:                               
                    # list_of_key["Test time"] = str(timedelta(seconds=float(self.jsondata["duration_seconds"]))).split('.')[0]  #second to 00:00:00
                    list_of_key["Test time"] = float(self.jsondata["duration_seconds"])
                except:
                    traceback.print_exc()

                try:                                        
                    list_of_key["slot"] = float(self.jsondata["slot"])
                except:
                    list_of_key["slot"] = self.jsondata["slot"]
                
                if self.jsondata["metadata"]["scof_on"] == "True":
                    list_of_key["Outcome"] = 'SCOF_' + self.jsondata["outcome"]
                else: list_of_key["Outcome"] = self.jsondata["outcome"]
                #---------------------------------------------------
                color_dict["SN"] = "W"
                color_dict["Test Mode"] = "W"
                color_dict["Test Version"] = "W"
                color_dict["BFT"] = "W"
                color_dict["Start time"] = "W"
                color_dict["End time"] = "W"
                color_dict["Test time"] = "W"
                color_dict["slot"] = "W"
                color_dict["Outcome"] = "W"


                #---------------------------------------------------
                station = self.jsondata["metadata"]["test_name"]
                self.sheet_name.setdefault(station, [])  #將Station name 加入Sheet_name dict
                self.sheet_color.setdefault(station, [])
                
                #if self.chk_opValue.get():
                try:
                    list_of_key["OP_ID"] = self.jsondata["metadata"]["op_id"]
                    color_dict["OP_ID"] = "W"
                except:
                    pass

                        # Random OP
                        #op_list = ['V0947667', 'V0973158', 'V0973555']
                        #list_of_key["OP_ID"] = random.choices(population=op_list, weights=self.weight)
                        #pos = self.op_list.index(list_of_key["OP_ID"][0])
                        #for i in range(len(self.weight)):
                            #if pos == i:								
                                #self.weight[i] -= 0.3
                            #else:								
                                #self.weight[i] += 0.15
                        #print(self.weight)


                Failure_code = ""
                failItemsDes = ""
                list_of_key["Failure Code"] = Failure_code
                list_of_key["Failure Items"] = failItemsDes
                color_dict["Failure Code"] = "W"
                color_dict["Failure Items"] = "W"
                #-----------------------------------------------------
                if self.jsondata["outcome"] != "PASS":
                    cnt_3_times_fail_sn.setdefault(self.jsondata["dut_id"], 0)
                    cnt_3_times_fail_sn[self.jsondata["dut_id"]] += 1 
                    # print(self.jsondata["dut_id"], list_of_key["Start time"])

                    if cnt_3_times_fail_sn[self.jsondata["dut_id"]] == 3:
                        # print(self.jsondata["dut_id"], list_of_key["Start time"], "fail")
                        color_dict["SN"] = "SN_R"
                #------------------------------------------------------
                if re.search("Mars", self.jsondata["metadata"]["framework"]):
                    for index in range(len(self.jsondata["phases"])):
                        if self.jsondata["phases"][index]["measurements"] is not None:
                            for item in range(len(self.jsondata["phases"][index]["measurements"])):
                                key = self.jsondata["phases"][index]["measurements"][item]["name"]
                                special_key = self.jsondata["phases"][index]["measurements"][item]["name"] + " spec"
                                try:
                                    list_of_key[key] = self.bool_float_check(
                                        self.jsondata["phases"][index]["measurements"][item]["measured_value"])
                                except:
                                    pass

                                try: 
                                    # list_of_key[special_key] = self.jsondata["phases"][index]["measurements"][item]["validators"][0]
                                    if self.jsondata["metadata"]["scof_on"] == "True" and len(self.jsondata["phases"][index]["measurements"][item]["validators"]) ==2 :
                                        list_of_key[special_key] = self.jsondata["phases"][index]["measurements"][item]["validators"][1]
                                    else: 
                                        list_of_key[special_key] = self.jsondata["phases"][index]["measurements"][item]["validators"][0]
                                except:
                                    pass
                                
                                if self.jsondata["phases"][index]["measurements"][item]["outcome"] == "FAIL":
                                    failItemsDes +=  "{}={}, [{}]\n".format(key, list_of_key[key], ", ".join(list_of_key[special_key]))
                                    Failure_code += "{}\n".format(key)

                                    list_of_key["Failure Code"] = Failure_code
                                    list_of_key["Failure Items"] = failItemsDes
                                    
                                    




                                color_dict[key] = self.define_color(self.jsondata["phases"][index]["measurements"][item]["outcome"])
                                color_dict[special_key] = "W"


                else:
                    #------------------------------------------------------
                    for index in range(len(self.jsondata["phases"])):
                        if self.jsondata["phases"][index]['name'] == 'teardown': break
                        for key in self.jsondata["phases"][index]["measurements"].keys():
                            special_key = key +" spec"
                            #print(self.jsondata["phases"][index]["measurements"][key])
                            
                            if self.check_key_ifexist(self.jsondata, index, key, "measured_value"):
                                list_of_key[key] = self.jsondata["phases"][index]["measurements"][key]["measured_value"]
                                list_of_key[key] = self.bool_float_check(list_of_key[key])

                            if self.check_key_ifexist(self.jsondata, index, key, "validators"):
                                if self.jsondata["metadata"]["scof_on"] == "True" and len(self.jsondata["phases"][index]["measurements"][key]["validators"]) == 2:
                                    list_of_key[special_key] = self.jsondata["phases"][index]["measurements"][key]["validators"][1]
                                else: list_of_key[special_key] = self.jsondata["phases"][index]["measurements"][key]["validators"][0]

                            if self.jsondata["phases"][index]["measurements"][key]["outcome"] == "FAIL":
                                #print(key, list_of_key[key], list_of_key[special_key])
                                failItemsDes +=  "{}={}, [{}]\n".format(key, list_of_key[key], ", ".join(list_of_key[special_key]))
                                Failure_code += "{}\n".format(key)

                                list_of_key["Failure Code"] = Failure_code
                                list_of_key["Failure Items"] = failItemsDes
                            
                            color_dict[key] = self.define_color(self.jsondata["phases"][index]["measurements"][item]["outcome"])
                            color_dict[special_key] = "W"

                
                #print(list_of_key)
                #print(pd.DataFrame(list(list_of_key.items())))

                # self.sheet_name[station].append(pd.DataFrame(data=list_of_key, index=[len(self.sheet_name[station])]))           
                # self.sheet_color[station].append(pd.DataFrame(data=color_dict, index=[len(self.sheet_color[station])]))

                # 先存 list_of_key (dict) 進去 sheet_name 
                self.sheet_name[station].append(list_of_key)
                self.sheet_color[station].append(color_dict)

        # Convert list of dictionaries to a pandas DataFrame
        # https://stackoverflow.com/questions/20638006/convert-list-of-dictionaries-to-a-pandas-dataframe
        for station in self.sheet_name.keys():
            self.sheet_name[station] = pd.DataFrame(self.sheet_name[station])
            self.sheet_color[station] = pd.DataFrame(self.sheet_color[station])
    
    def write_summary(self):
        self.set_output_summary_filename(self.sheet_name.keys())
        print('===== Writing to ' + self.summary_outputname  + '=====')     
        
        try:
            with pd.ExcelWriter(self.summary_outputname, engine='xlsxwriter') as writer:  
                workbook = writer.book

                failformat = workbook.add_format({
                    'bg_color': '#FFFF66',  # your setting
                    'bold': True,           # additional stuff...
                    'text_wrap': False,
                    'valign': 'bottom',
                    # 'align': 'left',
                    'border': 1})
                
                scof_failformat = workbook.add_format({
                    'bg_color': '#FFA500',  # your setting
                    'bold': True,           # additional stuff...
                    'text_wrap': False,
                    'valign': 'bottom',
                    # 'align': 'left',
                    'border': 1})
                
                sn_failformat = workbook.add_format({
                    'bg_color': '#CC0000',  # your setting
                    'bold': False,           # additional stuff...
                    'text_wrap': False,
                    'valign': 'bottom',
                    # 'align': 'left',
                    'border': 1})

                for key in self.sheet_name.keys():
                    self.work_status = Summary_WorkStatus.GENERATE.value
                    print("Sheet name: " , key)
                    
                    self.sheet_name[key].to_excel(writer, sheet_name=key, encoding='utf8') #一次寫入
                    	
                    #---------------------------------------------------------------
                    # color setting
                    self.work_status = Summary_WorkStatus.COLORING.value	
                    # color_result = pd.concat(self.sheet_color[key], axis=0)
                    
                    worksheet = writer.sheets[key]
                    # for col_num, value in enumerate(color_result.columns.values):
  
                    

                    # print("shape", color_result.shape)
                    for i in range(0, self.sheet_color[key].shape[0]):
                        # print(key, i)
                        color_rowSeries = self.sheet_color[key].iloc[i]
                        result_rowseries = self.sheet_name[key].iloc[i]
                        col = 0
                        for value in color_rowSeries:
                            col += 1
                            # result_rowseries[col]
                            if value == "R":
                                print(i, col)
                                print(color_rowSeries[col - 1])
                                worksheet.write(i + 1, col, result_rowseries[col-1], failformat)
                            elif value == "SCOF_R":
                                worksheet.write(i + 1, col, result_rowseries[col-1], scof_failformat)

                            elif value == "SN_R":
                                worksheet.write(i + 1, col - 1, result_rowseries[col-2], sn_failformat)

                        print(color_rowSeries.values)
                    #---------------------------------------------------------------

                    self.work_status = Summary_WorkStatus.FINISH.value

                print('Saved!')
            xl = Dispatch("Excel.Application")
            xl.Visible = True # otherwise excel is hidden
            xl.Workbooks.Open(os.getcwd() + '\\' + self.summary_outputname)
          

            
            

        except Exception as error:
            print(error)

    def define_color(self, value):
        if self.jsondata["metadata"]["scof_on"] == "True":
            if value.upper() != "PASS" and value.upper() != "UNSET":
                return "SCOF_R"
        elif value.upper() != "PASS" and value.upper() != "UNSET":
                return "R"
        else:
            return "W"

if __name__ == '__main__':
    print('Version 1.2.2')
    summ = RUN_SUMMARY(log_path='C:/Users/832816/Desktop/ftp_tool/output/logs/2021-10-26')
    summ.start()
    # user_interface = UI()
    # user_interface.func()
    """
    V1.2   修改bool, float型態轉換問題
    V1.2.1 json 資料夾可有可無
    V1.2.2 新增OP_ID，目前只有GRR的logs有提供
    """