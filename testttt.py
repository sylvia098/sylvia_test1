# num_list = []
# num_list.append(11)
# num_list.append(22)
# for i in range(6):
#     num_list.append(i)
# num_list = [i for i in range(4) if i % 2 == 0] #list comprehension
# # num_list = list(range(6))
# print(num_list)

# def ans():
#     print(42)
# def run(func):
#     func()
# run(ans)

# def add_args(arg1, arg2):
#     print(arg1+arg2)

# def run_add_args(func, arg1, arg2):
#     func(arg1, arg2)
# run_add_args(add_args, 2,3)

# def outer(a, b):
#     def inner(c, d):
#         return c + d
#     print (inner(a,b))

# outer(2,4)

#--------內部函式---------------
# def knights(say):
#     def inner(quote):
#         print("I like panda.'%s'" % quote)
#     return inner(say)

#--------Closure---------------
# def knights2(say):
#     def inner2():
#         print("happy,'%s'" % say)
#     return inner2

# a = knights2('frog')
# b = knights2('sheep')
# a()
# b()

# import paramiko
# from contextlib import suppress
# import traceback
# import re
# ssh_client = paramiko.SSHClient()
# ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
# with suppress(paramiko.ssh_exception.AuthenticationException):
#         ssh_client.connect('192.168.11.20', username = 'root', password = '')
# ssh_client.get_transport().auth_none('root')
# def Wifi_Check(freq):
#     def inner3():
#         stdin, stdout, stderr = ssh_client.exec_command(freq)
#         stdin.flush()
#         wifi_list = stdout.readlines()
#         try:
#             filtered_values = list(filter(lambda m: re.match(".*iPhone098", m), wifi_list))
#             if filtered_values:
#                 str_filtered_values =''.join(filtered_values)
#                 filtered_values_data = wifi_list.index(str_filtered_values)
#                 signal_level = wifi_list[filtered_values_data + 1]
#                 # signal_level_value = signal_level[48:51]
#                 get_signal_level_value = re.compile('Signal level=(-+\d+)')
#                 signal_level_value = get_signal_level_value.search(signal_level)
#                 signal_level_value = signal_level_value.group(1)
#                 signal_level_value = int(signal_level_value)
#                 # print(signal_level)
#                 # print(signal_level_value)
#                 # print(str_filtered_values)
#                 return str_filtered_values, signal_level_value
#             else:
#             # print('error')
#                 return None
#         except:
#             traceback.print_exc()
#     return inner3
# fiveG = Wifi_Check('5g')
# twoG = Wifi_Check('2g')
# fiveG()

#--------Generator---------------
# def my_range(first = 0, last = 10, step =2):
#     number = first
#     while number < last:
#         yield number
#         number += step
# ranger = my_range(0, 6)
# for i in ranger:
#     print(i)

#--------Decorator---------------
# def document_it(func):
#     def new_doc(*args, **kwargs):
#         print('running func:', func.__name__)
#         print('positional arguments', args)
#         print('keyword arguments:', kwargs)
#         result = func(*args, **kwargs)
#         print(result)
#         return result
#     return new_doc
# def square_it(func):
#     def new_doc(*args, **kwargs):
#         result = func(*args, **kwargs)
#         result = result * result
#         print(result)
#     return new_doc 
# @document_it
# @square_it
# def add_ints(a, b):
#     return a + b
# add_ints(3,4)

# periodic_table = {'a': 11, 'b': 22}
# abc = periodic_table.setdefault('c', 33) #沒在字典裡就加入，變成新值
# print(abc)
# print(periodic_table)
# cba = periodic_table.setdefault('a', 33) #回傳原始值
# print(cba)

# import collections
# quote = collections.OrderedDict([('sandy','she is a girl'),
#                                  ('mandy','she is a woman'),
#                                  ('ray','he is a man')])
# for i, j in quote.items(): #只給一個參數的話就會變成tuple回傳
#     print(i)

# import itertools
# aaa = [11,22,33,44,55,66]
# for i in itertools.chain(aaa):
#     print(i)

# class Person():
#     def __init__(self, name):
#         self.name = name
# class EmailPerson(Person):
#     def __init__(self, name, email):
#         super().__init__(name)   #用到父類別 super()
#         self.email = email
# bobo = EmailPerson('bobo', 'bobo@mail.com')
# print(bobo.name)
# print(bobo.email)

# class A():
#     count = 0
#     def __init__(self):
#         A.count += 1
#     def exclaim(self):
#         print("AAA")
#     @classmethod
#     def kids(cls):
#         print("A has", cls.count, "things")
# easyA = A()
# midA = A()
# diffA = A()
# print(A.kids())

# class Car:
#     def __init__(self, weight):
#        self.weight = weight

#     @property
#     def weight(self):
#         return self.__weight
    
#     @weight.setter
#     def weight(self, value):
#         if value <= 0:
#             raise ValueError("Weight can not be 0 or less.")
#         self.__weight = value
# car1 = Car(200)

# class Bank_Account:
#     def __init__(self):
#         self._password = '預設是0000'

#     @property
#     def password(self):
#         return self._password
#     @password.setter
#     def password(self, value):
#         self._password = value
    

# qoo_account = Bank_Account()
# qoo_account.password = '123'
# print(qoo_account.password)

# import subprocess
# ret = subprocess.getstatusoutput('date')
# print(ret)

# from datetime import timedelta, date
# right_now = date.today()
# a_day = timedelta(1)
# tomorrow = right_now + a_day
# print(right_now)
# print(tomorrow)

# import multiprocessing as mp
# def Washer(dishes, output):
#     for dish in dishes:
#         print('washing', dish, 'dish')
#         output.put(dish)
# def Dryer(input):
#     while True:
#         dish = input.get()
#         print('drying', dish, 'dish')
#         input.task_done()
# if __name__ == '__main__':
#     dish_queue =  mp.JoinableQueue()
#     dry_proc = mp.Process(target = Dryer, args = (dish_queue,))
#     dry_proc.daemon = True
#     dry_proc.start()
#     dishes = ['salad', 'main', 'cake', 'beverage']
#     Washer(dishes, dish_queue)
#     dish_queue.join()

import os
from win32com.client import Dispatch
xl = Dispatch("Excel.Application")
xl.Visible = True # otherwise excel is hidden

# wb = xl.Workbooks.Open(r'C:\Users\832816\Desktop\ftp_tool\summary_output_2020-10-22-17-56-18.xlsx')
wb = xl.Workbooks.Open(os.getcwd() + '\summary_output_2020-10-22-17-56-18.xlsx')