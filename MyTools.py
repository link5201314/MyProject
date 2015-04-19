# coding=UTF-8
import ctypes
import types
import locale
import random
import time
import subprocess

from itertools import imap

def pearsonr(x, y):
  # Assume len(x) == len(y)
  n = len(x)
  sum_x = float(sum(x))
  sum_y = float(sum(y))
  sum_x_sq = sum(map(lambda x: pow(x, 2), x))
  sum_y_sq = sum(map(lambda x: pow(x, 2), y))
  psum = sum(imap(lambda x, y: x * y, x, y))
  num = psum - (sum_x * sum_y/n)
  den = pow((sum_x_sq - pow(sum_x, 2) / n) * (sum_y_sq - pow(sum_y, 2) / n), 0.5)
  if den == 0: return 0
  return num / den

def list2str(l):
    str1=""
    # for item in list:
    #     str= str+item+","
    str1 = ",".join(str(s) for s in l)

    return str1

def randomPickNGroup(size, n):
    splitSize = size/n
    modSize = size%n
    #print(splitSize)
    s = set(range(1,size+1))
    #print(s)
    l = []
    for i in range(1,n+1):
        #print("i=" + str(i))
        if modSize != 0:
            splitSizeTmp = splitSize + 1
            modSize -= 1
        else:
            splitSizeTmp = splitSize

        subSet = []
        for j in range(1,splitSizeTmp+1):
            #print("j=" + str(j))
            selectNum = random.choice(list(s))
            s.remove(selectNum)
            subSet.append(selectNum)

        l.append(subSet)

    return l

def sleep(sec):
    time.sleep( sec )

def execCommand(cmdStr):
    #print("execute")
    p = subprocess.Popen(cmdStr, shell=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    retstr = ""
    for line in p.stdout.readlines():
        #print line.decode('cp950'),
        retstr += line.decode(locale.getdefaultlocale()[1])

    retval = p.wait()
    return [retval,retstr]

def getIfInteger(x):
    if type(x) is types.UnicodeType or type(x) is types.StringType:
        #print('type=str')
        try:
            if "." not in x:
                return int(x)
        except ValueError:
            #print(type(x))
            return None
    elif type(x) is types.IntType:
        #print('type=int')
        return x
    else:
        #print('type=else')
        #print(type(x))
        return None


def is_integer(x):
    if type(x) is types.UnicodeType or type(x) is types.StringType:
        #print('type=str')
        try:
            if "." not in x:
                int(x)
                return True
        except ValueError:
            #print(type(x))
            return False
    elif type(x) is types.IntType:
        #print('type=int')
        return True
    else:
        #print('type=else')
        #print(type(x))
        return False

def is_number_complex(x):
    #print('is_number_complex')
    return None


def getIfNumeric(s):
    try:
        c=float(s)
        n=str(c)
        if n=="nan" or n =="inf" or n=="-inf" : return None
    except ValueError:
        try:
            #print('try is_number_complex')
            #尚未實作
            c=is_number_complex(s)
            return None
        except ValueError, e:
            print('ValueError2')
            return None
    return c

def is_number(s):
    try:
        c=float(s)
        n=str(c)
        if n=="nan" or n =="inf" or n=="-inf" : return False
    except ValueError:
        try:
            #print('try is_number_complex')
            #尚未實作
            c=is_number_complex(s)
            if c is None:
                return False
        except ValueError, e:
            print('ValueError2')
            return False
    except Exception:
        print("is_numbe test fail: s = ",s)
    return True

def is_class(obj):
    obj_type = str(type(obj)).lower()
    flags = ['class', 'instance']
    return [flag for flag in flags if flag in obj_type]