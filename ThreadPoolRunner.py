# coding=UTF-8
__author__ = 'user'

import threading
from threading import Thread

class ThreadPoolRunner():

    def __init__(self):
        self.lock = threading.Lock()
        self.threadList = []
        self.runningList = []
        pass

    def addWorker(self, func, resultList, args=(), lockMode=False):
        t = threading.Thread(target=self.threadRunner, args=(len(self.threadList), lockMode,func, resultList, args))
        self.threadList.append(t)
        pass

    def threadRunner(self, id, lockMode, func, resultList, args=()):
        result = None
        try:
            if lockMode: self.lock.acquire()
            result = func(*args)
        finally:
            resultList.append([id, result])
            if lockMode: self.lock.release()

    def clearWorkers(self):
        self.waitWorkers()
        self.threadList[:] = []

    def runWorkers(self, tList):
        for thread in tList:
            thread.start()
            self.runningList.append(thread)

    def waitWorkers(self):
        for thread in self.runningList:
            thread.join()

        self.runningList[:] = []

    def runAllWorkerAndWait(self, max=None):
        if max is None: max = len(self.threadList)

        for i in xrange(0,len(self.threadList),max):
            tList = self.threadList[i:i+max]

            self.runWorkers(tList)
            self.waitWorkers()


