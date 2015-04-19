# coding=UTF-8
__author__ = 'use'

import sys

from PearsonAnalysis import *


if __name__ == "__main__":
    if len(sys.argv) > 1:
        sourceFolder = sys.argv[1]
    else:
        sourceFolder = None
    print("sourceFolder=", sourceFolder)
    print("TextSliceMap.sln.gvid = ", TextSliceMap.sln.gvid)
    print("TextSliceMap.sln.var1 = ", TextSliceMap.sln.var1)
    pearsonRunner = PearsonAnalysis(sourceFolder)
    pearsonRunner.run(2)
