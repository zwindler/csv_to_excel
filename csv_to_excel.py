#!/usr/bin/python
# Tool to convert CSV files (with configurable delimiter and text wrap
# character) to Excel spreadsheets.
################################################################################
# Found on http://sujitpal.blogspot.fr/2007/02/python-script-to-convert-csv-files-to.html
# Requires setup.py install of pyExcelerator found on
# https://pypi.python.org/pypi/pyExcelerator
# Modified to concat csv files in different sheets
################################################################################
import string
import sys
import getopt
import re
import os
import os.path
import csv
from pyExcelerator import *
 
def usage():
  ''' Display the usage '''
  print 'Usage:' + sys.argv[0] + ' [OPTIONS] csvfile'
  print 'OPTIONS:'
  print '--title|-t: If set, the first line is the title line'
  print '--lines|-l n: Split output into files of n lines or less each'
  print '--sep|-s c [def:,] : The character to use for field delimiter'
  print '--output|o: output file name/pattern, default: output.xls'
  print '--help|h: print this information'
  sys.exit(2)
 
def writeExcelHeader(worksheet, titleCols):
  ''' Write the header line into the worksheet '''
  cno = 0
  for titleCol in titleCols:
    worksheet.write(0, cno, titleCol)
    cno = cno + 1
 
def writeExcelRow(worksheet, lno, columns):
  ''' Write a non-header row into the worksheet '''
  cno = 0
  for column in columns:
    worksheet.write(lno, cno, column)
    cno = cno + 1
 
def closeExcelSheet(workbook, outputFileName):
  ''' Saves the in-memory WorkBook object into the specified file '''
  workbook.save(outputFileName)
 
def renameOutputFile(outputFileName, fno):
  ''' Renames the output file name by appending the current file number
      to it '''
  dirName, baseName = os.path.split(outputFileName)
  rootName, extName = os.path.splitext(baseName)
  backupFileBaseName = string.join([string.join([rootName, str(fno)], '-'), extName], '')
  backupFileName = os.path.join(dirName, backupFileBaseName)
  try:
    os.rename(outputFileName, backupFileName)
  except OSError:
    print 'Error renaming output file:', outputFileName, 'to', backupFileName, '...aborting'
    sys.exit(-1)
 
def validateOpts(opts):
  ''' Returns option values specified, or the default if none '''
  titlePresent = False
  linesPerFile = -1
  outputFileName = ''
  sepChar = ','
  for option, argval in opts:
    if (option in ('-t', '--title')):
      titlePresent = True
    if (option in ('-l', '--lines')):
      linesPerFile = int(argval)
    if (option in ('-s', '--sep')):
      sepChar = argval
    if (option in ('-o', '--output')):
      outputFileName = argval
    if (option in ('-h', '--help')):
      usage()
  return titlePresent, linesPerFile, sepChar, outputFileName
 
def main():
  ''' Main function '''
  try:
    opts,args = getopt.getopt(sys.argv[1:], 'tl:s:o:h', ['title', 'lines=', 'sep=', 'output=', 'help'])
  except getopt.GetoptError:
    usage()
  if (len(args) < 1):
    usage()
  ''' Opens a reference to an Excel WorkBook and Worksheet objects '''
  workbook = Workbook()
  for inputFileName in args:
    try:
      inputFile = open(inputFileName, 'r')
    except IOError:
      print 'File not found:', inputFileName, '...skipping'
      continue
    titlePresent, linesPerFile, sepChar, outputFileName = validateOpts(opts)
    if (outputFileName == ''):
      outputFileName = 'output.xls'
    worksheet = workbook.add_sheet(os.path.basename(inputFileName))
    fno = 0
    lno = 0
    titleCols = []
    reader = csv.reader(inputFile, delimiter=sepChar)
    for line in reader:
      if (lno == 0 and titlePresent):
        if (len(titleCols) == 0):
          titleCols = line
        writeExcelHeader(worksheet, titleCols)
      else:
        writeExcelRow(worksheet, lno, line)
      lno = lno + 1
      if (linesPerFile != -1 and lno >= linesPerFile):
        closeExcelSheet(workbook, outputFileName)
        renameOutputFile(outputFileName, fno)
        fno = fno + 1
        lno = 0
        worksheet = workbook.add_sheet(os.path.basename(inputFileName))
    inputFile.close()
    closeExcelSheet(workbook, outputFileName)
    if (fno > 0):
      renameOutputFile(outputFileName, fno)
 
if __name__ == '__main__':
  main()
