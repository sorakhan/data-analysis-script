#!/usr/bin/env python

# AUTHOR: Sora Khan
# DATE:   17 May 2021 (4hrs)
# DESC:   This script automates the process of analyzing data on multiple excel spreadsheets, to save hours of work.

import openpyxl
import os

def extractIntervals(max_rows):
  block_interval = 10
  new_interval = False
  intervals = []  # [[2,6],[7,11]] .. etc, so can refer to column G
  block = []      # [2,7] but range will be 2 -> 6 ;)

  for row in range(2,max_rows):
    b_col_val = sheet[f'B{row}'].value
    lastRow = row == max_rows - 1

    if (b_col_val > block_interval): # starting a new interval section, so finish up previous interval
      end = sheet[f'B{row-1}'].value
      block.append(row)
      # print(f'{end}, [B{row-1}] --- end\n')

      intervals.append(block)
      block = []
      block_interval += 10
      new_interval = True

    if (new_interval or row == 2): # set start interval
      # print(f'\n{b_col_val}, [B{row}] --- start')
      block.append(row)
      new_interval = False

    if (lastRow):
      # print(f'{b_col_val}, [B{row+1}] --- end')
      block.append(max_rows + 1)
      intervals.append(block)    

  return intervals

def calcResponseTimeAvg():
  # print(intervals)
  avg_resp_time = []

  for i in range(0,len(intervals)):
    count = 0.0
    total = 0
    for row in range (intervals[i][0], intervals[i][1]):
      col_val = sheet[f'G{row}'].value
      if (col_val != 0):
        count += 1
        total += col_val

    if (count != 0):
      total = total / count

    avg_resp_time.append(total)

  return avg_resp_time

def calcCrashes():
  # print(intervals)
  crashes = []
  for i in range(0,len(intervals)):
    start = intervals[i][0]
    end = intervals[i][1] - 1
    start_val = sheet[f'F{start}'].value
    end_val = sheet[f'F{end}'].value
    total = end_val - start_val
    # print(f'\n{i*10}: {total}')
    crashes.append(total)
    
  return crashes

def clearTable(fileName):
  cols = ['M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
  for letter in cols:
    for row in range(1,40): # M -> T
      sheet[f'{letter}{row}'].value = ''
  wb.save(fileName)


def addTableValues(fileName):
  # this is temp for testing.. should really be N,O,P
  sheet['N1'].value = 'Summary'
  sheet['N2'].value = 'In Game Time'
  sheet['O2'].value = 'Crashes for that period'
  sheet['P2'].value = 'Average response time'

  crashes = calcCrashes()
  avg_resp = calcResponseTimeAvg()

  row = 3
  for i in range(0, len(intervals)):
    sheet[f'N{row}'].value = i*10
    sheet[f'O{row}'].value = crashes[i]
    sheet[f'P{row}'].value = avg_resp[i]

    row += 1

  wb.save(fileName)

# Prints Table's output in console  
def tableOutput():
  crashes = calcCrashes()
  avg_resp = calcResponseTimeAvg()
  print('\n================================================')
  print(f'GT \t CRASHES \t\t AVG_RESP')
  print('\n================================================')

  for i in range(0, len(intervals)):
    print(f'{i*10} \t {crashes[i]} \t\t {avg_resp[i]}')

def run(fileName): 
  global wb
  global sheet 
  global intervals

  wb = openpyxl.load_workbook(fileName)
  sheet = wb.active
  # sheetName = sheet.title
  max_rows = sheet.max_row
  intervals = extractIntervals(max_rows)

  print(f'\nFINISHED: {fileName}')
  # print(f'ROWS: {max_rows}')
  # print(f'INTERVALS: {len(intervals)}')

  # tableOutput()
  clearTable(fileName)
  addTableValues(fileName)

if __name__ == "__main__":
  while True:
    print('\n============================================')
    filePath = input('\nEnter full file path / folder name: ')
    os.chdir(filePath)
    excelFiles = os.listdir('.')
    print(f'\n{len(excelFiles)} excel files found.')
    for i in range(0, len(excelFiles)):
      if not excelFiles[i].endswith('.xlsx'):
        print('\tNot a valid folder for .xlsx files')
        break
      run(excelFiles[i]) # end reached because all are excel files
