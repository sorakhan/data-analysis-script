#!/usr/bin/env python

# AUTHOR: Sora Khan
# DATE:   17 May 2021 (4hrs)
# DESC:   This script automates the process of analyzing data on multiple excel spreadsheets, to save hours of work.

import openpyxl
import os
import math

def extractIntervals(max_rows):
  # increment by 10 for range
  block_interval = 10
  new_interval = False
  intervals = [] # [[2,6],[7,11]] .. etc, so can refer to column G
  block = [] # [2,7] but range will be 2 -> 6 ;)

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

def calcResponseTimeAvg(withRBs):
  # print(intervals)
  avg_resp_time = []

  for i in range(0,len(intervals)):
    count = 0.0
    total = 0
    for row in range (intervals[i][0], intervals[i][1]):
      if(withRBs == "Y" or withRBs == "y"):
        col_val = sheet[f'H{row}'].value
      else:
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

def rbCrashes():
  # print(intervals)
  rbCrashes = []
  for i in range(0,len(intervals)):
    start = intervals[i][0]
    end = intervals[i][1] - 1
    start_val = sheet[f'G{start}'].value
    end_val = sheet[f'G{end}'].value
    total = end_val - start_val
    # print(f'\n{i*10}: {total}')
    rbCrashes.append(total)
    
  return rbCrashes

def getDifficulty():
  difficulty = []
  for i in range(0,len(intervals)):
    count = 0.0
    total = 0
    for row in range (intervals[i][0], intervals[i][1]):
      col_val = sheet[f'C{row}'].value
      if (col_val != 0):
        count += 1
        total += col_val

    if (count != 0):
      total = total / count

    difficulty.append(total)

  return difficulty

def clearTable(fileName):
  cols = ['K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']
  for letter in cols:
    for row in range(1,50): # M -> T
      sheet[f'{letter}{row}'].value = ''
  wb.save(fileName)

def getHalfCrashes(start, end):
  sum = 0
  for row in range (start, end):
      sum += sheet[f'L{row}'].value

  return sum

def getHalfDRT(start, end):
  count = 0
  sum = 0
  for row in range (start, end):
      col_val = sheet[f'N{row}'].value
      if(col_val!=0):
        count += 1
        sum += col_val

  return sum/count

def getHalfSD(start, end):
  mean = getHalfDRT(start, end)
  count = 0
  deviations = 0
  for row in range (start, end):
      col_val = sheet[f'N{row}'].value
      if(col_val!=0):
        count += 1
        deviations += math.pow((col_val - mean), 2)

  sd = math.sqrt(deviations/(count-1))
  return sd

def addTableValues(fileName, withRBs):

  # This adds the first summary table 
  sheet['K1'].value = 'Summary'
  sheet['K2'].value = 'In Game Time'
  sheet['L2'].value = 'Crashes for that period'
  sheet['M2'].value = 'RB crashes for that period'
  sheet['N2'].value = 'Average response time'
  sheet['O2'].value = 'Difficulty Level'

  # Calculations for the first summary table
  crashes = calcCrashes()
  rbs = rbCrashes()
  avg_resp = calcResponseTimeAvg(withRBs)
  difficulty = getDifficulty()

  row = 3
  for i in range(0, len(intervals)):
    sheet[f'K{row}'].value = i*10
    sheet[f'L{row}'].value = crashes[i]
    if(withRBs == "Y" or withRBs == "y"):
      sheet[f'M{row}'].value = rbs[i]
    else:
      sheet[f'M{row}'].value = -1
    sheet[f'N{row}'].value = avg_resp[i]
    sheet[f'O{row}'].value = difficulty[i]

    row += 1
  
  # This adds the second summary table 
  sheet['R1'].value = '0-119.99'
  sheet['U1'].value = '120-300'
  sheet['Q2'].value = 'Participant Id'
  sheet['R2'].value = 'Total Crashes'
  sheet['S2'].value = 'Mean DRT'
  sheet['T2'].value = 'SD DRT'
  sheet['U2'].value = 'Total crashes'
  sheet['V2'].value = 'Mean DRT'
  sheet['W2'].value = 'SD DRT'

  # Participant Id
  participantId = fileName[fileName.index("-")+1:fileName.rindex("-")]
  sheet['Q3'].value = participantId

  # 0-119.99 Number of Crashes
  firstHalfCrashes = getHalfCrashes(3, 15)
  sheet['R3'].value = firstHalfCrashes

  # 0-119.99 Mean DRT
  firstHalfDRT = getHalfDRT(3, 15)
  sheet['S3'].value = firstHalfDRT
  
  # 0-119.99 SD DRT
  firstHalfSD = getHalfSD(3, 15)
  sheet['T3'].value = firstHalfSD
  
  # 120-300 Number of Crashes
  secondHalfCrashes = getHalfCrashes(15, 33)
  sheet['U3'].value = secondHalfCrashes
  
  # 120-300 Mean DRT
  secondHalfDRT = getHalfDRT(15, 33)
  sheet['V3'].value = secondHalfDRT

  # 120-300 SD DRT
  secondHalfSD = getHalfSD(15, 33)
  sheet['W3'].value = secondHalfSD

  wb.save(fileName)

def run(fileName, withRBs): 
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

  # Clears any existing data
  clearTable(fileName)

  # Add new calculated data
  addTableValues(fileName, withRBs)
  
def tableOutput():
  crashes = calcCrashes()
  avg_resp = calcResponseTimeAvg()
  print('\n================================================')
  print(f'GT \t CRASHES \t\t AVG_RESP')
  print('\n================================================')

  for i in range(0, len(intervals)):
    print(f'{i*10} \t {crashes[i]} \t\t {avg_resp[i]}')

if __name__ == "__main__":
  while True:
    # User input
    print('\n============================================')
    print('\nPlease make sure the excel files are closed. ')
    filePath = input('\nEnter full file path / folder name: ')
    os.chdir(filePath)

    # Determine if we need to also calculate roadblocks
    withRBs = input('\nWith Roadblocks? (Y / N): ')
    if(withRBs == "Y" or withRBs == "y"):
      print("Roadblocks will be calculated.")
    else:
      print("Roadblocks will not be calculated.")

    # Open the folder location and find the excel files
    excelFiles = os.listdir('.')
    print(f'\n{len(excelFiles)} excel files found.')
    for i in range(0, len(excelFiles)):
      # if (i == len(excelFiles) - 1 and excelFiles[i].endswith('.xlsx')):
      #   return False
      if not excelFiles[i].endswith('.xlsx'):
        print('\tNot a valid folder for .xlsx files')
        break
      # end reached, so all excel files
      run(excelFiles[i], withRBs)

# [] grey alternate cells
# [] be able to enter or select file name
# [] Make .exe file .. bash script?
