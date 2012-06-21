#!/usr/local/bin/python
# -*- coding: utf-8 -*-
'''solver_excel_ole
1: K 2: W 3: R 4: G 5: B 6: Y 7: M 8: C
※ 100 x 100 のときは maximum recursion depth exceeded に達するので limit 変更
mat(r, c) を毎回実行するより m = mat(r, c) とした方が速そうだが Excel 側で
メモリ不足になる？ (毎回 del m すれば解決するかも知れないが再帰の都合で無理？)
maze_excel_ole で作ったサイズと合わせておく必要がある (自動サイズ検出未対応)
'''

import sys, os, random
import pywintypes, win32com.client

sys.setrecursionlimit(sys.getrecursionlimit() * 20) # ※
VISIBLE = True
TITLE = u'迷路'
HEIGHT, WIDTH, OFFSET_ROW, OFFSET_COL = 8, 16, 2, 2
MAX_ROW, MAX_COL = OFFSET_ROW + HEIGHT - 1, OFFSET_COL + WIDTH - 1
sheet = None

def mat(r, c):
  return sheet.Cells(OFFSET_ROW + r, OFFSET_COL + c)

def isExit(r, c):
  return r == HEIGHT - 1 and c == WIDTH - 1

def isWall(r, c, d):
  if mat(r, c).Borders(1 + d).LineStyle == 1: return True
  dr = 1 if d == 3 else -1 if d == 2 else 0
  dc = 1 if d == 1 else -1 if d == 0 else 0
  if mat(r + dr, c + dc).Interior.ColorIndex in [7, 8]: return True # M or C
  return False

def isDeadendWall(r, c, direc):
  if direc < 0: return False
  for d in xrange(4):
    if [1, 0, 3, 2][direc] == d: continue
    if not isWall(r, c, d): return False
  else:
    return True

def drawPath(r, c, solved, branch):
  mat(r, c).Interior.ColorIndex = 8 if solved and not branch else 7 # C or M
  return solved

def dug(r, c, direc, solved, branch):
  dlist = [0, 0, 0, 0]
  if direc >= 0: dlist[[1, 0, 3, 2][direc]] = 1
  mat(r, c).Interior.ColorIndex = 6 # Y
  while True:
    if isExit(r, c): solved = True
    if isDeadendWall(r, c, direc): return drawPath(r, c, solved, branch)
    d = random.randint(0, 3)
    if dlist[d] == 1: continue
    dlist[d] = 1
    if isWall(r, c, d): continue
    dr = 1 if d == 3 else -1 if d == 2 else 0
    dc = 1 if d == 1 else -1 if d == 0 else 0
    solved = dug(r + dr, c + dc, d, solved, solved)

def solver_excel_ole(filename):
  global sheet
  xl = win32com.client.Dispatch('Excel.Application')
  xl.Visible = VISIBLE
  try:
    book = xl.Workbooks.Open(filename)
    sheet = book.Worksheets(1)
    rg = sheet.Range(
      sheet.Cells(OFFSET_ROW, OFFSET_COL), sheet.Cells(MAX_ROW, MAX_COL))
    random.seed()
    dug(0, 0, 3, False, False)
    book.Save()
  finally:
    xl.ScreenUpdating = True
    xl.Workbooks.Close()
  xl.Quit()

if __name__ == '__main__':
  solver_excel_ole(os.path.abspath(u'./%s.xls' % TITLE))
