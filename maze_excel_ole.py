#!/usr/local/bin/python
# -*- coding: utf-8 -*-
'''maze_excel_ole
http://nyaos.org/d/index.cgi?p=[Ruby]+win32ole
http://jp.rubyist.net/magazine/?0004-Win32OLE
1: K 2: W 3: R 4: G 5: B 6: Y 7: M 8: C
※ 100 x 100 のときは maximum recursion depth exceeded に達するので limit 変更
入口(左上)から掘ると迷路が簡単になる傾向があるので出口(右下)から掘るように修正
罫線描画のタイミングを変更することで無駄を減らし高速化
mat(r, c) を毎回実行するより m = mat(r, c) とした方が速そうだが Excel 側で
メモリ不足になる？ (毎回 del m すれば解決するかも知れないが再帰の都合で無理？)
'''

import sys, os, random
import pywintypes, win32com.client

sys.setrecursionlimit(sys.getrecursionlimit() * 20) # ※
VISIBLE = True
TITLE = u'迷路'
HEIGHT, WIDTH, OFFSET_ROW, OFFSET_COL = 100, 100, 2, 2
MAX_ROW, MAX_COL = OFFSET_ROW + HEIGHT - 1, OFFSET_COL + WIDTH - 1
sheet = None

def mat(r, c):
  return sheet.Cells(OFFSET_ROW + r, OFFSET_COL + c)

def isPassed(r, c):
  try:
    return mat(r, c).Interior.ColorIndex != 6 # Y
  except pywintypes.com_error, e:
    return True

def isDeadend(r, c):
  for d in xrange(4):
    dr = 1 if d == 3 else -1 if d == 2 else 0
    dc = 1 if d == 1 else -1 if d == 0 else 0
    if not isPassed(r + dr, c + dc): return False
  else:
    return True

def drawWall(r, c, dlist):
  e = mat(r, c)
  for d in xrange(4):
    if dlist[d] == 0: e.Borders(1 + d).Weight = 2

def dig(r, c, direc, count):
  dlist = [0, 0, 0, 0]
  if direc >= 0: dlist[[1, 0, 3, 2][direc]] = 1
  mat(r, c).Interior.ColorIndex = 4 # G
  count -= 1
  if count == 0: return drawWall(r, c, dlist)
  while True:
    if isDeadend(r, c): return drawWall(r, c, dlist)
    d = random.randint(0, 3)
    dr = 1 if d == 3 else -1 if d == 2 else 0
    dc = 1 if d == 1 else -1 if d == 0 else 0
    if not isPassed(r + dr, c + dc):
      dlist[d] = 1
      dig(r + dr, c + dc, d, count)

def maze_excel_ole(filename):
  global sheet
  xl = win32com.client.Dispatch('Excel.Application')
  xl.Visible = VISIBLE
  try:
    book = xl.Workbooks.Add()
    sheet = book.Worksheets(1)
    sheet.Name = TITLE
    sheet.Cells(1, 2).Value = TITLE
    rg = sheet.Range(
      sheet.Cells(OFFSET_ROW, OFFSET_COL), sheet.Cells(MAX_ROW, MAX_COL))
    rg.RowHeight = 5.18
    rg.ColumnWidth = 0.58
    rg.Interior.ColorIndex = 6 # Y
    random.seed()
    dig(HEIGHT - 1, WIDTH - 1, -1, WIDTH * HEIGHT)
    book.SaveAs(filename)
  finally:
    xl.ScreenUpdating = True
    xl.Workbooks.Close()
  xl.Quit()

if __name__ == '__main__':
  maze_excel_ole(os.path.abspath(u'./%s.xls' % TITLE))
