import xlrd
fb=open('./macros.bas','w')
data=xlrd.open_workbook('./auto.xlsx')
table=data.sheets()[0]
nrows=table.nrows
fb.write('  ')
for i in range(nrows):
    fb.write('Selection.Find.ClearFormatting\n\
    Selection.Find.Replacement.ClearFormatting\n\
    With Selection.Find\n\
        .Text = "'+table.cell_value(i,0)+'"\n\
        .Replacement.Text = "'+table.cell_value(i,1)+'"\n\
        .Forward = True\n\
        .Wrap = wdFindContinue\n\
        .Format = False\n\
        .MatchCase = False\n\
        .MatchWholeWord = False\n\
        .MatchByte = True\n\
        .MatchWildcards = False\n\
        .MatchSoundsLike = False\n\
        .MatchAllWordForms = False\n\
    End With\n\
    Selection.Find.Execute Replace:=wdReplaceAll\n\
    ')
fb.close()
