import sqlite3

con = sqlite3.connect("TP/TuningParameter.sqlite")
cur = con.cursor()
TagName = '036PIA224_50'
sql_select = f""" 
                SELECT d.DataItemName, d.Value FROM Tag t, DataItem d 
                WHERE TagName = '{TagName}' and d.TagID = t.TagID and d.DataItemName in ('LL', 'PL', 'PH', 'HH') 
                ORDER BY 1 DESC
              """
res = cur.execute(sql_select)
setpoints = cur.fetchall()
print(len(setpoints))
test = [x[1] for x in setpoints]
print(test)
for sp in setpoints:
    print(sp)
con.close()
