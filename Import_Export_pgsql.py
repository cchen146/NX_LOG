import csv
import psycopg2
import win32com.client

from Report_Cleaning import get_latest_file

def truncate_n_update_tb(db, tb, raw_data, upload_sql):
    con = psycopg2.connect(db)
    cur = con.cursor()
    cur.execute("truncate table {}".format(tb))
    sql = upload_sql.format(tb)
    cur.execute(sql, raw_data)
    con.commit()
    con.close()

def truncate_n_upload_tb_fr_csv(db,sql_copy, tb, clean_file, sql_insert, **kwargs):
    con = psycopg2.connect(db)
    cur = con.cursor()
    cur.execute("truncate table {}".format(tb))
    sql = sql_copy.format(tb)
    nf = open(clean_file,'r',encoding = 'utf-8')
    cur.copy_expert(sql, nf)
    cur.callproc(sql_insert)
    nf.close()
    con.commit()
    con.close()

def fetch_n_write_csv(db, sql, tb_tempt):
    #store lookup table queried from database in a temporary table under tb_tempt path
    con = psycopg2.connect(db)
    cur = con.cursor()
    cur.execute(sql)
    mylist = cur.fetchall()
    myheader = [x.name for x in cur.description]
    myfile = open(tb_tempt, 'w', encoding='utf8', newline='')
    wr = csv.writer(myfile)
    wr.writerow(myheader)
    for row in mylist:
        wr.writerow(row)
    myfile.close()
    con.commit()


def trigger_update_macro(file_name, macro_to_run):
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = True
    wb = xl.Workbooks.Open(get_latest_file(file_name))
    for macro in macro_to_run:
        xl.Application.run(macro)
    wb.Save()
    wb.Close()
    xl.Application.Quit()
    del xl