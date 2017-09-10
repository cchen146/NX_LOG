import psycopg2

def truncate_n_update_tb(db, tb, raw_data, upload_sql):
    con = psycopg2.connect(db)
    cur = con.cursor()
    cur.execute("truncate table {}".format(tb))
    sql = upload_sql.format(tb)
    cur.execute(sql, raw_data)
    con.commit()
    con.close()

def truncate_n_upload_tb_fr_csv(db,sql_copy, tb, clean_file, sql_insert):
    con = psycopg2.connect(db)
    cur = con.cursor()
    cur.execute("truncate table {}".format(sql_stg))
    sql = sql_copy.format(tb)
    nf = open(clean_file,'r',encoding = 'utf-8')
    cur.copy_expert(sql, nf)
    cur.callproc(sql_insert)
    nf.close()
    con.commit()
    con.close()

