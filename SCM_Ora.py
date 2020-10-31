# -*- coding: utf-8 -*-
"""
Created on Wed Oct  9 10:07:57 2019

@author: u32445
"""

import os
import platform
import pandas as pd
from snipy.db.connect import dbclosing
from snipy.db import funcs_ora as ora
from itertools import count
import numpy as np

#import Functions.Func_Config as funcfg


def include_drivers():
    # include oracle dll files if running on windows
    if platform.system() == "Windows":
        driver = r'\\dom1\shared\Logistic\Shared Docs\BIA\Oracle_Drivers'
        os.environ['PATH'] = driver


def get_data(confile, sql):
    """
    Run an SQL Query via file and return a dataframe

    Parameters
    ----------
    confile: connection or connection string
    sql: filepath to file containing sql

    Returns
    -------
    df: dataframe
    """
    include_drivers()
    # Open and read the file as a single buffer
    fd = open(sql, 'r')
    sqlquery = fd.read()
    fd.close()

    with dbclosing(confile) as conn:
        df = pd.read_sql(sqlquery, con=conn)
    return df


def read(confile, table_name, select_list = None, **kwargs):
    """
    Generates SQL for a SELECT statement matching the kwargs passed.
    confile: file path for connection file to Oracle
    table_name
    select_list: names of columns to SELECT from table_name. If None (default), then SELECT *
    kwargs: conditions for WHERE ___ IN ___ AND ___ IN ___. 
    WHERE k IN v (k is keyword, v is value (v also a list)). If you pass multiple kwargs, it appends with ANDs. 
    """
    include_drivers()
    sql = list()
    if select_list:
        liststring = ",".join(select_list)
        sql.append("SELECT %s FROM %s " % (liststring, table_name))
    else:
        sql.append("SELECT * FROM %s " % (table_name))
    if kwargs:
        sql.append("WHERE " + " AND ".join("%s IN (%s)" 
                                            % (k, v) 
                                            for k, v in kwargs.items()))
    sql = "".join(sql)
    # print(sql)
    with dbclosing(confile) as conn:
        df = pd.read_sql(sql, con=conn)
    return df


def _ora_upsert(cn, df, table_name, constraint_cols, update_cols):
    """
    Upsert df into table_name using an Oracle MERGE INTO statement.

    Parameters
    ----------
    cn: connection or connection string
    df: DataFrame
    table_name: str
    contraint_cols: [str]
    update_cols: [str] or None

    Returns
    -------
    None
    """

    include_drivers()

    if (df is None or len(df) == 0):
        return  # nothing to do

    if update_cols is None:
        update_cols = df.columns.difference(constraint_cols.columns).tolist()

    sel_all_cols = (','.join((':%d %s' % (n, c))
                    for n, c in zip(count(1), df.columns)))
    constraints = (' and '.join(('T.%s = M.%s' % (c, c))
                                for c in constraint_cols))
    updates = ','.join(('%s = M.%s' % (c, c)) for c in update_cols)
    print(updates)
    cols = ','.join('%s' % c for c in update_cols)
    mcols = ','.join('M.%s' % c for c in update_cols)

    # Pandas NA values need to be passed to Oracle as None
    # values not NaT or NaN
    df = df.replace({pd.np.nan: None})
    rows = list(df.itertuples(index=False))

    sql = '''
       merge into %(table_name)s T using
          ( select %(sel_all_cols)s from dual ) M
       on ( %(constraints)s )
       when matched then update set %(updates)s
        when not matched then insert (%(cols)s)
        values (%(mcols)s) ''' % locals()
    # print(sql)
    with dbclosing(cn, CUR=1) as (con, cur):
        cur.executemany(sql, rows)
        con.commit()


def main():


    # Set path to ora file based on the location of the code
    db = funcfg.oraFilepath
    # print("db path is ", db)
    ecdw0 = os.path.join(db, 'ECDW0.ora')

#    pklist = [['46'], ['46']]
#    pkdf = pd.DataFrame(pklist, columns=["SANDBOX_ID"])
#
#    inputlist = [['46', 'NEWTEST', '0', '41234567890', '4987654321'],
#                 ['47', 'TEST2', '1', '21234567890', '2987654321']]
#    inputcol = ora.get_colnms_ora('IM_SANDBOX', 'SCMD', ecdw0)
#    inputdf = pd.DataFrame(inputlist, columns=inputcol)
#
#    _ora_upsert(ecdw0, inputdf, 'SCMD.IM_SANDBOX',
#                constraint_cols=pkdf, update_cols=None)
#    return data
    select = ["SANDBOX_ID", "SANDBOX_TEXT"]
    data = read(ecdw0, 'SCMD.IM_SANDBOX', select_list = select, **{"SANDBOX_ID": "'45', '84'","SANDBOX_TEXT": "'testing'"})
    return data

if __name__ == "__main__":

    test = main()
