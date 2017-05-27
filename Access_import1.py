import pyodbc

#Tests for access driver
#[x for x in pypyodbc.drivers() if x.startswith('Microsoft Access Driver')]

conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=C:\Users\alogan\Desktop\TEMP_GIS\Cape Poge DB\CapePogeDM_Rev.mdb;'
    )
cnxn = pyodbc.connect(conn_str)
crsr = cnxn.cursor()
#Lists all tables
#for table_info in crsr.tables(tableType='TABLE'):
#    print(table_info.table_name)

#Lists query
for row in crsr.columns(table='Intrusive_Results_Master_Query'):
    print(row.column_name)


tid = 'Target_ID'


    
#cnxn.commit()
#cnxn.close()


"""

for row in crsr.tables():
    print(row.table_name)

    







"""
