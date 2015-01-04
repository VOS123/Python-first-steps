import pymssql
conn = pymssql.connect("jdvos2", "VOID", "VOID", "agresso57")
cursor = conn.cursor()

sSQL = """
       SELECT c.client , s.apar_id , s.apar_name , s.last_update 
         FROM acrclient c
        INNER JOIN asuheader s ON s.client = c.pay_client       AND
                                  s.last_update > '2014-02-01'  AND
                                  c.client      = 'NL'
        ORDER BY c.client , s.apar_name
       """

cursor.execute(sSQL)

for row in cursor:
    print(row[0], row[1] , row[2] , row[3])


