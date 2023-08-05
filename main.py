import sqlite3
import pandas as pd
import re
from tabulate import tabulate
# table creation code for ceramic capacitor

def ceramic_cap():
    """the below code operates by connecting to the created database, creating a new table ceramic_capf
    and storing all the values of the records of ceramic capacitors"""
    df = pd.read_excel('C:/Users/PC/Desktop/visteondata/ceramiccap.xlsx')
    conn = sqlite3.connect('VISTEON.db')
    table_name = 'ceramiccap'
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    cursor = conn.cursor()
    cursor.execute('select * from ceramiccap ')
    records = cursor.fetchall()
    cursor.execute('create table ceramic_capf(partnumber TEXT primary key,description TEXT,value TEXT,tolerance TEXT,voltage TEXT,package TEXT,PS TEXT,PT TEXT,supplier TEXT )')
    cursor.executemany('insert into ceramic_capf values(?,?,?,?,?,?,?,?,?)', records)
    conn.commit()
    cursor.close()
    conn.close()

# ceramiccapacitor updation code

def update_ceramiccap():
    new_df = pd.read_excel('!!directory!!')
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    new_df.to_sql('ceramic_capf', conn, if_exists='append', index=False)
    conn.commit()
    cursor.close()
    conn.close()

# table for aluminium capacitor

def al_capacitor():
    """the below code works similar to the above code where a new table al_capf is created and all the records of
    aluminium capacitor is stored"""
    df = pd.read_excel('C:/Users/PC/Desktop/visteondata/alcap.xlsx')
    conn=sqlite3.connect('VISTEON.db')
    table_name= 'alcap'
    df.to_sql(table_name,conn, if_exists='replace',index=False)
    cursor=conn.cursor()
    cursor.execute('select * from alcap')
    records = cursor.fetchall()
    cursor.execute('create table al_capf(partnumber TEXT primary key,description TEXT,value TEXT,tolerance TEXT,voltage TEXT,type TEXT,PS TEXT)')
    cursor.executemany('insert into al_capf values(?,?,?,?,?,?,?)',records)
    conn.commit()
    cursor.close()
    conn.close()

# aluminium capacitor table updation code

def update_alcap():
    new_df = pd.read_excel('!!directory!!')
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    new_df.to_sql('al_capf', conn, if_exists='append', index=False)
    conn.commit()
    cursor.close()
    conn.close()

# table for LED

def LED():
    """ a new table ledf is created and the value records of all LED's are stored in it"""
    df = pd.read_excel('C:/Users/PC/Desktop/visteondata/LED.xlsx')
    conn=sqlite3.connect('VISTEON.db')
    table_name= 'LED'
    df.to_sql(table_name,conn, if_exists='replace',index=False)
    cursor=conn.cursor()
    cursor.execute('select * from LED')
    records = cursor.fetchall()
    cursor.execute('create table ledf(partnumber TEXT primary key,description TEXT,colour TEXT,package TEXT,PS TEXT)')
    cursor.executemany('insert into ledf values(?,?,?,?,?)',records)
    conn.commit()
    cursor.close()
    conn.close()

# led table updation code

def update_LED():
    new_df = pd.read_excel('!!directory!!')
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    new_df.to_sql('ledf', conn, if_exists='append', index=False)
    conn.commit()
    cursor.close()
    conn.close()

# table for transistor

def trans():
    """ the below code operates by creating a new table trans_tablef and storing all records of transistor data in it"""
    df = pd.read_excel('C:/Users/PC/Desktop/visteondata/transistor.xlsx')
    conn=sqlite3.connect('VISTEON.db')
    table_name= 'transistor'
    df.to_sql(table_name,conn, if_exists='replace',index=False)
    cursor=conn.cursor()
    cursor.execute('select * from transistor')
    records = cursor.fetchall()
    cursor.execute('create table trans_tablef(partnumber TEXT primary key,description TEXT,current TEXT,voltage TEXT,package TEXT,PS TEXT)')
    cursor.executemany('insert into trans_tablef values(?,?,?,?,?,?)',records)
    conn.commit()
    cursor.close()
    conn.close()

# transistor table updation code

def update_trans():
    new_df = pd.read_excel('!!directory!!')
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    new_df.to_sql('trans_tablef', conn, if_exists='append', index=False)
    conn.commit()
    cursor.close()
    conn.close()

# table for diode

def diode():
    """ diode_tablef is the new table created and all records of the diode are stored in it"""
    df = pd.read_excel('C:/Users/PC/Desktop/visteondata/diode.xlsx')
    conn=sqlite3.connect('VISTEON.db')
    table_name= 'diode'
    df.to_sql(table_name,conn, if_exists='replace',index=False)
    cursor=conn.cursor()
    cursor.execute('select * from diode')
    records = cursor.fetchall()
    cursor.execute('create table diode_tablef(partnumber TEXT primary key,description TEXT,voltage TEXT,current TEXT,package TEXT,tolerance TEXT,spec1 TEXT,spec2 TEXT,PS TEXT)')
    cursor.executemany('insert into diode_tablef values(?,?,?,?,?,?,?,?,?)',records)
    conn.commit()
    cursor.close()
    conn.close()

# diode table updation code

def update_diode():
    new_df = pd.read_excel('!!directory!!')
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    new_df.to_sql('diode_tablef', conn, if_exists='append', index=False)
    conn.commit()
    cursor.close()
    conn.close()

# table for res

def res():
    """ the below code creates a table resf and stores all the records of resistor in the table"""
    df = pd.read_excel('C:/Users/PC/Desktop/visteondata/RES.xlsx')
    conn = sqlite3.connect('VISTEON.db')
    table_name = 'res'
    df.to_sql(table_name, conn, if_exists='replace', index=False)
    cursor = conn.cursor()
    cursor.execute('select * from res ')
    records = cursor.fetchall()
    cursor.execute('create table resf(partnumber TEXT primary key,description TEXT,value TEXT,tolerance TEXT,power TEXT,package TEXT,spec1 TEXT,spec2 TEXT )')
    cursor.executemany('insert into resf values(?,?,?,?,?,?,?,?)', records)
    conn.commit()
    cursor.close()
    conn.close()

# resistor table updation code

def update_res():
    new_df = pd.read_excel('!!directory!!')
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    new_df.to_sql('resf', conn, if_exists='append', index=False)
    conn.commit()
    cursor.close()
    conn.close()

# table has been created by calling the above functions
# the tables are all stored in a single database called VISTEON.db

column_ceramiccap=['partnumber','description','value','tolerance','voltage','package','PS','PT','supplier']
column_alcap=['partnumber','description','value','tolerance','voltage','type','PS']
column_LED=['partnumber','description','colour','package','PS']
column_trans=['partnumber','decription','current','voltage','package','PS']
column_diode=['partnumber','description','voltage','current','package','tolerance','spec1','spec2','PS']
column_res=['partnumber','description','value','tolerance','power','package','spec1','spec2']

#search function for ceramic capacitor

def ceramiccap_search(**search_parameters):
    table_name = 'ceramic_capf'
    column_names = column_ceramiccap
    conn = sqlite3.connect("VISTEON.db")
    cursor = conn.cursor()
    query = f"SELECT {', '.join(column_names)} FROM {table_name} WHERE "
    conditions = []

    # Build the SQL query based on search parameters
    for key, value in search_parameters.items():
        conditions.append(f"LOWER({key}) = LOWER('{value}')")
    query += " AND ".join(conditions)
    cursor.execute(query)
    result = cursor.fetchall()

    query_all = f"SELECT {', '.join(column_names)} FROM {table_name}"
    cursor.execute(query_all)
    all_results = cursor.fetchall()
    target_value_str = search_parameters.get("value")
    text_part = extract_text(target_value_str)
    matching_records = textmatch_ceramiccap(text_part)
    target_numeric = convert_to_numeric(target_value_str)
    exact_matches = []
    nearest_matches = []

    for result in result:
        exact_matches.append(result)

    for result in all_results:
        value_str = result[column_names.index("value")]
        value_numeric = convert_to_numeric(value_str)

        if value_numeric and target_numeric:
            difference = abs(value_numeric - target_numeric)
            tolerance = target_numeric * 0.2  # 20% tolerance

            if difference <= tolerance:
                if result in matching_records and result not in exact_matches:
                    nearest_matches.append(result)

    if exact_matches or nearest_matches:
        headers = column_names
        tables = []

        if exact_matches:
            rows_exact = [headers] + exact_matches
            table_exact = tabulate(rows_exact, headers, tablefmt="html")
            tables.append("<h2>Exact Matches:</h2>")
            tables.append(table_exact)

        if nearest_matches:
            rows_nearest = [headers] + nearest_matches
            table_nearest = tabulate(rows_nearest, headers, tablefmt="html")
            tables.append("<h2>Nearest Matches:</h2>")
            tables.append(table_nearest)

        with open("output.html", "w") as file:
            file.write("\n\n".join(tables))
        print("Exported results to output.html.")
    else:
        print("No matches found.")

    cursor.close()
    conn.close()

# search for aluminium capacitor

def alcap_search(**search_parameters):
    table_name='al_capf'
    column_names=column_alcap
    conn = sqlite3.connect("VISTEON.db")
    cursor = conn.cursor()
    query = f"SELECT {', '.join(column_names)} FROM {table_name} WHERE "
    conditions = []
    """match is performed to find the records where the key value pair matches, if incase the match is found the 
    records is appended to the output table"""
    for key, value in search_parameters.items():
        conditions.append(f"LOWER({key}) = LOWER('{value}')")
    query += " AND ".join(conditions)
    cursor.execute(query)
    result = cursor.fetchall()
    query1 = f"SELECT {', '.join(column_names)} FROM {table_name}"
    cursor.execute(query1)
    all_results = cursor.fetchall()
    target_value_str = search_parameters.get("value")
    textpart=extract_text(target_value_str)
    matching_records=textmatch_alcap(textpart)
    target_numeric = convert_to_numeric(target_value_str)
    nearest_matches = []
    exact_matches=[]
    for result in result:
        exact_matches.append(result)
    for result in all_results:
        value_str = result[column_names.index("value")]
        value_numeric = convert_to_numeric(value_str)

        if value_numeric and target_numeric:
            difference = abs(value_numeric - target_numeric)
            tolerance = target_numeric * 0.2  # 20% tolerance

            if difference <= tolerance :
                if result in matching_records and result not in exact_matches:
                    nearest_matches.append(result)

    if exact_matches or nearest_matches:
        headers = column_names
        tables = []

        if exact_matches:
            rows_exact = [headers] + exact_matches
            table_exact = tabulate(rows_exact, headers, tablefmt="html")
            tables.append("<h2>Exact Matches:</h2>")
            tables.append(table_exact)

        if nearest_matches:
            rows_nearest = [headers] + nearest_matches
            table_nearest = tabulate(rows_nearest, headers, tablefmt="html")
            tables.append("<h2>Nearest Matches:</h2>")
            tables.append(table_nearest)

        with open("output.html", "w") as file:
            file.write("\n\n".join(tables))
        print("Exported results to output.html.")
    else:
        print("No matches found.")
    cursor.close()
    conn.close()

# search function for LED

def LED_search(**search_parameters):
    table_name='ledf'
    column_names=column_LED
    conn = sqlite3.connect("VISTEON.db")
    cursor = conn.cursor()
    query = f"SELECT {', '.join(column_names)} FROM {table_name} WHERE "
    conditions = []
    """for every records that satisfies the key value match, it is being appended to the output"""
    for key, value in search_parameters.items():
        conditions.append(f"{key} LIKE '%{value}%'")
    query += " AND ".join(conditions)
    cursor.execute(query)
    result = cursor.fetchall()
    if result:
        headers = column_names
        rows = [headers] + result  # Include headers as the first row
        table = tabulate(rows, tablefmt="html")
        with open("output.html", "w") as file:
            file.write(table)
        print("Exported results to output.html.")
    else:
        print('No matching records found')
    cursor.close()
    conn.close()

# search function for transistor

def transistor_search(**search_parameters):
    table_name = 'trans_tablef'
    column_names = column_trans
    conn = sqlite3.connect("VISTEON.db")
    cursor = conn.cursor()
    query = f"SELECT {', '.join(column_names)} FROM {table_name} WHERE "
    conditions = []
    """match is performed to find the records where the key value pair matches, if incase the match is found the 
    records is appended to the output table"""
    for key, value in search_parameters.items():
        conditions.append(f"({key}) = ('{value}')")
    query += " AND ".join(conditions)
    cursor.execute(query)
    result = cursor.fetchall()
    query1 = f"SELECT {', '.join(column_names)} FROM {table_name}"
    cursor.execute(query1)
    all_results = cursor.fetchall()
    target_value_str = search_parameters.get("current")
    textpart = extract_text(target_value_str)
    matching_records = textmatch_alcap(textpart)
    target_numeric = convert_to_numeric(target_value_str)
    nearest_matches = []
    exact_matches = []
    for result in result:
        exact_matches.append(result)
    for result in all_results:
        value_str = result[column_names.index("current")]
        value_numeric = convert_to_numeric(value_str)

        if value_numeric and target_numeric:
            difference = abs(value_numeric - target_numeric)
            tolerance = target_numeric * 0.2  # 20% tolerance

            if difference <= tolerance:
                if result in matching_records and result not in exact_matches:
                    nearest_matches.append(result)

    if exact_matches or nearest_matches:
        headers = column_names
        tables = []

        if exact_matches:
            rows_exact = [headers] + exact_matches
            table_exact = tabulate(rows_exact, headers, tablefmt="html")
            tables.append("<h2>Exact Matches:</h2>")
            tables.append(table_exact)

        if nearest_matches:
            rows_nearest = [headers] + nearest_matches
            table_nearest = tabulate(rows_nearest, headers, tablefmt="html")
            tables.append("<h2>Nearest Matches:</h2>")
            tables.append(table_nearest)

        with open("output.html", "w") as file:
            file.write("\n\n".join(tables))
        print("Exported results to output.html.")
    else:
        print("No matches found.")
    cursor.close()
    conn.close()


# search for diode

def diode_search(**search_parameters):
    table_name = 'diode_tablef'
    column_names = column_diode
    conn = sqlite3.connect("VISTEON.db")
    cursor = conn.cursor()
    query = f"SELECT {', '.join(column_names)} FROM {table_name} WHERE "
    conditions = []

    # Build the SQL query based on search parameters
    for key, value in search_parameters.items():
        conditions.append(f"LOWER({key}) = LOWER('{value}')")
    query += " AND ".join(conditions)
    cursor.execute(query)
    result = cursor.fetchall()

    query_all = f"SELECT {', '.join(column_names)} FROM {table_name}"
    cursor.execute(query_all)
    all_results = cursor.fetchall()
    target_value_str = search_parameters.get("voltage")
    target_numeric = convert_to_numeric(target_value_str)
    exact_matches = []
    nearest_matches = []

    for result in result:
        exact_matches.append(result)

    for result in all_results:
        value_str = result[column_names.index("voltage")]
        value_numeric = convert_to_numeric(value_str)

        if value_numeric and target_numeric:
            difference = abs(value_numeric - target_numeric)
            tolerance = target_numeric * 0.2  # 20% tolerance

            if difference <= tolerance:
                if result not in exact_matches:
                    nearest_matches.append(result)

    if exact_matches or nearest_matches:
        headers = column_names
        tables = []

        if exact_matches:
            rows_exact = [headers] + exact_matches
            table_exact = tabulate(rows_exact, headers, tablefmt="html")
            tables.append("<h2>Exact Matches:</h2>")
            tables.append(table_exact)

        if nearest_matches:
            rows_nearest = [headers] + nearest_matches
            table_nearest = tabulate(rows_nearest, headers, tablefmt="html")
            tables.append("<h2>Nearest Matches:</h2>")
            tables.append(table_nearest)

        with open("output.html", "w") as file:
            file.write("\n\n".join(tables))
        print("Exported results to output.html.")
    else:
        print("No matches found.")

    cursor.close()
    conn.close()

# search function for resistor

def resistor_search(**search_parameters):
    table_name = 'resf'
    column_names =column_res
    conn = sqlite3.connect("VISTEON.db")
    cursor = conn.cursor()
    query = f"SELECT {', '.join(column_names)} FROM {table_name} WHERE "
    conditions = []

    # Build the SQL query based on search parameters
    for key, value in search_parameters.items():
        if key == "value":
            pattern = f"{value}%"
            pattern_plural = f"{value[:-1]}%s"
            conditions.append(f"(LOWER({key}) LIKE LOWER('{pattern}') OR LOWER({key}) LIKE LOWER('{pattern_plural}'))")
        else:
            conditions.append(f"LOWER({key}) = LOWER('{value}')")
    query += " AND ".join(conditions)
    cursor.execute(query)
    result = cursor.fetchall()

    query_all = f"SELECT {', '.join(column_names)} FROM {table_name}"
    cursor.execute(query_all)
    all_results = cursor.fetchall()
    target_value_str = search_parameters.get("value")
    text_part = extract_text(target_value_str)
    matching_records = textmatch_resistor(text_part)
    target_numeric = convert_to_numeric(target_value_str)
    exact_matches = []
    nearest_matches = []

    for result in result:
        exact_matches.append(result)

    for result in all_results:
        value_str = result[column_names.index("value")]
        value_numeric = convert_to_numeric(value_str)

        if value_numeric and target_numeric:
            difference = abs(value_numeric - target_numeric)
            tolerance = target_numeric * 0.2  # 20% tolerance

            if difference <= tolerance:
                if result in matching_records and result not in exact_matches:
                    nearest_matches.append(result)

    if exact_matches or nearest_matches:
        headers = column_names
        tables = []

        if exact_matches:
            rows_exact = [headers] + exact_matches
            table_exact = tabulate(rows_exact, headers, tablefmt="html")
            tables.append("<h2>Exact Matches:</h2>")
            tables.append(table_exact)

        if nearest_matches:
            rows_nearest = [headers] + nearest_matches
            table_nearest = tabulate(rows_nearest, headers, tablefmt="html")
            tables.append("<h2>Nearest Matches:</h2>")
            tables.append(table_nearest)

        with open("output.html", "w") as file:
            file.write("\n\n".join(tables))
        print("Exported results to output.html.")
    else:
        print("No matches found.")

    cursor.close()
    conn.close()


def convert_to_numeric(value_str):
    units_mapping = {
        "n": 1e-9,
        "N": 1e-9,
        "u": 1e-6,
        "U": 1e-6,
        "m": 1e-3,
        "k": 1e3,
        "M": 1e6,
        "G": 1e9,
        "F": 1,
        "f": 1,
        "p": 1e-12,
        "P": 1e-12,
        "V": 1,
        "NA": 0,
        "OHMS":1,
        "OHM":1
    }

    matches = re.match(r"(\d+(\.\d+)?)([a-zA-Z]+)", str(value_str))
    if matches:
        numeric_part = float(matches.group(1))
        unit_part = matches.group(3)
        multiplier = units_mapping.get(unit_part, 1)
        numeric_value = numeric_part * multiplier
        return numeric_value

    return None

def textmatch_alcap(text_part):
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    query = "SELECT * FROM al_capf WHERE value LIKE '%" + text_part + "%'"
    cursor.execute(query)
    records = cursor.fetchall()
    cursor.close()
    conn.close()

    return records

def textmatch_ceramiccap(text_part):
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    query = "SELECT * FROM ceramic_capf WHERE value LIKE '%" + text_part + "%'"
    cursor.execute(query)
    records = cursor.fetchall()
    cursor.close()
    conn.close()

    return records

def textmatch_resistor(text_part):
    conn = sqlite3.connect('VISTEON.db')
    cursor = conn.cursor()
    query = "SELECT * FROM resf WHERE value LIKE '%" + text_part + "%' or value LIKE '%"+ text_part+"s%'"
    cursor.execute(query)
    records = cursor.fetchall()
    cursor.close()
    conn.close()

    return records

def extract_text(input_string):
    text_part = re.search(r'[a-zA-Z]+', input_string)
    if text_part:
        return text_part.group()
    else:
        return None
# the above functions perform exact match search operation.
