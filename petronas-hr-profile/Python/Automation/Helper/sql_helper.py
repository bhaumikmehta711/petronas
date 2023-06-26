import pandas as pd

def sql_read(engine, query) -> pd:
    try:
        df = pd.read_sql(query, engine)
        return df
    except Exception as e:
        raise ValueError(e)

def sql_execute(engine, query):
    try:
        with engine.begin() as conn:
            conn.execute(query)
    except Exception as e:
        raise ValueError(e)

def sql_insert(engine, df: pd, table_name, schema_name = 'dbo', if_exists = 'append'):
    try:
        row_count = df.to_sql(table_name, schema = schema_name, con = engine, if_exists = if_exists, index = False)
        return row_count
    except Exception as e:
        raise ValueError(e)