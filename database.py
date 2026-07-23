import psycopg2


def get_connection():
    conn = psycopg2.connect(
        host="localhost",
        database="cricketdb",
        user="postgres",
        password="YourPassword",
        port="5432"
    )
    
    return conn