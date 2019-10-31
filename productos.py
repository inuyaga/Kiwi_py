import sqlite3
import  os
class Productos:
    def abrir(self):
        conexion=sqlite3.connect(os.path.abspath('database_kiwi.db'))
        return conexion

    def consulta(self, datos):
        try:
            cone=self.abrir()
            cone.row_factory = sqlite3.Row
            cursor=cone.cursor()
            sql="SELECT PRODUCTO, DESC1 FROM M_PROD WHERE PRODUCTO LIKE '%'||?||'%' OR DESC1 LIKE '%'||?||'%' LIMIT 200"
            cursor.execute(sql, datos)
            return cursor.fetchall()
        finally:
            cone.close()
    def get(self, datos):
        try:
            cone=self.abrir()
            cone.row_factory = sqlite3.Row
            cursor=cone.cursor()
            sql = "SELECT * FROM M_PROD WHERE PRODUCTO=?"
            cursor.execute(sql, datos)
            return cursor.fetchone()
        finally:
            cone.close()