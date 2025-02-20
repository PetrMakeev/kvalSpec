import sqlite3
from datetime import datetime


class DB_Local:
    def __init__(self, db_name="kvalSpec.db"):
        self.db_name = db_name
        self.DBconn = sqlite3.connect(self.db_name)
        self.cursor = self.DBconn.cursor()
        
        query = """ CREATE TABLE IF NOT EXISTS "prikaz" (
                    "idPrikaz"	integer NOT NULL,
                    "nomPrikaz"	varchar(50) NOT NULL COLLATE NOCASE,
                    "dtPrikaz"	datetime NOT NULL,
                    "pdfLink"	varchar(200),
                    PRIMARY KEY("idPrikaz" AUTOINCREMENT)
                )"""
        cursor = self.execute_query(query)             
        
        query = """ CREATE TABLE IF NOT EXISTS "sotrudnik" (
                    "id"	integer NOT NULL,
                    "family"	varchar(50) NOT NULL COLLATE NOCASE,
                    "name"	varchar(50) NOT NULL COLLATE NOCASE,
                    "lastname"	varchar(50) NOT NULL COLLATE NOCASE,
                    "podr"	varchar(50),
                    "currDol"	varchar(500) NOT NULL COLLATE NOCASE,
                    "zvan"	varchar(50),
                    "kval"	varchar(50),
                    "idPrikaz"	integer NOT NULL,
                    "dtSledKval"	datetime NOT NULL,
                    "status"	INTEGER NOT NULL,
                    "podtvKval"	integer NOT NULL DEFAULT 0,
                    PRIMARY KEY("id" AUTOINCREMENT)
                )"""
        cursor = self.execute_query(query)  
        

        

    def close_connection(self):
        if self.DBconn:
            self.DBconn.close()

    def execute_query(self, query, params=None):
        try:
            with sqlite3.connect(self.db_name) as conn:
                cursor = conn.cursor()
                cursor.execute(query, params or ())
                conn.commit()
                return cursor
        except sqlite3.Error as e:
            print(f"Database error: {e}")
            return None



    def read_spec(self):
        query = (
            "SELECT id, "
            "TRIM(family) || ' ' || TRIM(name) || ' ' || TRIM(lastname) AS fio, "
            "podr, currDol, kval, dtSledKval "
            "FROM sotrudnik ORDER BY family, name, lastname"
        )
        cursor = self.execute_query(query)
        if cursor:
            results = cursor.fetchall()
            processed_results = []
            for el in results:
                try:
                    dt_sled_kval = datetime.strptime(el[5], "%Y-%m-%d %H:%M:%S") if el[5] else None
                except ValueError as e:
                    print(f"Error parsing date for id {el[0]}: {e}")
                    dt_sled_kval = None  # Можно установить значение по умолчанию

                processed_results.append([
                    el[0],  # id
                    el[1],  # fio
                    el[2],  # podr
                    el[3],  # currDol
                    el[4],  # kval
                    dt_sled_kval  # Преобразованная дата или None
                ])
            return processed_results
        return []

    def read_spec_to_edit(self, curr_id):
        #             0       1        2        3          4         5        6       7          8          9           10         11           
        query = (
            "SELECT s.id, s.family, s.name, s.lastname, s.podr, s.currDol, s.zvan, s.kval, s.podtvKval, s.idPrikaz, s.dtSledKval, s.status, " 
            "p.nomPrikaz, p.dtPrikaz "
            "FROM sotrudnik AS s "
            "JOIN prikaz AS p ON s.idPrikaz = p.idPrikaz "
            "WHERE s.id = ?"
        )
        cursor = self.execute_query(query, (curr_id,))
        result = cursor.fetchone()
        return list(result) if result else []


    def spec_save(self, spec):
        if spec.get("id") == -1:
            # сохраняем нового специалиста
            query = """
                INSERT INTO sotrudnik (
                    family, name, lastname, podr, currDol, zvan, kval, podtvKval, idPrikaz, dtSledKval, status
                ) VALUES (
                    :family, :name, :lastname, :podr, :dolzh, :zvan, :kval, :podtv, :id_prikaz, :sledKval, :status
                )
            """
        else:
            # сохраняем изменения в сведения о специалисте
             query = """
                UPDATE sotrudnik
                SET 
                    family = :family,
                    name = :name,
                    lastname = :lastname,
                    podr = :podr,
                    currDol = :dolzh,
                    zvan = :zvan,
                    kval = :kval,
                    podtvKval = :podtv,
                    idPrikaz = :id_prikaz,
                    dtSledKval = :sledKval,
                    status = :status
                WHERE id = :id
            """
        
        try:
            # Проверяем дату на корректный формат
            spec['zvan_dt'] = spec['zvan_dt'].toString("yyyy-MM-dd HH:mm:ss") if spec.get('zvan_dt') else None
            spec['sledKval'] = spec['sledKval'].strftime("%Y-%m-%d %H:%M:%S") if spec.get('sledKval') else None
            
            # Выполняем запрос
            self.execute_query(query, spec)
            self.DBconn.commit()
            print("Запись успешно добавлена.")
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении записи: {e}")
        except Exception as ex:
            print(f"Непредвиденная ошибка: {ex}")
        
        
    def spec_del(self, curr_id):
        query = (
            "DELETE FROM sotrudnik    "
            "WHERE id = ?"
        )
        cursor = self.execute_query(query, (curr_id,))
        self.DBconn.commit() 
    
    
        

    def read_prikaz(self):
        query = "SELECT idPrikaz, nomPrikaz, dtPrikaz, pdfLink FROM prikaz order by dtPrikaz"
        cursor = self.execute_query(query)
        if cursor:
            rezult = cursor.fetchall()
            return [
                [
                    el[0],
                    el[1],
                    datetime.strptime(el[2], "%Y-%m-%d %H:%M:%S"),
                    el[3]
                ]
                for el in rezult
            ]
        return []


    def read_prikaz_to_edit(self, curr_id):
        query = (
            "SELECT p.nomPrikaz, p.dtPrikaz, p.pdfLink "
            "FROM prikaz AS p  "
            "WHERE p.idPrikaz = ?"
        )
        cursor = self.execute_query(query, (curr_id,))
        result = cursor.fetchone()
        return list(result) if result else []



    def prikaz_save(self, prikaz):
        if prikaz.get("id") == -1:
            # сохраняем нового специалиста
            query = """
                INSERT INTO prikaz (
                    nomPrikaz, dtPrikaz, pdfLink
                ) VALUES (
                    :prikaz_nom, :prikaz_dt, :pdf_link
                )
            """
        else:
            # сохраняем изменения в сведения о специалисте
             query = """
                UPDATE prikaz
                SET 
                    nomPrikaz = :prikaz_nom,
                    dtPrikaz = :prikaz_dt,
                    pdfLink = :pdf_link
                WHERE idPrikaz = :id
            """
        
        try:
            # устанавливаем  формат даты Python
            prikaz['prikaz_dt'] = datetime(prikaz['prikaz_dt'].year, 
                                           prikaz['prikaz_dt'].month, 
                                           prikaz['prikaz_dt'].day, 
                                           0,0,0)
            
            # Выполняем запрос
            self.execute_query(query, prikaz)
            self.DBconn.commit()
            print("Запись успешно добавлена.")
        except sqlite3.Error as e:
            print(f"Ошибка при добавлении записи: {e}")
        except Exception as ex:
            print(f"Непредвиденная ошибка: {ex}")
            
            

    def prikaz_del(self, curr_id):
        query = (
            "DELETE FROM prikaz AS p  "
            "WHERE p.idPrikaz = ?"
        )
        cursor = self.execute_query(query, (curr_id,)) 
        self.DBconn.commit()       
        

    def check_spec_on_prikaz(self, prikaz_id):
        query = (
            "SELECT id  "
            "FROM sotrudnik  "
            "WHERE idPrikaz = ?"
        )
        cursor = self.execute_query(query, (prikaz_id,))
        result = cursor.fetchone()
        return list(result) if result else []
                   

    # применяем фильтр и получаем данные для грида
    def read_spec_filter(self, filter_spec, filter_podr, filter_kval, filter_zvan, filter_prikaz, filter_tg, filter_end):
        str_where = ""
        if filter_spec:
            str_where = filter_spec
        if filter_podr:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_podr} "
        if filter_kval:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_kval} "           
        if filter_zvan:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_zvan} "   
        if filter_prikaz:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_prikaz} "   
        if filter_tg:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_tg} "    
        if filter_end:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_end} "    
        
        str_where = ("WHERE " + str_where) if not str_where == '' else ''
        
        
        query = (
            "SELECT id, " +
            "TRIM(family) || ' ' || TRIM(name) || ' ' || TRIM(lastname) AS fio, " +
            "podr, currDol, kval, dtSledKval " +
            "FROM sotrudnik " + 
            str_where +
            "ORDER BY family, name, lastname"
        )
        cursor = self.execute_query(query)
        if cursor:
            results = cursor.fetchall()
            processed_results = []
            for el in results:
                try:
                    dt_sled_kval = datetime.strptime(el[5], "%Y-%m-%d %H:%M:%S") if el[5] else None
                except ValueError as e:
                    print(f"Error parsing date for id {el[0]}: {e}")
                    dt_sled_kval = None  # Можно установить значение по умолчанию

                processed_results.append([
                    el[0],  # id
                    el[1],  # fio
                    el[2],  # podr
                    el[3],  # currDol
                    el[4],  # kval
                    dt_sled_kval  # Преобразованная дата или None
                ])
            
            return processed_results
        return []


    # применяем фильтр и получаем данные для грида
    def read_spec_export(self, filter_spec, filter_podr, filter_kval, filter_zvan, filter_prikaz, filter_tg, filter_end):
        str_where = ""
        if filter_spec:
            str_where = filter_spec
        if filter_podr:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_podr} "
        if filter_kval:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_kval} "           
        if filter_zvan:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_zvan} "   
        if filter_prikaz:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_prikaz} "   
        if filter_tg:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_tg} "    
        if filter_end:
            str_where = str_where + f"{' AND ' if not str_where == '' else ''} {filter_end} "    
        
        str_where = ("WHERE " + str_where) if not str_where == '' else ''
        
        
        query = (
            "SELECT  " +
            "TRIM(s.family) || ' ' || TRIM(s.name) || ' ' || TRIM(s.lastname) AS fio, " +
            "s.podr, s.currDol, s.zvan, s.kval, s.dtSledKval, s.podtvKval, p.nomPrikaz, p.dtPrikaz " +
            "FROM sotrudnik AS s " +
            "JOIN prikaz AS p ON s.idPrikaz = p.idPrikaz " +   
            str_where +
            "ORDER BY family, name, lastname" 
        )
        cursor = self.execute_query(query)
        if cursor:
            results = cursor.fetchall()
            processed_results = []
            for i, el in enumerate(results):
                processed_results.append([
                    i+1,     # №
                    el[0],  # fio
                    el[1],  # podr
                    el[2],  # dolhz
                    el[3],  # zvan
                    el[4],  # kval
                    f"от {datetime.strftime(datetime.strptime(el[8], "%Y-%m-%d %H:%M:%S"), "%d.%m.%Y")} № {el[7]}", # dtPrikaz и nomPrikaz
                    datetime.strftime(datetime.strptime(el[5], "%Y-%m-%d %H:%M:%S"), "%d.%m.%Y"),  # dtSledKval
                    "Подтверждено" if el[6] == 1 else "Присвоено" # podtvKval
                    
                    
                ])
            
            return processed_results
        return []