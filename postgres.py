import pandas as pd
import logging
import psycopg2
from sqlalchemy import create_engine

logging.basicConfig(filename="logs.log", level=logging.DEBUG)

class Employee:
    def __init__(self):
        pass

    def get_connection(self):
        try:
            connection = psycopg2.connect(database="postgres_assignment",
                                          user="user", password="password@123",
                                          host="localhost", port=5432)
            return connection
        except:
            logging.error("Exception occurred while connecting to database")
            raise Exception("Exception occurred.")

    def get_engine(self):
        try:
            engine = create_engine("postgresql+psycopg2://user:postgres@localhost:5432/postgres_assignment")
            logging.info("Engine object created. Returning it")
            return engine
        except:
            logging.error("Exception occurred while creating engine with sqlalchemy.")
            raise Exception("Exception occurred.")

    # Question 1
    def list_of_employees(self):
        connection = self.get_connection()
        cursor = connection.cursor()
        cursor.execute(
            "SELECT t1.empno as EmployeeNumber, t1.ename as EmployeeName, "
            "t2.ename as Manager FROM emp t1, emp t2 WHERE t1.mgr=t2.empno;"
        )
        view = cursor.fetchall()
        df = pd.DataFrame(view, columns=['Number', 'Name', 'Manager'])
        print("\n\nQuestion 1:\n", df.head().to_string(index=False))
        writer = pd.ExcelWriter('ques1.xlsx')
        df.to_excel(writer, sheet_name='ques1', index=False)
        writer.save()
        connection.close()
        logging.info("Connection close for question 1")

    # Question 2
    def total_compensation(self):
        connection = self.get_connection()
        cursor = connection.cursor()
        cursor.execute("UPDATE jobhist SET enddate=CURRENT_DATE WHERE enddate IS NULL;")
        cursor.execute(
            "SELECT emp.ename, "
            "jh.empno, dept.dname, jh.deptno, "
            "ROUND((jh.enddate-jh.startdate)/30*jh.sal,0) "
            "AS total_compensation, ROUND((jh.enddate-jh.startdate)/30,0) as months_spent FROM "
            "jobhist as jh INNER JOIN dept ON jh.deptno=dept.deptno INNER JOIN emp ON jh.empno=emp.empno;")
        view = cursor.fetchall()

        df = pd.DataFrame(view, columns=['Employee_Name', 'Employee_No', 'Dept_Name',
                                         'Dept_Number', 'Total_Compensation', 'Months_Spent'])

        print("\n\nQuestion 2:\n", df.head().to_string(index=False))
        writer = pd.ExcelWriter('ques2.xlsx')
        df.to_excel(writer, sheet_name='ques2', index=False)
        writer.save()
        connection.close()
        logging.info("Connection close for question 2")

    # Question 3
    def upload_to_database(self, excel_file):
        engine = self.get_engine()
        try:
            df = pd.read_excel(excel_file)
            df.to_sql(name='new_table_compensation', con=engine, if_exists='replace', index=False)
            logging.info("Successfully uploaded the excel file to database.")
        except:
            logging.error("Couldn't upload the excel file to database.")

    # Question 4
    def group_by_department(self, excel_file):
        df = pd.read_excel(excel_file)
        df = df.groupby(['Dept_Name', 'Dept_Number']).agg(
            Total_Compensation=pd.NamedAgg(column='Total_Compensation', aggfunc="sum")).reset_index()
        print("\n\nQuestion 4:\n", df.head().to_string(index=False))
        writer = pd.ExcelWriter('ques4.xlsx')
        df.to_excel(writer, sheet_name='ques4', index=False)
        writer.save()


if __name__ == '__main__':

    obj = Employee()
    # Question 1
    obj.list_of_employees()
    # Question 2
    obj.total_compensation()
    # Question 3
    obj.upload_to_database(pd.ExcelFile('ques2.xlsx'))
    # Question 4
    obj.group_by_department(pd.ExcelFile('ques2.xlsx'))

