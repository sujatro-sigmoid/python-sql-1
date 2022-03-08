import pandas as pd
import logging
import psycopg2
from sqlalchemy import create_engine

class Employee:
    def __init__(self):
        pass

    def get_connection(self):
        try:
            connection = psycopg2.connect(database="postgres_assignment", user="user", password="password@123",
                                          host="localhost", port=5432)
            return connection
        except:
            logging.info("Error in Connection")
            raise Exception("Execution unsuccessful. Exception occurred.")
        finally:
            logging.info("No issues found")


    def get_engine(self):
        try:
            engine = create_engine("postgresql+psycopg2://user:postgres@localhost:5432/postgres_assignment")
        except:
            raise Exception("Execution unsuccessful. Exception occurred.")
        return engine

    # Query 1...
    def list_of_employees(self):
        connection = self.get_connection()
        cur = connection.cursor()
        data = cur.execute(
            "SELECT t1.empno as EmployeeNumber, t1.ename as EmployeeName, t2.ename as Manager FROM emp t1, emp t2 WHERE t1.mgr=t2.empno;")
        rows = cur.fetchall()
        li1 = []
        li2 = []
        li3 = []

        for row in rows:
            temp_list = list(row)
            li1.append(temp_list[0])
            li2.append(temp_list[1])
            li3.append(temp_list[2])

        df = pd.DataFrame({'Number': li1, 'Name': li2, 'Manager': li3})
        print(df.head())
        writer = pd.ExcelWriter('ques1.xlsx')
        df.to_excel(writer, sheet_name='ques1', index=False)
        writer.save()
        connection.close()
        logging.info("Connection Close")

    # query 2

    def total_compensation(self):
        connection = self.get_connection()
        cur = connection.cursor()
        # Query 1...
        cur.execute("UPDATE jobhist SET enddate=CURRENT_DATE WHERE enddate IS NULL;")
        data = cur.execute(
            "SELECT emp.ename, "
            "jh.empno, dept.dname, jh.deptno, "
            "ROUND((jh.enddate-jh.startdate)/30*jh.sal,0) "
            "AS total_compensation, ROUND((jh.enddate-jh.startdate)/30,0) as months_spent FROM "
            "jobhist as jh INNER JOIN dept ON jh.deptno=dept.deptno INNER JOIN emp ON jh.empno=emp.empno;")
        rows = cur.fetchall()
        li1 = []
        li2 = []
        li3 = []
        li4 = []
        li5 = []
        li6 = []

        for row in rows:
            temp_list = list(row)
            li1.append(temp_list[0])
            li2.append(temp_list[1])
            li3.append(temp_list[2])
            li4.append(temp_list[3])
            li5.append(temp_list[4])
            li6.append(temp_list[5])
        df = pd.DataFrame(
            {'Employee_Name': li1, 'Employee_No': li2, 'Dept_Name': li3,
             'Dept_Number' : li4, 'Total_Compensation': li5, 'Months_Spent': li6})
        print("\n\n", df.head())
        writer = pd.ExcelWriter('ques2.xlsx')
        df.to_excel(writer, sheet_name='Q2', index=False)
        writer.save()
        connection.close()
        logging.info("Connection Close")

    # Query-3
    def file_to_table(self, data, file):
        engine = self.get_engine()
        try:
            if data == 'Q2':
                df = pd.read_excel(file, 'Q2')
                df.to_sql(name='total_compensation', con=engine, if_exists='append', index=False)
        except:
            logging.info("Execution unsuccessful. Exception occurred.")
        finally:
            logging.info("Execution Successful.")


    def compensation_at_dept_level(self, data, file):
        try:
            if data == 'Q2':
                df = pd.read_excel(file, 'Q2')
                print(df.head())
                return df
        except:
            logging.info("Execution unsuccessful. Exception occurred.")
        finally:
            logging.info("Execution Successful.")


obj = Employee()
# q1
obj.list_of_employees()
# q2
obj.total_compensation()
# q3
with pd.ExcelFile('ques2.xlsx') as xls1:
    for sheet_name in xls1.sheet_names:
        obj.file_to_table(sheet_name, xls1)
# q4
with pd.ExcelFile('ques2.xlsx') as xls2:
    for sheet_name in xls2.sheet_names:
        new_df = obj.compensation_at_dept_level(sheet_name, xls2)

temp1_df = new_df.groupby(['Dept_Name', 'Dept_Number']).agg(
    Total_Compensation=pd.NamedAgg(column='Total_Compensation', aggfunc="sum")).reset_index()

writer = pd.ExcelWriter('ques4.xlsx')
temp1_df.to_excel(writer, sheet_name='Excel_file_Q4', index=False)
writer.save()
