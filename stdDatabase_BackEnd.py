import sqlite3
import re


def student_data():
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    cur.execute(
        'CREATE TABLE IF NOT EXISTS student (FName text, Nationality text, StudID text, Faculty text,  Address text, Mobile text, Whats text, Disability text)')
    con.commit()
    con.close()


def add_std_rec(name_stud, nationality, stud_id, faculty, address, mobile, whats, disability):
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    cur.execute('INSERT INTO student VALUES (?,?,?,?,?,?,?,?)',
                (name_stud, nationality, stud_id, faculty, address, mobile, whats, disability))
    con.commit()
    con.close()


def view_data():
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    cur.execute('SELECT * FROM student')
    row = cur.fetchall()
    con.close()
    return row


def view(item):
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    cur.execute(
        'SELECT * FROM student WHERE FName=?',
        (item,))
    row = cur.fetchall()
    con.close()
    return row


def delete_rec(item):
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    cur.execute('DELETE FROM student WHERE FName=?', (item,))
    cur.fetchall()
    con.commit()
    con.close()


def search_data(name_stud='', nationality='', stud_id='', faculty='', mobile='', whats='', disability=''):
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    cur.execute(
        'SELECT * FROM student WHERE FName=? or Nationality=?  or StudID=? or Faculty=? or Mobile=? or Whats=? or Disability=?',
        (name_stud, nationality, stud_id, faculty, mobile, whats, disability))
    row = cur.fetchall()
    con.close()
    return row


# ====================

def function_regex(value, pattern):
    c_pattern = re.compile(pattern.lower() + r"/...")
    return c_pattern.search(value) is not None


def search_address(item):
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    con.create_function("REGEXP", 2, function_regex)
    cur.execute('SELECT * FROM student WHERE REGEXP(Address, ?)', (item,))
    row = cur.fetchall()
    con.close()
    return row


# ====================


# def update_data(FName='', Nationality='', StudID='', Faculty='', Address='', Mobile='', Whats='', Disability=''):
#     con = sqlite3.connect('student.sqlite')
#     cur = con.cursor()
#     cur.execute(
#         'UPDATE student SET FName=?, Nationality=?, StudID=?, Faculty=?, Address=?, Mobile=?, Whats=?, Disability=?',
#         (FName, Nationality, StudID, Faculty, Address, Mobile, Whats, Disability))
#     con.commit()
#     con.close()

# ======================

def visit_data():
    con = sqlite3.connect("result.sqlite")
    cur = con.cursor()
    cur.execute(
        "CREATE TABLE IF NOT EXISTS result(id integer primary key, DateVisit text, FName text, StudID text, Address text, ResultVisit text)")
    con.commit()
    con.close()


def get_name(item):
    # con = sqlite3.connect('student.sqlite')
    # cur = con.cursor()
    # names_stud = cur.execute("SELECT FName FROM student")
    # names = names_stud.fetchall()
    # a = map(' '.join, names)
    # con.close()
    # return [''.join(item) for item in a]

    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    con.create_function("REGEXP", 2, function_regex)
    names_stud = cur.execute('SELECT FName FROM student WHERE REGEXP(Address, ?)', (item,))
    names = names_stud.fetchall()
    # a = map(' '.join, names)
    # return [''.join(item) for item in a]

    a = map(' '.join, names)
    ret_names = [''.join(item) for item in a]

    address_stud = cur.execute('SELECT Address FROM student WHERE REGEXP(Address, ?)', (item,))
    address = address_stud.fetchall()
    b = map(' '.join, address)
    ret_address = [''.join(item) for item in b]
    con.close()
    return ret_names, ret_address

def get_items(item):
    con = sqlite3.connect('student.sqlite')
    cur = con.cursor()
    [stud_id], = cur.execute("SELECT StudID FROM student WHERE FName=?", (item,))
    [address], = cur.execute("SELECT Address FROM student WHERE FName=?", (item,))
    con.close()
    return stud_id, address


def add_visit(date_visit, name_stud, stud_id, address, result_visit):
    con = sqlite3.connect('result.sqlite')
    cur = con.cursor()
    cur.execute("INSERT INTO result VALUES(NULL, ?,?,?,?,?)",
                (date_visit, name_stud, stud_id, address, result_visit))
    con.commit()
    con.close()


def view_result_visit(date_visit):
    con = sqlite3.connect('result.sqlite')
    cur = con.cursor()
    cur.execute(
        'SELECT DateVisit, FName, StudID, Address, ResultVisit FROM result WHERE DateVisit=?',
        (date_visit,))
    row = cur.fetchall()
    con.close()
    return row


student_data()

visit_data()
