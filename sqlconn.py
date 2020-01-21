from sqlalchemy import create_engine
from sqlalchemy import Table, Column,VARCHAR,INTEGER,Float,String, MetaData,ForeignKey,Date,DateTime,Text,DECIMAL
from sqlalchemy.sql import exists
import os
class Database():
    engine = create_engine('sqlite:///{}/db/library.db'.format(os.getcwd()),connect_args={'check_same_thread': False})
    meta = MetaData()


    pgso_admin = Table('pgso_admin',meta,
                          Column('userid',INTEGER,primary_key=True),
                          Column('username',VARCHAR(50)),
                          Column('password',VARCHAR(50)))

    pgso_department = Table('pgso_department',meta,
                            Column('id', INTEGER, primary_key=True),
                            Column('type', VARCHAR(50)),
                            Column('name', VARCHAR(50)))

    pgso_procurement = Table('pgso_procurement',meta,
                    Column('id', INTEGER, primary_key=True),
                    Column('department_id',INTEGER, ForeignKey("pgso_department.id"), nullable=False),
                    Column('date_archived',Date),
                    Column('status', String),
                    )

    pgso_procurement_data = Table('pgso_procurement_data',meta,
                         Column('id',INTEGER,primary_key=True),
                         Column('description', String),
                         Column('quantity', INTEGER),
                         Column('unit',String),
                         Column('unit_cost',INTEGER),
                         Column('po_id',INTEGER,ForeignKey('pgso_procurement.id'),nullable=False))

    pgso_price_list = Table('pgso_price_list',meta,
                            Column('id',INTEGER,primary_key=True),
                            Column('item',String),
                            Column('price',String))


    meta.create_all(engine)

    conn = engine.connect()

    s = pgso_admin.select()
    s_value = conn.execute(s)
    z = 0
    for val in s_value:
        z += 1

    if z == 0:
        ins = pgso_admin.insert().values(username = 'admin',
                                            password = 'admin')
        result = conn.execute(ins)

