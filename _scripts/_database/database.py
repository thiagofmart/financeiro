from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import json

user = 'Thiago'
passw = '3hYK0q7CAmg5boOhWTde'
endpoint_write_cluster = 'solardb.cluster-czl3t2lkshwh.us-east-1.rds.amazonaws.com'
endpoint_write_instance = 'solardb-instance-1.czl3t2lkshwh.us-east-1.rds.amazonaws.com'
endpoint_read_cluster = 'solardb.cluster-ro-czl3t2lkshwh.us-east-1.rds.amazonaws.com'
endpoint_read_instance = 'solardb-instance-2.czl3t2lkshwh.us-east-1.rds.amazonaws.com'
port = 3306

engine_write_instance = create_engine(f"mysql://{user}:{passw}@{endpoint_write_instance}:{port}/Solar")
engine_read_instance = create_engine(f"mysql://{user}:{passw}@{endpoint_read_instance}:{port}/Solar") #connect_args={"check_same_thread": False}
Session_write = sessionmaker(bind=engine_write_instance)
Session_read = sessionmaker(bind=engine_read_instance)
Base = declarative_base()

