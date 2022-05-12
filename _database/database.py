from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import json


engine = create_engine("postgresql://postgres:123456@localhost:5433/") #connect_args={"check_same_thread": False}
Session = sessionmaker(bind=engine)
session = Session()
Base = declarative_base()
