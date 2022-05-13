from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import json


engine = create_engine("sqlite:///./_database/financeiro_test.db", connect_args={"check_same_thread": False}) #connect_args={"check_same_thread": False}
Session = sessionmaker(bind=engine)
session = Session()
Base = declarative_base()
