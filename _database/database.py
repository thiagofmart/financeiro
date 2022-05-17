from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker
import json


engine_solar = create_engine("sqlite:///./_database/solar.db", connect_args={"check_same_thread": False}) #connect_args={"check_same_thread": False}
engine_contruar = create_engine("sqlite:///./_database/contruar.db", connect_args={"check_same_thread": False}) #connect_args={"check_same_thread": False}
SessionSolar = sessionmaker(bind=engine_solar)
SessionContruar = sessionmaker(bind=engine_contruar)
Base = declarative_base()

