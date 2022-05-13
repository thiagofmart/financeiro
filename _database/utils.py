from .database import Base, engine, Session
from . import models

def _create_database():
    Base.metadata.create_all(engine)

def get_db():
    db = Session()
    return db
