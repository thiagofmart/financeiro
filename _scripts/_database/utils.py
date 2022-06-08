from .database import Base, engine_solar, engine_contruar, SessionSolar, SessionContruar
from . import models
from sqlalchemy_utils import database_exists

def _create_database():
    Base.metadata.create_all(endpoint_write_instance)
    Base.metadata.create_all(endpoint_read_instance)


async def get_db_write():
    db = Session_write()
    try:
        yield db
    finally:
        db.close()

async def get_db_read():
    db = Session_read()
    try:
        yield db
    finally:
        db.close()
