from .database import Base, engine_solar, engine_contruar, SessionSolar, SessionContruar
from . import models

def _create_database():
    Base.metadata.create_all(engine_solar)
    Base.metadata.create_all(engine_contruar)

def get_db(empresa):
    match empresa:
        case 'SOLAR':
            db = SessionSolar()
        case 'CONTRUAR':
            db = SessionContruar()
        case _:
            db = None
    return db
