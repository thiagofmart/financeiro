from .database import Base, engine


def _create_database():
    Base.metadata.create_all(engine)
