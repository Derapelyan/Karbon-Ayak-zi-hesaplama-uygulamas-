from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, declarative_base
# this creates a local file called carbon.db in your project folder
DATABASE_URL = "sqlite:///carbon.db"

engine = create_engine(DATABASE_URL,    echo = False)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()

def get_session():
    return SessionLocal()