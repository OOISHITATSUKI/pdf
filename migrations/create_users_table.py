from sqlalchemy import create_engine
from database import Base
from models.user import User

def run_migration():
    from database import SQLALCHEMY_DATABASE_URL
    engine = create_engine(SQLALCHEMY_DATABASE_URL)
    Base.metadata.create_all(bind=engine)

if __name__ == '__main__':
    run_migration() 