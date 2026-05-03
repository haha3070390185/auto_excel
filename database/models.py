from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, DateTime, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

Base = declarative_base()

class HistoryRecord(Base):
    __tablename__ = 'history_records'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    operation_type = Column(String(50), nullable=False)
    source_files = Column(Text, nullable=False)
    target_file = Column(String(500), nullable=True)
    options = Column(Text, nullable=True)
    timestamp = Column(DateTime, default=datetime.now)
    status = Column(String(20), default="completed")
    message = Column(Text, nullable=True)

def init_db(db_url='sqlite:///excel_automation.db'):
    engine = create_engine(db_url, echo=False)
    Base.metadata.create_all(engine)
    Session = sessionmaker(bind=engine)
    return Session()
