import json
from typing import List, Optional, Dict
from datetime import datetime
from sqlalchemy import create_engine, desc
from sqlalchemy.orm import sessionmaker
from database.models import Base, HistoryRecord, init_db


class DatabaseManager:
    def __init__(self, db_url: str = 'sqlite:///excel_automation.db'):
        self.db_url = db_url
        self.engine = create_engine(db_url, echo=False)
        self.Session = sessionmaker(bind=self.engine)
        self._init_database()

    def _init_database(self):
        Base.metadata.create_all(self.engine)

    def add_history_record(
        self,
        operation_type: str,
        source_files: List[str],
        target_file: Optional[str] = None,
        options: Optional[Dict] = None,
        status: str = "completed",
        message: str = ""
    ) -> bool:
        try:
            session = self.Session()
            
            record = HistoryRecord(
                operation_type=operation_type,
                source_files=json.dumps(source_files, ensure_ascii=False),
                target_file=target_file or "",
                options=json.dumps(options or {}, ensure_ascii=False),
                timestamp=datetime.now(),
                status=status,
                message=message
            )
            
            session.add(record)
            session.commit()
            session.close()
            return True
        
        except Exception as e:
            print(f"添加历史记录失败: {str(e)}")
            return False

    def get_all_history(self, limit: int = 100) -> List[HistoryRecord]:
        try:
            session = self.Session()
            records = session.query(HistoryRecord).order_by(
                desc(HistoryRecord.timestamp)
            ).limit(limit).all()
            session.close()
            return records
        except Exception as e:
            print(f"获取历史记录失败: {str(e)}")
            return []

    def get_history_by_type(self, operation_type: str, limit: int = 50) -> List[HistoryRecord]:
        try:
            session = self.Session()
            records = session.query(HistoryRecord).filter(
                HistoryRecord.operation_type == operation_type
            ).order_by(desc(HistoryRecord.timestamp)).limit(limit).all()
            session.close()
            return records
        except Exception as e:
            print(f"获取历史记录失败: {str(e)}")
            return []

    def get_history_by_date_range(
        self,
        start_date: datetime,
        end_date: datetime
    ) -> List[HistoryRecord]:
        try:
            session = self.Session()
            records = session.query(HistoryRecord).filter(
                HistoryRecord.timestamp >= start_date,
                HistoryRecord.timestamp <= end_date
            ).order_by(desc(HistoryRecord.timestamp)).all()
            session.close()
            return records
        except Exception as e:
            print(f"获取历史记录失败: {str(e)}")
            return []

    def delete_history_record(self, record_id: int) -> bool:
        try:
            session = self.Session()
            record = session.query(HistoryRecord).filter(
                HistoryRecord.id == record_id
            ).first()
            
            if record:
                session.delete(record)
                session.commit()
                session.close()
                return True
            else:
                session.close()
                return False
        
        except Exception as e:
            print(f"删除历史记录失败: {str(e)}")
            return False

    def clear_all_history(self) -> bool:
        try:
            session = self.Session()
            session.query(HistoryRecord).delete()
            session.commit()
            session.close()
            return True
        except Exception as e:
            print(f"清空历史记录失败: {str(e)}")
            return False

    def get_statistics(self) -> Dict:
        try:
            session = self.Session()
            
            total_count = session.query(HistoryRecord).count()
            
            from sqlalchemy import func
            type_stats = session.query(
                HistoryRecord.operation_type,
                func.count(HistoryRecord.id).label('count')
            ).group_by(HistoryRecord.operation_type).all()
            
            session.close()
            
            stats = {
                'total_count': total_count,
                'by_type': {stat.operation_type: stat.count for stat in type_stats}
            }
            
            return stats
        
        except Exception as e:
            print(f"获取统计信息失败: {str(e)}")
            return {'total_count': 0, 'by_type': {}}

    def parse_source_files(self, record: HistoryRecord) -> List[str]:
        try:
            return json.loads(record.source_files)
        except:
            return []

    def parse_options(self, record: HistoryRecord) -> Dict:
        try:
            return json.loads(record.options)
        except:
            return {}
