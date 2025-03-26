from sqlalchemy import create_engine, Column, String, DateTime, Enum, JSON, ForeignKey, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.sql import func
import uuid

Base = declarative_base()

class UploadedFile(Base):
    __tablename__ = 'uploaded_files'
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    user_id = Column(UUID(as_uuid=True), nullable=False)
    filename = Column(String, nullable=False)
    file_path = Column(String, nullable=False)
    uploaded_at = Column(DateTime(timezone=True), server_default=func.now())
    status = Column(
        Enum('pending', 'processed', 'failed', name='processing_status'),
        nullable=False,
        default='pending'
    )

class UserInput(Base):
    __tablename__ = 'user_inputs'
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    file_id = Column(UUID(as_uuid=True), ForeignKey('uploaded_files.id'), nullable=False)
    document_type = Column(String, nullable=False)
    declared_date = Column(String)  # YYYY/MM/DD形式
    additional_note = Column(Text)

class OCRResult(Base):
    __tablename__ = 'ocr_results'
    
    id = Column(UUID(as_uuid=True), primary_key=True, default=uuid.uuid4)
    file_id = Column(UUID(as_uuid=True), ForeignKey('uploaded_files.id'), nullable=False)
    extracted_text = Column(Text, nullable=False)
    layout_json = Column(JSON, nullable=False)
    ocr_engine = Column(String, nullable=False)
    processed_at = Column(DateTime(timezone=True), server_default=func.now()) 