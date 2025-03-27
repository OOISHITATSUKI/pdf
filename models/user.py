from sqlalchemy import Column, Integer, String, DateTime, Enum as SQLAlchemyEnum
from sqlalchemy.sql import func
from database import Base
import enum

class PlanType(enum.Enum):
    FREE = "free"
    PRO = "pro"
    PREMIUM = "premium"

class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String(50), unique=True, nullable=False)
    email = Column(String(120), unique=True, nullable=False)
    password_hash = Column(String(128), nullable=False)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())
    plan_type = Column(SQLAlchemyEnum(PlanType), default=PlanType.FREE)
    stripe_customer_id = Column(String(255), unique=True, nullable=True)

    def __repr__(self):
        return f'<User {self.username}>'

    def to_dict(self):
        return {
            "id": self.id,
            "username": self.username,
            "email": self.email,
            "created_at": self.created_at.isoformat() if self.created_at else None,
            "updated_at": self.updated_at,
            "plan_type": self.plan_type.value if self.plan_type else None,
            "stripe_customer_id": self.stripe_customer_id
        } 