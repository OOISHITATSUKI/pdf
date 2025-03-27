from sqlalchemy import Column, Integer, String, DateTime, Enum
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
    email = Column(String, unique=True, index=True)
    hashed_password = Column(String)
    created_at = Column(DateTime(timezone=True), server_default=func.now())
    updated_at = Column(DateTime(timezone=True), onupdate=func.now())
    plan_type = Column(Enum(PlanType), default=PlanType.FREE)
    stripe_customer_id = Column(String, unique=True, nullable=True)

    def to_dict(self):
        return {
            "id": self.id,
            "email": self.email,
            "created_at": self.created_at,
            "updated_at": self.updated_at,
            "plan_type": self.plan_type.value if self.plan_type else None,
            "stripe_customer_id": self.stripe_customer_id
        } 