from sqlalchemy import Column, Integer, String, Text, ForeignKey
from db import Base


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(100), unique=True)
    password = Column(String(255))


class Report(Base):
    __tablename__ = "reports"

    id = Column(Integer, primary_key=True, index=True)

    user_id = Column(Integer, ForeignKey("users.id"))

    title = Column(String(200))

    resume_text = Column(Text)

    results = Column(Text)


class CoachReport(Base):
    __tablename__ = "coach_reports"

    id = Column(Integer, primary_key=True, index=True)

    user_id = Column(Integer, ForeignKey("users.id"))

    target_role = Column(String(200))

    job_description = Column(Text)

    resume_text = Column(Text)

    results = Column(Text)
