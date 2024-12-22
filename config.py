from dotenv import load_dotenv
import os

load_dotenv()

STUDENT_FILE = 'students.xlsx'
TEACHER_FILE = 'teachers.xlsx'
PROJECTS_FILE = 'projects.xlsx'
PROPOSED_PROJECTS_FILE = 'proposed_projects.xlsx'
TOKEN = os.getenv("TOKEN")
TEACHER_PASSWORD = os.getenv("TEACHER_PASSWORD")
