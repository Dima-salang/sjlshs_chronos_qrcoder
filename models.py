from dataclasses import dataclass, asdict

@dataclass
class StudentList:
    """Data model for the student master list."""
    lrn: str
    last_name: str
    first_name: str
    student_year: str
    section: str
    adviser: str
    gender: str
    
    def toDict(self) -> dict:
        """Converts the Student object to a dictionary."""
        # asdict converts the dataclass to a dict.
        return asdict(self)

    def fromDict(self, data: dict) -> 'StudentList':
        """Converts a dictionary to a Student object."""
        return StudentList(**data)
    

