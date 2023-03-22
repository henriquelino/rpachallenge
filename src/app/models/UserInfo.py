from dataclasses import dataclass


@dataclass
class UserInfo:
    first_name: str
    last_name: str
    company_name: str
    role_in_company: str
    address: str
    email: str
    phone_number: str
