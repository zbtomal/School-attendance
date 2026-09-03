import requests
from django.conf import settings


def append_school_name(base_message):
    """
    Uses the short school name when it keeps the message within
    one SMS segment (160 chars). If the message is already longer than
    that (multi-part SMS regardless), the fuller school name is used
    instead since there's no length benefit left to protect.
    """
    short_name = getattr(settings, 'SCHOOL_SHORT_NAME', 'School')
    full_name = getattr(settings, 'SCHOOL_FULL_NAME', short_name)
    short_version = f"{base_message}\n{short_name}"
    if len(short_version) <= 160:
        return short_version
    return f"{base_message}\n{full_name}"


def build_absent_message(student_name, date_str):
    base = (
        f"Dear Parents,\n"
        f"Your child {student_name} was ABSENT on {date_str}. "
        f"Contact the Authority if this is a mistake."
    )
    return append_school_name(base)


def build_teacher_absent_message(teacher_name, date_str):
    base = (
        f"Dear {teacher_name},\n"
        f"You have been marked ABSENT today on {date_str}. "
        f"Please contact the Authority if this is a mistake."
    )
    return append_school_name(base)


def send_sms(number, message):
    token = getattr(settings, 'SMS_TOKEN', None)
    if not token:
        return False, "SMS token is not configured in environment variables."

    url = "https://api.bdbulksms.net/api.php"
    params = {
        "token": token,
        "to": number,
        "message": message,
    }
    try:
        response = requests.get(url, params=params, timeout=10)
        is_success = response.status_code == 200 and "success" in response.text.lower()
        return is_success, response.text
    except requests.RequestException:
        # Avoid exposing token or sensitive parameters in error messages
        return False, "SMS delivery failed due to a network connection error."