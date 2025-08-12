from PySide6.QtCore import QDate

def is_document_expiring(document, days_threshold):
    """Проверяет, истекает ли документ в течение указанного количества дней"""
    if not document.end_date or document.end_date.isNull():
        return False
    
    current_date = QDate.currentDate()
    days_left = current_date.daysTo(document.end_date)
    
    return 0 <= days_left <= days_threshold

def format_days_left(end_date):
    """Форматирует оставшееся количество дней"""
    if not end_date or end_date.isNull():
        return "Бессрочный"
    
    days_left = QDate.currentDate().daysTo(end_date)
    if days_left < 0:
        return f"Просрочен ({abs(days_left)} дн.)"
    return f"{days_left} дн."