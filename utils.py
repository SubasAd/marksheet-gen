from datetime import date, datetime

def datetime_converter(o):
    if isinstance(o, (datetime, date)):
        return o.isoformat()
    raise TypeError(f"{o.__class__.__name__} is not JSON serializable")
