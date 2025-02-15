from django.http import HttpResponse
from django.views.decorators.cache import cache_page
from django.views.decorators.csrf import csrf_exempt  # CSRF Exempt decorator
import io
import pytz
from datetime import datetime, timedelta
from .utils import read_xls, find_shifts
from ics import Calendar, Event

# Exempt CSRF protection for this view
@csrf_exempt
@cache_page(60 * 60 * 24 * 365)  # Cache the page for 1 year
def upload_file(request):
    # Define a file size limit (e.g., 5MB)
    MAX_FILE_SIZE = 5 * 1024 * 1024  # 5MB

    # If the form is submitted
    if request.method == 'POST' and request.FILES.get('excel_file'):
        uploaded_file = request.FILES['excel_file']

        # Check the file size
        if uploaded_file.size > MAX_FILE_SIZE:
            return HttpResponse("File size exceeds the allowed limit of 5MB.", status=400)

        # Read the file content into memory
        file_content = uploaded_file.read()  # This reads the file content into memory
        file_stream = io.BytesIO(file_content)  # Create an in-memory byte stream

        # Process the file
        data_frames = read_xls(file_stream)  
        name_to_search = request.POST.get('name_to_search')
        shifts = find_shifts(data_frames, name_to_search)

        # Create an ICS calendar file
        cal = Calendar()
        local_time_zone = pytz.timezone("Australia/Sydney")
        for shift in shifts:
            shift_date = datetime.strptime(shift["Date"], "%d/%m/%Y")
            start_time = datetime.strptime(shift["Start Time"], "%H:%M").time()
            
            # Combine date and time
            start_dt = datetime.combine(shift_date, start_time)

            # Localize the start time
            localized_start = local_time_zone.localize(start_dt)

            # Calculate the end time using the duration
            localized_end = localized_start + timedelta(hours=shift["Duration"])

            # Create the event
            event = Event()
            event.name = "Work Shift"
            event.begin = localized_start
            event.end = localized_end

            cal.events.add(event)

        # Generate the .ics file as a response
        response = HttpResponse(cal, content_type='text/calendar')
        response['Content-Disposition'] = f'attachment; filename=shifts.ics'
        return response

    # If no file is uploaded, return an inline HTML response
    html = f'''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload Excel File</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            background-color: #f4f4f4;
            margin: 0;
        }}
        .container {{
            background: white;
            padding: 2rem;
            border-radius: 10px;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
            text-align: center;
            max-width: 400px;
            width: 100%;
        }}
        h1, h2 {{
            color: #333;
        }}
        p {{
            color: #555;
        }}
        input[type="text"], input[type="file"] {{
            width: 100%;
            padding: 10px;
            margin: 10px 0;
            border: 1px solid #ccc;
            border-radius: 5px;
        }}
        button {{
            background-color: #007BFF;
            color: white;
            border: none;
            padding: 10px 15px;
            border-radius: 5px;
            cursor: pointer;
            width: 100%;
            font-size: 16px;
            transition: background 0.3s ease;
        }}
        button:hover {{
            background-color: #0056b3;
        }}
        a {{
            color: #007BFF;
            text-decoration: none;
        }}
        a:hover {{
            text-decoration: underline;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Shift Calendar</h1>
        <p>Upload an Excel file containing your work shifts and generate an ICS calendar file.</p>
        <p><a href="https://github.com/leele2/timesheet-xlsx-to-ics" target="_blank">Learn More</a></p>
        <h2>Upload Your File</h2>
        <form method="POST" enctype="multipart/form-data">
            <label for="name_to_search">Name:</label>
            <input type="text" name="name_to_search" id="name_to_search" required>
            <input type="file" name="excel_file" accept=".xlsx" required>
            <button type="submit">Upload</button>
        </form>
    </div>
</body>
</html>
    '''
    return HttpResponse(html)
