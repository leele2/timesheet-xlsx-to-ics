import io
from django.http import HttpResponse
from django.middleware.csrf import get_token
from datetime import datetime, timedelta
from .utils import read_xls, find_shifts
from ics import Calendar, Event
from django.utils import timezone

def upload_file(request):
    # If the form is submitted
    if request.method == 'POST' and request.FILES['excel_file']:
        uploaded_file = request.FILES['excel_file']
        
        # Read the file content into memory
        file_content = uploaded_file.read()
        file_stream = io.BytesIO(file_content)

        # Process the file
        data_frames = read_xls(file_stream)
        name_to_search = request.POST.get('name_to_search')
        shifts = find_shifts(data_frames, name_to_search)

        # Create the calendar
        cal = Calendar()

        # Pre-calculate the timezone (no need to localize on every shift)
        local_time_zone = timezone.get_current_timezone()

        for shift in shifts:
            shift_date = datetime.strptime(shift["Date"], "%d/%m/%Y")
            start_time = datetime.strptime(shift["Start Time"], "%H:%M").time()
            start_dt = datetime.combine(shift_date, start_time)
            localized_start = timezone.make_aware(start_dt, local_time_zone)

            # Calculate the end time once, adjusting for the shift duration
            end_dt = localized_start + timedelta(hours=shift["Duration"])

            # Create the event
            event = Event()
            event.name = "Work Shift"
            event.begin = localized_start
            event.end = end_dt

            # Add event to calendar
            cal.events.add(event)

        # Generate the .ics file as a response
        response = HttpResponse(cal, content_type='text/calendar')
        response['Content-Disposition'] = f'attachment; filename=shifts.ics'
        return response

    # If no file is uploaded, return an inline HTML response
    csrf_token = get_token(request)

    html = f'''
    <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Upload Excel File</title>
        </head>
        <body>
            <h1>Welcome to the Shift Calendar</h1>
            <p>
                This website allows you to upload an Excel file containing your work shifts,
                and it will generate an ICS calendar file that you can use in your preferred calendar application.
            </p>
            <p>
                Select the Excel file and type the name you would like to find the shifts for, then hit upload.
            </p>
            <p>
                Please note this will only work with Excel .xlsx files and the format must be accepted.
            </p>
            <h2>Upload Your Excel File</h2>
            <form method="POST" enctype="multipart/form-data">
                <input type="hidden" name="csrfmiddlewaretoken" value="{csrf_token}">
                <label for="name_to_search">Name to Search:</label>
                <input type="text" name="name_to_search" id="name_to_search" required>
                <br><br>
                <input type="file" name="excel_file" accept=".xlsx" required>
                <button type="submit">Upload</button>
            </form>
        </body>
    </html>
    '''
    return HttpResponse(html)