from django.shortcuts import render
from django.http import HttpResponse
from django.views.decorators.cache import cache_page
from django.views.decorators.csrf import csrf_exempt
from .utils import generate_uid, read_xls, find_shifts
from ics import Calendar, Event
import io
import pytz
from datetime import datetime, timedelta

@csrf_exempt
@cache_page(60 * 60 * 24 * 365)  # Cache the page for 1 year
def upload_file(request):
    # Handle the POST request
    if request.method == 'POST' and request.FILES.get('excel_file'):
        uploaded_file = request.FILES['excel_file']
        MAX_FILE_SIZE = 20 * 1024 * 1024  # 5MB

        # Check file size
        if uploaded_file.size > MAX_FILE_SIZE:
            return HttpResponse("File size exceeds the allowed limit of 5MB.", status=400)

        file_content = uploaded_file.read()
        file_stream = io.BytesIO(file_content)

        data_frames = read_xls(file_stream)
        name_to_search = request.POST.get('name_to_search')
        shifts = find_shifts(data_frames, name_to_search)

        # Create the ICS file
        cal = Calendar()
        local_time_zone = pytz.timezone("Australia/Sydney")
        for shift in shifts:
            shift_date = datetime.strptime(shift["Date"], "%d/%m/%Y")
            start_time = datetime.strptime(shift["Start Time"], "%H:%M").time()
            start_dt = datetime.combine(shift_date, start_time)
            localized_start = local_time_zone.localize(start_dt)
            localized_end = localized_start + timedelta(hours=shift["Duration"])

            event_uid = generate_uid(shift_date, name_to_search.lower())

            event = Event()
            event.name = "Work Shift"
            event.begin = localized_start
            event.end = localized_end
            event.uid = event_uid
            cal.events.add(event)

        response = HttpResponse(cal, content_type='text/calendar')
        response['Content-Disposition'] = f'attachment; filename=shifts.ics'
        return response

    # Render the template if it's a GET request
    return render(request, 'upload_form.html')  # This renders the template
