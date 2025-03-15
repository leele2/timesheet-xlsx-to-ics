from os import getenv
from django.shortcuts import render
from django.http import HttpResponse, JsonResponse
from django.views.decorators.cache import cache_page
from django.views.decorators.csrf import csrf_exempt
from .utils import generate_uid, read_xls, find_shifts
from ics import Calendar, Event
import io
import requests
import pytz
from datetime import datetime, timedelta

VERCEL_BLOB_URL = "https://blob.vercel-storage.com/"
VERCEL_BLOB_TOKEN = getenv('BLOB_READ_WRITE_TOKEN')

def delete_blob(file_url):
    """Deletes a file from Vercel Blob Storage using its URL."""
    
    try:
        # Prepare the headers for the DELETE request
        headers = {
            "Authorization": f"Bearer {VERCEL_BLOB_TOKEN}",
            "x-api-version": "7",
        }
        # Send DELETE request to Vercel Blob Storage delete endpoint
        delete_response = requests.post(
            f"{VERCEL_BLOB_URL}/delete",
            headers=headers,
            json={"urls": [file_url] if isinstance(file_url, str) else file_url},
        )

        # Check if the deletion was successful
        if delete_response.status_code == 200:
            logger1.info(f"File Deleted")
            return JsonResponse({"message": "File successfully deleted."}, status=200)
        else:
            # If deletion fails, return the response error
            print(delete_response.json())
            return JsonResponse(delete_response.json(), status=delete_response.status_code)

    except Exception as e:
        return JsonResponse({"error": f"Error deleting file: {str(e)}"}, status=500)

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

        # Upload file to Vercel Blob Storage
        headers = {
            "Authorization": f"Bearer {VERCEL_BLOB_TOKEN}",
            "Content-Type": "application/octet-stream",
        }
        upload_response = requests.post(
            f"{VERCEL_BLOB_URL}/put",
            headers=headers,
            files={"file": (uploaded_file.name, uploaded_file.read())},
        )

        if upload_response.status_code != 200:
            return JsonResponse({"error": "Failed to upload file"}, status=500)

        blob_data = upload_response.json()
        file_url = blob_data.get("url")

        try:
            # Download file from Blob Storage
            response = requests.get(file_url)
            response.raise_for_status()
            file_stream = io.BytesIO(response.content)

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

            delete_blob(blob_data["blob"]["key"])

            response = HttpResponse(cal, content_type='text/calendar')
            response['Content-Disposition'] = f'attachment; filename=shifts.ics'
            return response
        
        except Exception as e:
            delete_blob(blob_data["blob"]["key"])  # Ensure the file is deleted even if an error occurs
            return JsonResponse({"error": str(e)}, status=500)

    # Render the template if it's a GET request
    return render(request, 'upload_form.html')  # This renders the template
