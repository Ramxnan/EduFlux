from django.shortcuts import render, redirect, get_object_or_404
from . forms import CreateUserForm, LoginForm
from django.contrib.auth import login as auth_login
from django.contrib.auth.models import auth
from .models import File
from django.contrib.auth import authenticate
from django.contrib.auth.decorators import login_required
from django.views.decorators.csrf import csrf_exempt
from django.http import JsonResponse, HttpResponse
from django.core.files.storage import FileSystemStorage
import sys
import os
from django.conf import settings
import json
from django.http import FileResponse, JsonResponse
from django.http import Http404
from django.contrib.auth.models import User

sys.path.append('/path/to/copilot')

from .Part_1.driver import main1
from .Part_3.driver import driver_part3

def homepage(request):
    return render(request, 'nba/index.html')

def register(request):
    form = CreateUserForm()
    if request.method == 'POST':
        form = CreateUserForm(request.POST)
        if form.is_valid():
            user = form.save(commit=False)
            # Extract username part from the email
            display_name = user.username.split('@')[0]
            user.save()
            user_base_directory = os.path.join(settings.MEDIA_ROOT, 'storage', display_name)

            try:
                os.makedirs(user_base_directory, exist_ok=True)
                empty_templates_dir = os.path.join(user_base_directory, 'empty_templates')
                os.makedirs(empty_templates_dir, exist_ok=True)

            except Exception as e:
                # Handle exceptions, such as permission issues
                # messages.error(request, "Error during registration: " + str(e))
                pass

            return redirect('login')  # Redirect to login after successful registration

    context = {'registerform': form}
    return render(request, 'nba/register.html', context=context)


def login(request):
    form = LoginForm()
    if request.method == 'POST':
        form = LoginForm(request, data=request.POST)

        if form.is_valid():
            email = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')

            user = authenticate(request, username=email, password=password)

            if user is not None:
                auth.login(request, user)
                return redirect("dashboard")
            else:
                form.add_error(None, "Invalid email or password")
    context = {'loginform': form}

    return render(request, 'nba/login.html', context=context)



@login_required(login_url = "login")
def dashboard(request):
    display_name = request.user.username.split('@')[0]
    user_directory = os.path.join(settings.MEDIA_ROOT, 'storage', display_name)
    empty_templates_dir = os.path.join(user_directory, 'empty_templates')

    # List all files in the user directory
    empty_templates_files = os.listdir(empty_templates_dir)
    
    return render(request, 'nba/dashboard.html', {'empty_templates_files': empty_templates_files})


def logout(request):
    auth.logout(request)
    return render(request, 'nba/index.html')

@csrf_exempt  # Note: Use this decorator if you want to allow POST requests without CSRF token validation
def submit(request):
    if request.method == 'POST':
        data = {
            "Teacher": str(request.POST.get('teacher')),
            "Academic_year": str(request.POST.get('academicYearStart')) + "-" + str(request.POST.get('academicYearEnd')),
            "Semester": str(request.POST.get('semester')),
            "Branch": str(request.POST.get('branch')),
            "Batch": int(request.POST.get('batch')),
            "Section": str(request.POST.get('section')),
            "Subject_Code": str(request.POST.get('subjectCode')),
            "Subject_Name": str(request.POST.get('subjectName')).replace(" ", "_"),
            "Number_of_Students": int(request.POST.get('numberOfStudents')),
            "Number_of_COs": int(request.POST.get('numberOfCOs')),
        }


        num_components = int(request.POST.get('numberOfComponents'))
        Component_Details = {}
        for i in range(1, num_components+1):
            component_name_key = 'componentName' + str(i)
            component_value_key = 'componentValue' + str(i)
            component_type_key = 'componentType' + str(i)
            
            component_name = request.POST.get(component_name_key)
            component_value = request.POST.get(component_value_key)
            component_type = request.POST.get(component_type_key)

            full_component_name = f"{component_name}_{component_type[0]}"
            Component_Details[full_component_name] = int(component_value)

        print(Component_Details)
        # Directory for empty templates
        display_name = request.user.username.split('@')[0]
        empty_templates_dir = os.path.join(settings.MEDIA_ROOT, 'storage', display_name, 'empty_templates')

        # Generate file name for the Excel file
        excel_file_path = os.path.join(empty_templates_dir)

        # Call main1 function with the necessary data and file path
        generated_file_name = main1(data, Component_Details, excel_file_path)
        file_path = os.path.join(settings.MEDIA_ROOT, 'storage', display_name, 'empty_templates', generated_file_name)
        if os.path.exists(file_path):
            response = FileResponse(open(file_path, 'rb'), as_attachment=True, filename=generated_file_name)
            return response
        return redirect('dashboard')
    # If the request method is not POST, handle accordingly (redirect to a form, etc.)
    return render(request, 'nba/dashboard.html')

from django.http import FileResponse
    


@csrf_exempt   
def upload_multiple_files_po(request):
    if request.method == 'POST':
        uploaded_files = request.FILES.getlist('files')
        PO_uploaded = os.path.join(settings.MEDIA_ROOT, 'storage', request.user.username, 'POCalculations')
        fs = FileSystemStorage(location=PO_uploaded)
        for uploaded_file in uploaded_files:
            fs.save(uploaded_file.name, uploaded_file)
        po_calc_file_name=driver_part3(PO_uploaded, PO_uploaded)
        po_calc_file_path = os.path.join(PO_uploaded, po_calc_file_name)
        if os.path.exists(po_calc_file_path):
            return FileResponse(open(po_calc_file_path, 'rb'), as_attachment=True, filename=po_calc_file_name)
        

        return redirect('dashboard')

        # You can redirect to a success page or do something else after processing files
        #return redirect('success_page')  # Redirect to a success page or dashboard

    else:
        # If it's not a POST request, redirect to dashboard or appropriate page
        return redirect('dashboard')

from django.http import JsonResponse
from django.core.cache import cache
from django.views.decorators.cache import cache_page
from django.utils.decorators import method_decorator
@method_decorator(csrf_exempt, name='dispatch')  # Allow the view to be accessed without CSRF token for simplicity
@method_decorator(cache_page(5), name='dispatch')  # Cache the view response for 5 seconds
@login_required
def fetch_file_lists(request):
    user_directory = os.path.join(settings.MEDIA_ROOT, 'storage', request.user.username)

    # Use cache to store and retrieve the file lists
    cache_key = f'user_{request.user.username}_file_lists'
    file_lists = cache.get(cache_key)

    if not file_lists:
        # Fetch empty template files
        empty_templates_dir = os.path.join(user_directory, 'empty_templates')
        empty_templates_files = os.listdir(empty_templates_dir) if os.path.exists(empty_templates_dir) else []
        # Store file lists in cache for 5 seconds
        file_lists = {
            'empty_templates': empty_templates_files
        }
        cache.set(cache_key, file_lists, 5)

    return JsonResponse(file_lists)



from django.shortcuts import get_object_or_404
from django.http import HttpResponse, Http404
import os
from django.conf import settings

def download_file(request, file_name):
    # Paths to the different folders
    display_name = request.user.username.split('@')[0]
    empty_templates_path = os.path.join(settings.MEDIA_ROOT, 'storage', display_name, 'empty_templates')
    # Check which folder contains the file
    if os.path.isfile(os.path.join(empty_templates_path, file_name)):
        file_path = os.path.join(empty_templates_path, file_name)
    else:
        raise Http404("File not found")

    # If the file exists, serve it
    if os.path.exists(file_path):
        with open(file_path, 'rb') as fh:
            response = HttpResponse(fh.read(), content_type="application/octet-stream")
            response['Content-Disposition'] = 'attachment; filename=' + os.path.basename(file_path)
            return response

    # If the file does not exist
    raise Http404("File not found")

def delete_file(request, file_name):
    # Paths to the different folders
    display_name = request.user.username.split('@')[0]
    empty_templates_path = os.path.join(settings.MEDIA_ROOT, 'storage', display_name, 'empty_templates')

    # Check which folder contains the file
    if os.path.isfile(os.path.join(empty_templates_path, file_name)):
        file_path = os.path.join(empty_templates_path, file_name)
    else:
        raise Http404("File not found")

    # If the file exists, delete it
    if os.path.exists(file_path):
        os.remove(file_path)
        return redirect('dashboard')

    # If the file does not exist
    raise Http404("File not found")




