from django.shortcuts import render, redirect, get_object_or_404
from . forms import CreateUserForm, LoginForm
from django.contrib.auth import login as auth_login
from django.contrib.auth.models import auth
from .models import File
from django.contrib.auth import authenticate
from django.contrib.auth.decorators import login_required
from django.views.decorators.csrf import csrf_exempt
from django.http import JsonResponse, HttpResponse, FileResponse
from django.core.files.storage import FileSystemStorage
import os
from django.conf import settings
from django.http import Http404
from django.contrib.auth.models import User
from django.contrib import messages
from django.core.cache import cache
from django.views.decorators.cache import cache_page
from django.utils.decorators import method_decorator
from datetime import datetime
import uuid
import zipfile


from .Part_1.driver import main1
from .Part_2.driver import driver_part3

def homepage(request):
    return render(request, 'nba/index.html')

def register(request):
    form = CreateUserForm()
    if request.method == 'POST':
        form = CreateUserForm(request.POST)
        if form.is_valid():
            # Check if the college code is correct
            college_code = form.cleaned_data.get('college_code')
            if college_code == '560035':
                user = form.save(commit=False)
                display_name = user.username.split('@')[0]
                user.save()
            user_base_directory = os.path.join(settings.MEDIA_ROOT, 'storage', display_name)
            try:
                os.makedirs(user_base_directory, exist_ok=True)
                Generated_Templates_dir = os.path.join(user_base_directory, 'Generated_Templates')
                os.makedirs(Generated_Templates_dir, exist_ok=True)
                Branch_Calculation_dir = os.path.join(user_base_directory, 'Branch_Calculation')
                os.makedirs(Branch_Calculation_dir, exist_ok=True)
                Batch_Calculation_dir = os.path.join(user_base_directory, 'Batch_Calculation')
                os.makedirs(Batch_Calculation_dir, exist_ok=True)
            except Exception as e:
                pass
            messages.success(request, "Registration successful, please login.")
            return redirect('login')
        else:
            if User.objects.filter(username=request.POST.get('email')).exists():
                messages.error(request, "Email already registered, please login.")
            else:
                messages.error(request, "Invalid college code, please contact system admin for further instructions.")

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
    Generated_Templates_dir = os.path.join(user_directory, 'Generated_Templates')

    #Template Generation
    Generated_Templates = os.listdir(Generated_Templates_dir)
    file_time_stamp = []
    for file in Generated_Templates:
        file_time_stamp.append(os.path.getmtime(os.path.join(Generated_Templates_dir, file)))
    file_time_stamp = [datetime.fromtimestamp(i).strftime("%H:%M:%S") for i in file_time_stamp]
    Generated_Templates = dict(zip(Generated_Templates, file_time_stamp))
    # End of Template Generation

    #Branch Calculation
    Branch_Calculation_dir = os.path.join(user_directory, 'Branch_Calculation')
    Branch_Calculation = os.listdir(Branch_Calculation_dir)
    file_time_stamp = []
    files=[]
    for folder in Branch_Calculation:
        file_time_stamp.append(os.path.getmtime(os.path.join(Branch_Calculation_dir, folder)))
        files.append(os.listdir(os.path.join(Branch_Calculation_dir, folder)))
    file_time_stamp = [datetime.fromtimestamp(i).strftime("%d-%m-%Y %H:%M") for i in file_time_stamp]

    #i want the dictionary as {folder_name: [[file_name,file_name],file_time_stamp]}
    Branch_Calculation = dict(zip(Branch_Calculation, zip(files,file_time_stamp)))

    # End of Branch Calculation


    return render(request, 'nba/dashboard.html', {'Generated_Templates': Generated_Templates, 
                                                  'Branch_Calculation': Branch_Calculation})

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
            "Subject_Name": str(request.POST.get('subjectName')),
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


@csrf_exempt
def upload_multiple_files_branch(request):
    if request.method == 'POST':
        uploaded_files = request.FILES.getlist('ExcelFiles')
        num_files = len(uploaded_files)
        display_name = request.user.username.split('@')[0]
        user_directory = os.path.join(settings.MEDIA_ROOT, 'storage', display_name)
        branch_directory = os.path.join(user_directory, 'Branch_Calculation')
        if num_files == 0:
            return JsonResponse({'status': 'error', 'message': 'No files were uploaded.'})
        
        unique_id = str(uuid.uuid4())
        unique_folder_name = f"{num_files}Files_BranchCalculation_{unique_id}"
        unique_folder_path = os.path.join(branch_directory, unique_folder_name)
        os.makedirs(unique_folder_path, exist_ok=True)
        fs = FileSystemStorage(location=unique_folder_path)
        
        saved_files = []
        for uploaded_file in uploaded_files:
            filename=fs.save(uploaded_file.name, uploaded_file)
            saved_files.append(os.path.join(unique_folder_path, filename))


        return JsonResponse({'status': 'success', 'files': saved_files, 'folder': unique_folder_name})
    else:
        # If it's not a POST request, handle accordingly
        return JsonResponse({'status': 'error', 'message': 'Invalid request method.'})





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

def download_folder(request, folder_name):
    display_name = request.user.username.split('@')[0]
    branch_calculation_path = os.path.join(settings.MEDIA_ROOT, 'storage', display_name, 'Branch_Calculation', folder_name)

    # Check if the directory exists
    if not os.path.isdir(branch_calculation_path):
        raise Http404("Folder not found")

    # Create a ZIP file in memory
    response = HttpResponse(content_type='application/zip')
    zip_filename = f"{folder_name}.zip"
    response['Content-Disposition'] = f'attachment; filename={zip_filename}'

    with zipfile.ZipFile(response, 'w') as zip_file:
        for foldername, subfolders, filenames in os.walk(branch_calculation_path):
            for filename in filenames:
                # Create complete filepath of file in directory
                file_path = os.path.join(foldername, filename)
                # Add file to zip
                zip_file.write(file_path, os.path.relpath(file_path, branch_calculation_path))

    return response


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




