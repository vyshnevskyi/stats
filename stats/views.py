from django.shortcuts import render, redirect
from django.http import HttpResponse, HttpResponseRedirect
from django.contrib import messages
import os
from . import excel_parse

def home(request):
    return render(request, 'home.html')

def upload(request):
    if request.method == 'POST':
        start = request.POST['start']
        end = request.POST['end']
        end1 = request.POST.get('end1', 0)
        if 'file2' in request.FILES:
            handle_uploaded_file(request.FILES['file'], 'noc_schedule1.xlsx')
            handle_uploaded_file(request.FILES['file2'], 'noc_schedule2.xlsx')
            excel_parse.get_all_stats(start, end, 2, end1)
        else:
            handle_uploaded_file(request.FILES['file'], 'noc_schedule1.xlsx')
            excel_parse.get_all_stats(start, end, 1, end1)

        link = "http://" + request.META['HTTP_HOST'] + "/stats/upload/noc_schedule_stats.xlsx"
        messages.success(request, 'Here is a link for stats download: \n' + '<font size="+2"><a href='+link+'>link </a></font>')
        return redirect('home')
    messages.error(request, 'Something went wrong, please try again')
    return redirect('home')

def handle_uploaded_file(file, filename):
    if not os.path.exists('stats/upload/'):
        os.mkdir('stats/upload/')
    with open('stats/upload/' + filename, 'wb+') as destination:
        for chunk in file.chunks():
            destination.write(chunk)
