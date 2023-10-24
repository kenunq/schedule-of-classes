from django.shortcuts import render
from django.http import JsonResponse

from .services import get_json, get_filename_json
# Create your views here.

def index(request, filename, find):
    result = get_json(filename, find)
    return JsonResponse(result)

def file_name(request):
    result = get_filename_json()
    return JsonResponse(result)