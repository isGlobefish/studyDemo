from django.shortcuts import render
from django.http import HttpResponse, HttpRequest
from book.models import BookInfo
import json

# Create your views here.

def index(request):
    books = BookInfo.objects.all()
    # name = '元芳'
    # context = {'name': name}
    # return render(request, 'index.html', locals())
    # return render(request, 'index.html', context)
    context = {
        'books': books
    }
    # print(json.dumps(booklist))
    return render(request, 'index.html', context)
    # return HttpResponse('index')
