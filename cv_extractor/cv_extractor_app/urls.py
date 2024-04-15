from django.urls import path
from .views import *

urlpatterns=[
    path('',extractorgeneric.as_view(),name='extractor_input'),
]