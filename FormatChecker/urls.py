from django.urls import path
from .views import index, check_format

urlpatterns = [
    path("", index, name="index"),
    path("check-format/", check_format, name="check_format"),
]