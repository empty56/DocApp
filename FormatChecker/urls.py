from django.urls import path
from .views import index, check_document

urlpatterns = [
    path("", index, name="index"),
    path("check-format/", check_document, name="check_format"),
]