from rest_framework import serializers
from .models import Document  # Assuming Document model exists

class DocumentSerializer(serializers.ModelSerializer):
    class Meta:
        model = Document
        fields = ['file']