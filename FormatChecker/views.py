from django.http import JsonResponse, FileResponse, HttpResponse
from django.shortcuts import render
from .doc_checker import check_document_rules
from io import BytesIO
import os
import json
def index(request):
    return render(request, 'main_page.html')

def check_document(request):
    if request.method == "POST":
        uploaded_file = request.FILES.get("document")
        if not uploaded_file:
            return JsonResponse({"error": "No file uploaded"}, status=400)
        document_part = request.POST.get("document_part")
        formatting_check = request.POST.get("formatting_check") == "on"
        grammar_check = request.POST.get("grammar_check") == "on"
        exception_words_raw = request.POST.get("exception_words")
        exception_words = []
        if exception_words_raw:
            try:
                exception_words = json.loads(exception_words_raw)
            except Exception as e:
                print("JSON decode error:", e)

        from io import BytesIO
        file_stream = BytesIO(uploaded_file.read())

        from .doc_checker import check_document_rules
        result_text = check_document_rules(file_stream, document_part, formatting_check, grammar_check, exception_words)
        return JsonResponse(result_text)
    return JsonResponse({"error": "No file uploaded"}, status=400)