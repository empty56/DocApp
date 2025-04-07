from django.http import JsonResponse
from django.shortcuts import render
from .doc_checker import check_document_rules
from io import BytesIO

def index(request):
    return render(request, 'main_page.html')

def check_document(request):
    if request.method == "POST":
        uploaded_file = request.FILES.get("document")
        document_part = request.POST.get("document_part")
        formatting_check = request.POST.get("formatting_check") == "on"
        grammar_check = request.POST.get("grammar_check") == "on"

        if not uploaded_file:
            return JsonResponse({"error": "No file uploaded"}, status=400)

        file_stream = BytesIO(uploaded_file.read())  # Convert to memory stream

        # Pass the memory stream instead of a file path
        result = check_document_rules(file_stream, document_part, formatting_check, grammar_check)
        return JsonResponse({"message": result})

    return JsonResponse({"error": "No file uploaded"}, status=400)
