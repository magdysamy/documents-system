from django.urls import path
from .views import customer_page ,export_pdf ,export_excel

urlpatterns = [
    path('customers/', customer_page, name='customers'),
    path("customers/pdf/", export_pdf),
    path("customers/excel/", export_excel),
]