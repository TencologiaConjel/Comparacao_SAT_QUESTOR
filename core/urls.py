from django.urls import path
from . import views
from django.contrib.auth import views as auth_views

urlpatterns = [
    path("painel-inicial/", views.painel_inicial, name="painel_inicial"),
    path("sat/importar/", views.sat_importar, name="sat_importar"),

    path("questor/form/", views.comparar_questor_form, name="comparar_questor_form"),         
    path("questor/comparar/", views.comparar_questor, name="comparar_questor"),              
    path("questor/resultado/", views.comparar_questor_resultado, name="comparar_questor_resultado"),  
    path("questor/resultado.csv", views.comparar_questor_csv, name="comparar_questor_csv"),    

    path("", views.login_view, name="login"),
    path("logout/", views.logout_view, name="logout"),
    path('crate-admin/', views.create_admin, name='admin')
]