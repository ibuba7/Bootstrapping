from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('curve_bootstrapping/', views.bootstrapping, name='bootstrapping'),
]
