"""
URL configuration for control_asistencia project.

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/5.1/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.contrib import admin
from django.urls import path, include
from django.contrib.auth import views as auth_views
from django.http import JsonResponse
from app.views import (
    api_identificar_por_fingerprint,
    api_vincular_fingerprint,
    api_desvincular_fingerprint,
)

def health_check(request):
    return JsonResponse({"status": "ok"})

urlpatterns = [
    path('health/', health_check, name='health_check'),
    path('admin/', admin.site.urls),
    path('login/', auth_views.LoginView.as_view(template_name='login.html'), name='login'),
    # API explícita en raíz para evitar 404 en despliegue (proxy, etc.)
    path('api/identificar-fingerprint/', api_identificar_por_fingerprint, name='api_identificar_por_fingerprint'),
    path('api/vincular-fingerprint/', api_vincular_fingerprint, name='api_vincular_fingerprint'),
    path('api/desvincular-fingerprint/', api_desvincular_fingerprint, name='api_desvincular_fingerprint'),
    path('', include('app.urls')),
]
