from django.urls import path
from . import views

urlpatterns = [
    # Página principal
    path('', views.pagina_principal, name='pagina_principal'),
    
    # Paso previo: Control de Actividades (DESHABILITADO TEMPORALMENTE)
    # path('actividades/<int:empleado_id>/', views.control_actividades, name='control_actividades'),

    # Sistema tradicional (mantenido para compatibilidad)
    path('manual/', views.registrar_asistencia, name='registrar_asistencia'),
    
    # Sistema de QR por empleado (existente)
    path('qr/', views.escanear_qr, name='escanear_qr'),
    path('qr/<str:codigo_qr>/', views.registrar_asistencia_qr, name='registrar_asistencia_qr'),
    path('api/buscar-empleado-qr/', views.api_buscar_empleado_qr, name='api_buscar_empleado_qr'),

    # QR general: auto-identificación por dispositivo (API en control_asistencia/urls.py)
    path('auto/', views.identificar_dispositivo, name='identificar_dispositivo'),
    path('auto/empleado/<int:empleado_id>/', views.registrar_asistencia_auto, name='registrar_asistencia_auto'),

    # Reportes (solo para staff)
    path('login/descarga/', views.pagina_descarga_excel, name='pagina_descarga_excel'),
    path('login/descargar/asistencia/', views.exportar_reporte_mensual_empleados, name='descargar_excel'),
    path('login/descargar/resumen/', views.exportar_resumen_excel, name='resumen_excel'),
]
