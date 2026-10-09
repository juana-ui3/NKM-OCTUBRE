from django.contrib import admin
from .models import Empleado, TipoAsistencia, RegistroAsistencia, DispositivoEmpleado, ActividadProyecto


@admin.register(Empleado)
class EmpleadoAdmin(admin.ModelAdmin):
    list_display = ('id_empleado', 'apellidos', 'nombres', 'codigo_qr')
    search_fields = ('nombres', 'apellidos')
    ordering = ('apellidos', 'nombres')


@admin.register(TipoAsistencia)
class TipoAsistenciaAdmin(admin.ModelAdmin):
    list_display = ('id_tipo', 'nombre_asistencia')


@admin.register(RegistroAsistencia)
class RegistroAsistenciaAdmin(admin.ModelAdmin):
    list_display = ('id_registro', 'empleado', 'tipo', 'fecha_registro', 'hora_registro', 'descripcion')
    list_filter = ('tipo', 'fecha_registro')
    search_fields = ('empleado__nombres', 'empleado__apellidos')
    date_hierarchy = 'fecha_registro'
    ordering = ('-fecha_registro', '-hora_registro')


@admin.register(DispositivoEmpleado)
class DispositivoEmpleadoAdmin(admin.ModelAdmin):
    list_display = ('empleado', 'fingerprint', 'creado_en')
    search_fields = ('empleado__nombres', 'empleado__apellidos', 'fingerprint')


@admin.register(ActividadProyecto)
class ActividadProyectoAdmin(admin.ModelAdmin):
    list_display = ('empleado', 'proyecto', 'actividad', 'fecha', 'hora')
    list_filter = ('fecha',)
