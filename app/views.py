from django.shortcuts import render, get_object_or_404
from django.contrib import messages
from django.contrib.auth.decorators import user_passes_test
from django.http import HttpResponse, JsonResponse
from django.views.decorators.csrf import ensure_csrf_cookie
from django.views.decorators.http import require_http_methods
from .models import Empleado, RegistroAsistencia
from .services import ReporteService
from .controllers import (
    procesar_registro_por_empleado,
    procesar_registro_manual,
    crear_contexto_exito,
    buscar_empleado_por_qr_api,
    identificar_por_fingerprint_api,
    vincular_fingerprint_api,
    desvincular_fingerprint_api,
)
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import json


def es_staff(user):
    return user.is_authenticated and user.is_staff


@user_passes_test(es_staff)
def pagina_descarga_excel(request):
    return render(request, 'pagina_descarga_excel.html')


@user_passes_test(es_staff)
def exportar_resumen_excel(request):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen Diario"

    encabezados = [
        "Empleado", "Fecha", "Tiempo de Almuerzo",
        "Horas por Comisión", "Horas por Permiso (Otros)",
        "Horas Trabajadas Totales"
    ]
    ws.append(encabezados)

    datos_diarios = ReporteService.obtener_datos_resumen()

    for (id_empleado, fecha), data in datos_diarios.items():
        empleado = data["empleado"]
        horas = ReporteService.calcular_horas_empleado(data)
        ws.append([
            empleado.nombre_completo,
            fecha.strftime("%Y-%m-%d"),
            horas['almuerzo'],
            horas['comision'],
            horas['permiso'],
            horas['trabajadas']
        ])

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    if ws.max_row > 1:
        tabla = Table(
            displayName="ResumenAsistencia",
            ref=f"A1:F{ws.max_row}"
        )
        style = TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False
        )
        tabla.tableStyleInfo = style
        ws.add_table(tabla)

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = 'attachment; filename=resumen_asistencia.xlsx'
    wb.save(response)
    return response


@user_passes_test(es_staff)
def exportar_asistencia_excel(request):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Asistencia"

    encabezados = ["Empleado", "Tipo de Asistencia", "Fecha", "Hora", "Descripción", "ID Dispositivo"]
    ws.append(encabezados)

    registros = RegistroAsistencia.objects.select_related('empleado', 'tipo') \
        .order_by('-fecha_registro', '-hora_registro')

    for reg in registros:
        ws.append([
            reg.empleado.nombre_completo,
            reg.tipo.nombre_asistencia,
            reg.fecha_registro.strftime('%Y-%m-%d'),
            reg.hora_registro.strftime('%H:%M:%S'),
            reg.descripcion or '',
            reg.fingerprint or '',
        ])

    for col in ws.columns:
        max_length = max(len(str(cell.value)) for cell in col if cell.value)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4F81BD")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.border = thin_border

    if ws.max_row > 1:
        tabla = Table(
            displayName="RegistroAsistencia",
            ref=f"A1:F{ws.max_row}"
        )
        style = TableStyleInfo(
            name="TableStyleMedium9", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False
        )
        tabla.tableStyleInfo = style
        ws.add_table(tabla)

    response = HttpResponse(
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )
    response['Content-Disposition'] = 'attachment; filename=registro_asistencia.xlsx'
    wb.save(response)
    return response


def pagina_principal(request):
    return render(request, 'pagina_principal.html')


def escanear_qr(request):
    return render(request, 'escanear_qr.html')


@ensure_csrf_cookie
def identificar_dispositivo(request):
    empleados = Empleado.objects.order_by('apellidos', 'nombres')
    return render(request, 'identificar.html', {'empleados': empleados})


def registrar_asistencia_qr(request, codigo_qr):
    empleado = Empleado.buscar_por_codigo_qr(codigo_qr)
    if not empleado:
        messages.error(request, 'Código QR no válido o empleado no encontrado.')
        return render(request, 'error_qr.html')

    resultado = procesar_registro_por_empleado(request, empleado)
    if resultado['is_post']:
        if resultado['success']:
            messages.success(request, resultado['message'])
            return render(request, 'asistencia_exitosa.html', crear_contexto_exito(resultado['registro']))
        messages.error(request, resultado['message'])

    return render(request, 'formulario_qr.html', resultado['context'])


def registrar_asistencia_auto(request, empleado_id):
    empleado = get_object_or_404(Empleado, id_empleado=empleado_id)
    resultado = procesar_registro_por_empleado(request, empleado)
    if resultado['is_post']:
        if resultado['success']:
            messages.success(request, resultado['message'])
            return render(request, 'asistencia_exitosa.html', crear_contexto_exito(resultado['registro']))
        messages.error(request, resultado['message'])

    return render(request, 'formulario_qr.html', resultado['context'])


@require_http_methods(["POST", "OPTIONS"])
def api_buscar_empleado_qr(request):
    if request.method == 'OPTIONS':
        return JsonResponse({'success': True})

    try:
        data = json.loads(request.body)
        resultado, status_code = buscar_empleado_por_qr_api(data)
        return JsonResponse(resultado, status=status_code)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Datos JSON inválidos'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': f'Error del servidor: {str(e)}'}, status=500)


@require_http_methods(["POST", "OPTIONS"])
def api_identificar_por_fingerprint(request):
    if request.method == 'OPTIONS':
        return JsonResponse({'success': True})

    try:
        data = json.loads(request.body)
        resultado, status_code = identificar_por_fingerprint_api(data)
        return JsonResponse(resultado, status=status_code)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Datos JSON inválidos'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': f'Error del servidor: {str(e)}'}, status=500)


@require_http_methods(["POST", "OPTIONS"])
def api_vincular_fingerprint(request):
    if request.method == 'OPTIONS':
        return JsonResponse({'success': True})

    try:
        data = json.loads(request.body)
        resultado, status_code = vincular_fingerprint_api(data)
        return JsonResponse(resultado, status=status_code)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Datos JSON inválidos'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': f'Error del servidor: {str(e)}'}, status=500)


@require_http_methods(["POST", "OPTIONS"])
def api_desvincular_fingerprint(request):
    if request.method == 'OPTIONS':
        return JsonResponse({'success': True})

    try:
        data = json.loads(request.body)
        resultado, status_code = desvincular_fingerprint_api(data)
        return JsonResponse(resultado, status=status_code)
    except json.JSONDecodeError:
        return JsonResponse({'success': False, 'error': 'Datos JSON inválidos'}, status=400)
    except Exception as e:
        return JsonResponse({'success': False, 'error': f'Error del servidor: {str(e)}'}, status=500)


def registrar_asistencia(request):
    resultado = procesar_registro_manual(request)
    if resultado['is_post']:
        if resultado['success']:
            messages.success(request, resultado['message'])
            return render(request, 'asistencia_exitosa.html', crear_contexto_exito(resultado['registro']))
        messages.error(request, resultado['message'])

    return render(request, 'formulario.html', resultado['context'])
