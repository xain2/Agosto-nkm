from .models import Empleado, TipoAsistencia, DispositivoEmpleado
from .services import AsistenciaService
from .qr_service import QRService
from .utils import obtener_fecha_hora_actual


def obtener_tipos_evento():
    return TipoAsistencia.objects.all()


def preparar_contexto_formulario(empleado=None):
    return {
        'empleado': empleado,
        'tipos_evento': obtener_tipos_evento()
    }


def preparar_contexto_formulario_manual():
    return {
        'empleados': Empleado.objects.order_by('apellidos', 'nombres'),
        'tipos_evento': obtener_tipos_evento()
    }


def procesar_registro_por_empleado(request, empleado):
    if request.method != 'POST':
        return {'is_post': False, 'context': preparar_contexto_formulario(empleado)}

    tipo_id = request.POST.get('tipo_evento')
    descripcion = request.POST.get('descripcion') or ''
    fingerprint = request.POST.get('fingerprint')

    if not tipo_id:
        return {
            'is_post': True,
            'success': False,
            'message': 'Debe seleccionar un tipo de asistencia.',
            'context': preparar_contexto_formulario(empleado)
        }

    success, message, registro = AsistenciaService.crear_registro_asistencia(
        empleado.id_empleado,
        tipo_id,
        descripcion,
        fingerprint
    )

    return {
        'is_post': True,
        'success': success,
        'message': message,
        'registro': registro,
        'context': preparar_contexto_formulario(empleado)
    }


def procesar_registro_manual(request):
    if request.method != 'POST':
        return {'is_post': False, 'context': preparar_contexto_formulario_manual()}

    empleado_id = request.POST.get('empleado')
    tipo_id = request.POST.get('tipo_evento')
    descripcion = request.POST.get('descripcion') or ''
    fingerprint = request.POST.get('fingerprint')

    if not empleado_id or not tipo_id:
        return {
            'is_post': True,
            'success': False,
            'message': 'Debe seleccionar un empleado y tipo de asistencia.',
            'context': preparar_contexto_formulario_manual()
        }

    success, message, registro = AsistenciaService.crear_registro_asistencia(
        empleado_id,
        tipo_id,
        descripcion,
        fingerprint
    )

    return {
        'is_post': True,
        'success': success,
        'message': message,
        'registro': registro,
        'context': preparar_contexto_formulario_manual()
    }


def crear_contexto_exito(registro):
    fecha, hora = obtener_fecha_hora_actual()
    return {
        'fecha': fecha,
        'hora': hora,
        'empleado': registro.empleado
    }


def buscar_empleado_por_qr_api(request_body):
    codigo_qr = request_body.get('codigo_qr')
    if not codigo_qr:
        return {'success': False, 'error': 'Código QR requerido'}, 400

    resultado = QRService.buscar_empleado_por_qr(codigo_qr)
    status_code = 200 if resultado.get('success') else 404
    return resultado, status_code


def identificar_por_fingerprint_api(request_body):
    fingerprint = request_body.get('fingerprint')
    if not fingerprint:
        return {'success': False, 'error': 'Fingerprint requerido'}, 400

    empleado = DispositivoEmpleado.obtener_empleado_por_fingerprint(fingerprint)
    if empleado:
        return {
            'success': True,
            'empleado': {
                'id': empleado.id_empleado,
                'nombres': empleado.nombres,
                'apellidos': empleado.apellidos,
                'nombre_completo': empleado.nombre_completo,
            }
        }, 200

    return {'success': False, 'error': 'Dispositivo no vinculado a un empleado'}, 404


def vincular_fingerprint_api(request_body):
    empleado_id = request_body.get('empleado_id')
    fingerprint = request_body.get('fingerprint')
    if not empleado_id or not fingerprint:
        return {'success': False, 'error': 'Empleado y fingerprint requeridos'}, 400

    try:
        empleado = Empleado.objects.get(id_empleado=empleado_id)
    except Empleado.DoesNotExist:
        return {'success': False, 'error': 'Empleado no encontrado'}, 404

    DispositivoEmpleado.objects.update_or_create(
        fingerprint=fingerprint,
        defaults={'empleado': empleado}
    )

    return {'success': True, 'empleado_id': empleado.id_empleado}, 201


def desvincular_fingerprint_api(request_body):
    fingerprint = request_body.get('fingerprint')
    if not fingerprint:
        return {'success': False, 'error': 'Fingerprint requerido'}, 400

    borrados, _ = DispositivoEmpleado.objects.filter(fingerprint=fingerprint).delete()
    return {'success': True, 'deleted': borrados}, 200
