from django.contrib.auth.models import AnonymousUser, User
from django.test import TestCase, RequestFactory
from django.urls import reverse

from .views import exportar_asistencia_excel


class ReporteExportTests(TestCase):
    def setUp(self):
        self.factory = RequestFactory()
        self.user = User.objects.create_user(username='staff', password='pass')
        self.user.is_staff = True
        self.user.save()

    def test_export_asistencia_requires_staff(self):
        request = self.factory.get(reverse('descargar_excel'))
        request.user = AnonymousUser()

        response = exportar_asistencia_excel(request)

        self.assertEqual(response.status_code, 302)
        self.assertIn('/login/', response['Location'])

    def test_export_asistencia_for_staff_returns_excel(self):
        request = self.factory.get(reverse('descargar_excel'))
        request.user = self.user

        response = exportar_asistencia_excel(request)

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response['Content-Type'],
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
