import datetime
from datetime import date as d
from io import BytesIO

from django.core.exceptions import ValidationError, FieldError
from django.http import HttpResponse
from django.views.generic import View
from django.utils.translation import gettext_lazy as _
from django.template.defaultfilters import date
from xlsxwriter import Workbook


class ModelExportView(View):
    model = None
    list_filter = []

    def get_queryset(self, request):
        queryset = self.model.objects.all()
        for get in request.GET:
            if not self.list_filter or get in self.list_filter:
                try:
                    queryset = queryset.filter(**{get: request.GET[get]})
                except ValidationError:
                    try:
                        value = request.GET[get]
                        data = d(int(value[0:4]), int(value[4:6]), 1)
                        queryset.filter(**{get + '__month': data.month, get + '__year': data.year})
                    except FieldError:
                        pass
                except FieldError:
                    pass
        return queryset


class XlsxExportView:
    fields = []
    worksheet_name = ''
    freeze_panes = True
    autofit = True
    header_format = {'bold': True}

    def get_worksheet(self, workbook, queryset):
        worksheet = workbook.add_worksheet(str(self.worksheet_name))
        bold = workbook.add_format(self.header_format)

        for i, field in enumerate(self.fields):
            if type(field) is tuple:
                worksheet.write(0, i, str(field[1]), bold)
            else:
                worksheet.write(0, i, field, bold)

        row = 1
        for linha in queryset:
            col = 0
            for field in self.fields:
                try:
                    formato = None
                    if type(field) is tuple:
                        name = field[0]
                        if len(field) > 2:
                            formato = workbook.add_format(field[2])
                    else:
                        name = field
                    try:
                        value = getattr(linha, name)

                        if callable(value):
                            value = value()

                        if type(value) is datetime.date:
                            worksheet.write(row, col, date(value, _('d/m/Y')), formato)
                            continue
                        try:
                            worksheet.write(row, col, value, formato)
                        except TypeError:
                            worksheet.write(row, col, str(value), formato)
                    except AttributeError:
                        value = None
                        if callable(name):
                            value = name(linha)
                        elif hasattr(self, name):
                            attr = getattr(self, name)
                            value = attr(linha)
                        if value:
                            worksheet.write(row, col, value, formato)
                finally:
                    col += 1
            row += 1

        if self.freeze_panes:
            worksheet.freeze_panes(1, 0)

        if self.autofit:
            worksheet.autofit()

        return worksheet

    def get_workbook(self, output, queryset):
        workbook = Workbook(output)
        self.get_worksheet(workbook, queryset)
        workbook.close()
        return workbook

    def get_xlsx_response(self, request, queryset):
        output = BytesIO()
        self.get_workbook(output, queryset)
        response = HttpResponse(content_type='application/vnd.ms-excel')
        response['Content-Disposition'] = 'attachment;filename="{}.xlsx"'.format(self.get_filename(request))
        response.write(output.getvalue())
        return response

    def get_filename(self, request):
        return request.resolver_match.view_name


class ModelXlsxView(ModelExportView, XlsxExportView):

    def get(self, request, *args, **kwargs):
        queryset = self.get_queryset(request)
        return self.get_xlsx_response(request, queryset)