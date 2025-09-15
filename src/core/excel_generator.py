"""
Generador de archivos Excel
"""

import os
from datetime import datetime
from typing import Dict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from config.default_config import DEFAULT_CONFIG


class ExcelReportGenerator:
    def __init__(self):
        self.output_dir = os.path.expanduser(DEFAULT_CONFIG['output_directory'])
        self.filename_format = DEFAULT_CONFIG['filename_format']

        # Estilos
        self.header_font = Font(bold=True, color="FFFFFF")
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.regular_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        self.extra_50_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        self.extra_100_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        self.night_fill = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
        self.holiday_fill = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")
        self.pending_fill = PatternFill(start_color="F5C6CB", end_color="F5C6CB", fill_type="solid")

        self.thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )
        self.center_alignment = Alignment(horizontal='center', vertical='center')

    # -------- público --------

    def generate_report(self, processed_data: Dict, start_date: str, end_date: str, output_filename: str = None) -> str:
        wb = Workbook()
        wb.remove(wb.active)

        self._create_summary_sheet(wb, processed_data, start_date, end_date)
        self._create_daily_sheet(wb, processed_data, start_date, end_date)
        self._create_statistics_sheet(wb, processed_data, start_date, end_date)
        self._create_config_sheet(wb, start_date, end_date)

        if not output_filename:
            output_filename = self.filename_format.format(
                start_date=start_date.replace('-', ''), end_date=end_date.replace('-', '')
            )
        os.makedirs(self.output_dir, exist_ok=True)
        filepath = os.path.join(self.output_dir, output_filename)
        wb.save(filepath)
        print(f"✅ Reporte Excel generado: {filepath}")
        return filepath

    # -------- hojas --------

    def _create_summary_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        """
        Resumen Consolidado:
        Solo columnas: ID Empleado, Nombre, Apellido, Total Horas, Horas Regulares,
        Horas Extra 50%, Horas Extra 100%, Horas Nocturnas, Horas Feriado.
        """
        ws = wb.create_sheet("Resumen Consolidado")
        ws['A1'] = "REPORTE DE ASISTENCIA - RESUMEN CONSOLIDADO"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)

        headers = [
            'ID Empleado', 'Nombre', 'Apellido',
            'Total Horas', 'Horas Regulares', 'Horas Extra 50%',
            'Horas Extra 100%', 'Horas Nocturnas', 'Horas Feriado'
        ]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.thin_border
            cell.alignment = self.center_alignment

        row = 6
        for emp in processed_data.values():
            info = emp['employee_info']
            totals = emp['totals']
            data_row = [
                info.get('employeeInternalId', ''),
                info.get('firstName', ''),
                info.get('lastName', ''),
                round(totals.get('total_hours_worked', 0.0), 2),
                round(totals.get('total_regular_hours', 0.0), 2),
                round(totals.get('total_extra_hours_50', 0.0), 2),
                round(totals.get('total_extra_hours_100', 0.0), 2),
                round(totals.get('total_night_hours', 0.0), 2),
                round(totals.get('total_holiday_hours', 0.0), 2),
            ]
            for col, value in enumerate(data_row, 1):
                c = ws.cell(row=row, column=col, value=value)
                c.border = self.thin_border
                if col == 5: c.fill = self.regular_fill
                elif col == 6: c.fill = self.extra_50_fill
                elif col == 7: c.fill = self.extra_100_fill
                elif col == 8: c.fill = self.night_fill
                elif col == 9: c.fill = self.holiday_fill
            row += 1

        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 17

        total_row = row + 1
        ws.cell(row=total_row, column=1, value="TOTALES").font = Font(bold=True)
        # Sumar desde "Total Horas" (col 4) hasta "Horas Feriado" (col 9)
        for col in range(4, 10):
            col_letter = get_column_letter(col)
            total_formula = f"=SUM({col_letter}6:{col_letter}{row - 1})"
            c = ws.cell(row=total_row, column=col, value=total_formula)
            c.font = Font(bold=True)
            c.border = self.thin_border

    def _create_daily_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        """
        Detalle Diario:
        Agrega Inicio Turno, Fin Turno, Es Franco, Horas Feriado.
        """
        ws = wb.create_sheet("Detalle Diario")
        ws['A1'] = "DETALLE DIARIO DE ASISTENCIA"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)

        headers = [
            'ID Empleado', 'Nombre', 'Apellido', 'Fecha', 'Día',
            'Inicio Turno', 'Fin Turno', 'Es Franco',             # NUEVOS
            'Horas Trabajadas', 'Horas Regulares', 'Horas Extra 50%',
            'Horas Extra 100%', 'Horas Nocturnas', 'Horas Feriado',  # NUEVO
            'Horas Pendientes', 'Es Feriado', 'Nombre Feriado',
            'Tiene Licencia', 'Tipo Licencia', 'Tiene Ausencia',
            'Observaciones', 'Cálculo (explicación)'
        ]
        for col, header in enumerate(headers, 1):
            c = ws.cell(row=5, column=col, value=header)
            c.font = self.header_font
            c.fill = self.header_fill
            c.border = self.thin_border
            c.alignment = self.center_alignment

        row = 6
        for emp in processed_data.values():
            info = emp['employee_info']
            for d in emp['daily_data']:
                observations = []
                if d.get('is_holiday'):
                    observations.append(f"Feriado: {d.get('holiday_name') or 'N/A'}")
                if d.get('has_time_off'):
                    observations.append(f"Licencia: {d.get('time_off_name') or 'N/A'}")
                if d.get('has_absence'):
                    observations.append("Ausencia")
                if d.get('pending_hours', 0) > 0:
                    observations.append(f"{d['pending_hours']:.1f}h pendientes")
                if d.get('day_of_week') in ['Sábado', 'Domingo']:
                    observations.append("Fin de semana")

                data_row = [
                    info.get('employeeInternalId', ''),
                    info.get('firstName', ''),
                    info.get('lastName', ''),
                    d.get('date', ''),
                    d.get('day_of_week', ''),
                    d.get('shift_start', ''),                      # NUEVO
                    d.get('shift_end', ''),                        # NUEVO
                    'Sí' if d.get('is_rest_day') else 'No',        # NUEVO
                    round(d.get('hours_worked', 0.0), 2),
                    round(d.get('regular_hours', 0.0), 2),
                    round(d.get('extra_hours_50', 0.0), 2),
                    round(d.get('extra_hours_100', 0.0), 2),
                    round(d.get('night_hours', 0.0), 2),
                    round(d.get('holiday_hours', 0.0), 2),         # NUEVO
                    round(d.get('pending_hours', 0.0), 2),
                    'Sí' if d.get('is_holiday') else 'No',
                    d.get('holiday_name') or '',
                    'Sí' if d.get('has_time_off') else 'No',
                    d.get('time_off_name') or '',
                    'Sí' if d.get('has_absence') else 'No',
                    ', '.join(observations) if observations else '',
                    d.get('calc_note', '')
                ]

                for col, value in enumerate(data_row, 1):
                    c = ws.cell(row=row, column=col, value=value)
                    c.border = self.thin_border
                    # Colores por métricas
                    if col == 10: c.fill = self.regular_fill       # Horas Regulares
                    elif col == 11: c.fill = self.extra_50_fill    # Extra 50
                    elif col == 12: c.fill = self.extra_100_fill   # Extra 100
                    elif col == 13: c.fill = self.night_fill       # Nocturnas
                    elif col == 14: c.fill = self.holiday_fill     # Horas Feriado
                    elif col == 15: c.fill = self.pending_fill     # Pendientes

                    # sombreado rápido por fin de semana / feriado
                    if d.get('day_of_week') in ['Sábado', 'Domingo']:
                        c.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                    elif d.get('is_holiday'):
                        c.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                row += 1

        # Anchos y wrap
        for col in range(1, len(headers) + 1):
            width = 14
            # Nombre Feriado (col 16), Observaciones (col 21), Explicación (col 22)
            if col in [16, 21, 22]:
                width = 28 if col == 21 else (55 if col == 22 else 22)
            ws.column_dimensions[get_column_letter(col)].width = width

        for r in range(6, row):
            ws.cell(row=r, column=21).alignment = Alignment(wrap_text=True, vertical='top')  # Observaciones
            ws.cell(row=r, column=22).alignment = Alignment(wrap_text=True, vertical='top')  # Explicación

    def _create_statistics_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        ws = wb.create_sheet("Estadísticas")
        ws['A1'] = "ESTADÍSTICAS Y GRÁFICOS"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        # Espacio para gráficos/indicadores si luego querés agregarlos.

    def _create_config_sheet(self, wb: Workbook, start_date: str, end_date: str):
        ws = wb.create_sheet("Configuración")
        ws['A1'] = "CONFIGURACIÓN DEL SISTEMA"
        ws['A2'] = "Parámetros utilizados para el reporte"
        ws['A3'] = f"Período: {start_date} al {end_date}"
        ws['A4'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 5):
            ws[f'A{row}'].font = Font(bold=True, size=12)

        # Imprime algunos parámetros clave
        rows = [
            ("output_directory", DEFAULT_CONFIG.get('output_directory', '')),
            ("filename_format", DEFAULT_CONFIG.get('filename_format', '')),
            ("extras_al_50", DEFAULT_CONFIG.get("extras_al_50", 2)),
            ("hora_nocturna_inicio", DEFAULT_CONFIG.get('hora_nocturna_inicio', 21)),
            ("hora_nocturna_fin", DEFAULT_CONFIG.get('hora_nocturna_fin', 6)),
            ("sabado_limite_hora", DEFAULT_CONFIG.get('sabado_limite_hora', 13)),
            ("local_timezone", DEFAULT_CONFIG.get('local_timezone', 'America/Argentina/Buenos_Aires')),
        ]

        start_row = 6
        ws['A5'] = "Clave"
        ws['B5'] = "Valor"
        ws['A5'].font = ws['B5'].font = self.header_font
        ws['A5'].fill = ws['B5'].fill = self.header_fill
        ws['A5'].alignment = ws['B5'].alignment = self.center_alignment
        ws['A5'].border = ws['B5'].border = self.thin_border

        for i, (k, v) in enumerate(rows, start=start_row):
            a = ws.cell(row=i, column=1, value=k)
            b = ws.cell(row=i, column=2, value=v)
            a.border = b.border = self.thin_border

        ws.column_dimensions[get_column_letter(1)].width = 28
        ws.column_dimensions[get_column_letter(2)].width = 55
