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

        self.header_font = Font(bold=True, color="FFFFFF")
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.regular_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        self.extra_50_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        self.extra_100_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        self.night_fill = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
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
        # (tu implementación actual sin cambios)
        ws = wb.create_sheet("Resumen Consolidado")
        ws['A1'] = "REPORTE DE ASISTENCIA - RESUMEN CONSOLIDADO"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        headers = [
            'ID Empleado', 'Nombre', 'Apellido', 'Días Trabajados', 'Total Horas',
            'Horas Regulares', 'Horas Extra 50%', 'Horas Extra 100%', 'Horas Nocturnas',
            'Horas Pendientes Iniciales', 'Compensado con 50%', 'Compensado con 100%',
            'Horas Netas Extra 50%', 'Horas Netas Extra 100%', 'Horas Pendientes Finales'
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
            comp = emp['compensations']
            data_row = [
                info.get('employeeInternalId',''),
                info.get('firstName',''),
                info.get('lastName',''),
                totals['total_days_worked'],
                round(totals['total_hours_worked'],2),
                round(totals['total_regular_hours'],2),
                round(totals['total_extra_hours_50'],2),
                round(totals['total_extra_hours_100'],2),
                round(totals['total_night_hours'],2),
                round(totals['total_pending_hours'],2),
                round(comp['compensated_with_50'],2),
                round(comp['compensated_with_100'],2),
                round(comp['net_extra_hours_50'],2),
                round(comp['net_extra_hours_100'],2),
                round(comp['remaining_pending_hours'],2),
            ]
            for col, value in enumerate(data_row, 1):
                c = ws.cell(row=row, column=col, value=value)
                c.border = self.thin_border
                if col == 6: c.fill = self.regular_fill
                elif col in [7,13]: c.fill = self.extra_50_fill
                elif col in [8,14]: c.fill = self.extra_100_fill
                elif col == 9: c.fill = self.night_fill
                elif col in [10,15]: c.fill = self.pending_fill
            row += 1

        for col in range(1, len(headers)+1):
            ws.column_dimensions[get_column_letter(col)].width = 15

        total_row = row + 1
        ws.cell(row=total_row, column=1, value="TOTALES").font = Font(bold=True)
        for col in range(4, len(headers)+1):
            total_formula = f"=SUM({get_column_letter(col)}6:{get_column_letter(col)}{row-1})"
            c = ws.cell(row=total_row, column=col, value=total_formula)
            c.font = Font(bold=True)
            c.border = self.thin_border

    def _create_daily_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        ws = wb.create_sheet("Detalle Diario")
        ws['A1'] = "DETALLE DIARIO DE ASISTENCIA"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)

        headers = [
            'ID Empleado','Nombre','Apellido','Fecha','Día','Horas Trabajadas',
            'Horas Regulares','Horas Extra 50%','Horas Extra 100%','Horas Nocturnas',
            'Horas Pendientes','Es Feriado','Nombre Feriado','Tiene Licencia',
            'Tipo Licencia','Tiene Ausencia','Observaciones',
            'Cálculo (explicación)'  # NUEVO
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
                if d['is_holiday']:
                    observations.append(f"Feriado: {d.get('holiday_name') or 'N/A'}")
                if d['has_time_off']:
                    observations.append(f"Licencia: {d.get('time_off_name') or 'N/A'}")
                if d['has_absence']:
                    observations.append("Ausencia")
                if d['pending_hours'] > 0:
                    observations.append(f"{d['pending_hours']:.1f}h pendientes")
                if d['day_of_week'] in ['Sábado','Domingo']:
                    observations.append("Fin de semana")

                data_row = [
                    info.get('employeeInternalId',''),
                    info.get('firstName',''),
                    info.get('lastName',''),
                    d['date'],
                    d['day_of_week'],
                    round(d['hours_worked'],2),
                    round(d['regular_hours'],2),
                    round(d['extra_hours_50'],2),
                    round(d['extra_hours_100'],2),
                    round(d['night_hours'],2),
                    round(d['pending_hours'],2),
                    'Sí' if d['is_holiday'] else 'No',
                    d.get('holiday_name') or '',
                    'Sí' if d['has_time_off'] else 'No',
                    d.get('time_off_name') or '',
                    'Sí' if d['has_absence'] else 'No',
                    ', '.join(observations) if observations else '',
                    d.get('calc_note','')  # NUEVO
                ]

                for col, value in enumerate(data_row, 1):
                    c = ws.cell(row=row, column=col, value=value)
                    c.border = self.thin_border
                    if col == 7: c.fill = self.regular_fill
                    elif col == 8: c.fill = self.extra_50_fill
                    elif col == 9: c.fill = self.extra_100_fill
                    elif col == 10: c.fill = self.night_fill
                    elif col == 11: c.fill = self.pending_fill

                    # sombreado rápido por fin de semana / feriado
                    if d['day_of_week'] in ['Sábado','Domingo']:
                        c.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                    elif d['is_holiday']:
                        c.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                row += 1

        # Anchos y wrap para columnas largas
        for col in range(1, len(headers)+1):
            width = 12
            if col in [13, 17, 18]:  # Nombre feriado, Observaciones, Explicación
                width = 28 if col == 17 else (55 if col == 18 else 22)
            ws.column_dimensions[get_column_letter(col)].width = width

        # Wrap text en Observaciones y Explicación
        for r in range(6, row):
            ws.cell(row=r, column=17).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=r, column=18).alignment = Alignment(wrap_text=True, vertical='top')

    def _create_statistics_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        # (puedes dejar tu implementación actual)
        ws = wb.create_sheet("Estadísticas")
        ws['A1'] = "ESTADÍSTICAS Y GRÁFICOS"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        # … (resto igual a tu versión)

    def _create_config_sheet(self, wb: Workbook, start_date: str, end_date: str):
        ws = wb.create_sheet("Configuración")
        ws['A1'] = "CONFIGURACIÓN DEL SISTEMA"
        ws['A2'] = "Parámetros utilizados para el reporte"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        # … (resto igual a tu versión)
