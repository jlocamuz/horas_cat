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
        """
        Resumen Consolidado (reducido):
        ID Empleado | Nombre | Apellido | Total Horas | Horas Regulares | Horas Extra 50% |
        Horas Extra 100% | Horas Nocturnas | Horas Feriado
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
            totals = emp.get('totals', {})

            # Si no viene 'total_holiday_hours' en totals, lo calculo desde daily_data
            total_holiday_hours = totals.get('total_holiday_hours')
            if total_holiday_hours is None:
                total_holiday_hours = round(sum(d.get('holiday_hours', 0.0) for d in emp.get('daily_data', [])), 2)

            data_row = [
                info.get('employeeInternalId',''),
                info.get('firstName',''),
                info.get('lastName',''),
                round(totals.get('total_hours_worked', 0), 2),
                round(totals.get('total_regular_hours', 0), 2),
                round(totals.get('total_extra_hours_50', 0), 2),
                round(totals.get('total_extra_hours_100', 0), 2),
                round(totals.get('total_night_hours', 0), 2),
                round(total_holiday_hours, 2),
            ]

            for col, value in enumerate(data_row, 1):
                c = ws.cell(row=row, column=col, value=value)
                c.border = self.thin_border
                if col == 5: c.fill = self.regular_fill
                elif col == 6: c.fill = self.extra_50_fill
                elif col == 7: c.fill = self.extra_100_fill
                elif col == 8: c.fill = self.night_fill
            row += 1

        for col in range(1, len(headers)+1):
            ws.column_dimensions[get_column_letter(col)].width = 16

        total_row = row + 1
        ws.cell(row=total_row, column=1, value="TOTALES").font = Font(bold=True)
        # Sumar desde columna 4 (Total Horas) hasta última
        for col in range(4, len(headers)+1):
            total_formula = f"=SUM({get_column_letter(col)}6:{get_column_letter(col)}{row-1})"
            c = ws.cell(row=total_row, column=col, value=total_formula)
            c.font = Font(bold=True)
            c.border = self.thin_border

    def _create_daily_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        """
        Detalle Diario + columnas nuevas:
        - Inicio Turno, Fin Turno
        - Día Franco (TRUE/FALSE)
        - Horas Feriado
        """
        ws = wb.create_sheet("Detalle Diario")
        ws['A1'] = "DETALLE DIARIO DE ASISTENCIA"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)

        headers = [
            'ID Empleado','Nombre','Apellido','Fecha','Día',
            'Inicio Turno','Fin Turno','Día Franco',
            'Horas Trabajadas','Horas Regulares','Horas Extra 50%','Horas Extra 100%','Horas Nocturnas',
            'Horas Feriado','Horas Pendientes','Es Feriado','Nombre Feriado',
            'Tiene Licencia','Tipo Licencia','Tiene Ausencia','Observaciones',
            'Cálculo (explicación)'
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
                if d.get('day_of_week') in ['Sábado','Domingo']:
                    observations.append("Fin de semana")

                data_row = [
                    info.get('employeeInternalId',''),
                    info.get('firstName',''),
                    info.get('lastName',''),
                    d.get('date',''),
                    d.get('day_of_week',''),
                    d.get('shift_start',''),                 # NUEVO
                    d.get('shift_end',''),                   # NUEVO
                    'TRUE' if d.get('is_rest_day') else 'FALSE',  # NUEVO
                    round(d.get('hours_worked',0),2),
                    round(d.get('regular_hours',0),2),
                    round(d.get('extra_hours_50',0),2),
                    round(d.get('extra_hours_100',0),2),
                    round(d.get('night_hours',0),2),
                    round(d.get('holiday_hours',0),2),       # NUEVO
                    round(d.get('pending_hours',0),2),
                    'Sí' if d.get('is_holiday') else 'No',
                    d.get('holiday_name') or '',
                    'Sí' if d.get('has_time_off') else 'No',
                    d.get('time_off_name') or '',
                    'Sí' if d.get('has_absence') else 'No',
                    ', '.join(observations) if observations else '',
                    d.get('calc_note','')
                ]

                for col, value in enumerate(data_row, 1):
                    c = ws.cell(row=row, column=col, value=value)
                    c.border = self.thin_border
                    if col == 10: c.fill = self.regular_fill         # Horas Regulares
                    elif col == 11: c.fill = self.extra_50_fill       # Horas Extra 50
                    elif col == 12: c.fill = self.extra_100_fill      # Horas Extra 100
                    elif col == 13: c.fill = self.night_fill          # Nocturnas
                    elif col == 15: c.fill = self.pending_fill        # Pendientes
                    # Sombreado rápido por fin de semana / feriado
                    if d.get('day_of_week') in ['Sábado','Domingo']:
                        c.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                    elif d.get('is_holiday'):
                        c.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                row += 1

        # Anchos y wrap para columnas largas
        for col in range(1, len(headers)+1):
            width = 12
            if col in [17, 21, 22]:  # Nombre feriado, Observaciones, Explicación
                width = 28 if col == 21 else (55 if col == 22 else 22)
            if col in [6, 7]:  # Inicio/Fin turno
                width = 18
            ws.column_dimensions[get_column_letter(col)].width = width

        # Wrap text en Observaciones y Explicación
        for r in range(6, row):
            ws.cell(row=r, column=21).alignment = Alignment(wrap_text=True, vertical='top')
            ws.cell(row=r, column=22).alignment = Alignment(wrap_text=True, vertical='top')

    def _create_statistics_sheet(self, wb: Workbook, processed_data: Dict, start_date: str, end_date: str):
        ws = wb.create_sheet("Estadísticas")
        ws['A1'] = "ESTADÍSTICAS Y GRÁFICOS"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        # (Espacio para gráficos si luego los agregás)

    def _create_config_sheet(self, wb: Workbook, start_date: str, end_date: str):
        ws = wb.create_sheet("Configuración")
        ws['A1'] = "CONFIGURACIÓN DEL SISTEMA"
        ws['A2'] = "Parámetros utilizados para el reporte"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)

        # Parámetros básicos
        params = [
            ("Jornada completa (h)", DEFAULT_CONFIG.get('jornada_completa_horas')),
            ("Inicio horario nocturno", DEFAULT_CONFIG.get('hora_nocturna_inicio')),
            ("Fin horario nocturno", DEFAULT_CONFIG.get('hora_nocturna_fin')),
            ("Límite sábado", DEFAULT_CONFIG.get('sabado_limite_hora')),
            ("Tolerancia (min)", DEFAULT_CONFIG.get('tolerancia_minutos')),
            ("Fragmento (min)", DEFAULT_CONFIG.get('fragmento_minutos')),
            ("Zona horaria", DEFAULT_CONFIG.get('local_timezone', 'America/Argentina/Buenos_Aires')),
            ("Directorio salida", DEFAULT_CONFIG.get('output_directory')),
            ("Formato de archivo", DEFAULT_CONFIG.get('filename_format')),
        ]

        ws.append(())  # fila en blanco
        start_row = ws.max_row + 1
        ws.cell(row=start_row, column=1, value="Clave").font = self.header_font
        ws.cell(row=start_row, column=2, value="Valor").font = self.header_font
        for c in (1, 2):
            ws.cell(row=start_row, column=c).fill = self.header_fill
            ws.cell(row=start_row, column=c).border = self.thin_border
            ws.cell(row=start_row, column=c).alignment = self.center_alignment

        r = start_row + 1
        for k, v in params:
            ws.cell(row=r, column=1, value=k).border = self.thin_border
            ws.cell(row=r, column=2, value=str(v)).border = self.thin_border
            r += 1
