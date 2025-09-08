"""
Generador de archivos Excel
Crea reportes con múltiples hojas y formato profesional
"""

import os
from datetime import datetime
from typing import Dict, List, Optional
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from config.default_config import DEFAULT_CONFIG


class ExcelReportGenerator:
    """Generador de reportes Excel con múltiples hojas"""
    
    def __init__(self):
        self.output_dir = os.path.expanduser(DEFAULT_CONFIG['output_directory'])
        self.filename_format = DEFAULT_CONFIG['filename_format']
        
        # Estilos predefinidos
        self.header_font = Font(bold=True, color="FFFFFF")
        self.header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        self.regular_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
        self.extra_50_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        self.extra_100_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
        self.night_fill = PatternFill(start_color="D1ECF1", end_color="D1ECF1", fill_type="solid")
        self.pending_fill = PatternFill(start_color="F5C6CB", end_color="F5C6CB", fill_type="solid")
        
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        self.center_alignment = Alignment(horizontal='center', vertical='center')
    
    def generate_report(self, processed_data: Dict, start_date: str, end_date: str, 
                       output_filename: str = None) -> str:
        """
        Genera el reporte Excel completo
        Args:
            processed_data: Datos procesados de empleados
            start_date: Fecha de inicio
            end_date: Fecha de fin
            output_filename: Nombre del archivo (opcional)
        Returns:
            Ruta del archivo generado
        """
        try:
            # Crear workbook
            wb = Workbook()
            
            # Eliminar hoja por defecto
            wb.remove(wb.active)
            
            # Crear hojas
            self._create_summary_sheet(wb, processed_data, start_date, end_date)
            self._create_daily_sheet(wb, processed_data, start_date, end_date)
            self._create_statistics_sheet(wb, processed_data, start_date, end_date)
            self._create_config_sheet(wb, start_date, end_date)
            
            # Generar nombre de archivo
            if not output_filename:
                output_filename = self.filename_format.format(
                    start_date=start_date.replace('-', ''),
                    end_date=end_date.replace('-', '')
                )
            
            # Asegurar que el directorio existe
            os.makedirs(self.output_dir, exist_ok=True)
            
            # Guardar archivo
            filepath = os.path.join(self.output_dir, output_filename)
            wb.save(filepath)
            
            print(f"✅ Reporte Excel generado: {filepath}")
            return filepath
            
        except Exception as e:
            print(f"❌ Error generando reporte Excel: {str(e)}")
            raise e
    
    def _create_summary_sheet(self, wb: Workbook, processed_data: Dict, 
                            start_date: str, end_date: str):
        """Crea la hoja de resumen consolidado"""
        ws = wb.create_sheet("Resumen Consolidado")
        
        # Título
        ws['A1'] = f"REPORTE DE ASISTENCIA - RESUMEN CONSOLIDADO"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Aplicar estilo al título
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        
        # Headers
        headers = [
            'ID Empleado', 'Nombre', 'Apellido', 'Días Trabajados', 'Total Horas',
            'Horas Regulares', 'Horas Extra 50%', 'Horas Extra 100%', 'Horas Nocturnas',
            'Horas Pendientes Iniciales', 'Compensado con 50%', 'Compensado con 100%',
            'Horas Netas Extra 50%', 'Horas Netas Extra 100%', 'Horas Pendientes Finales'
        ]
        
        # Escribir headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.thin_border
            cell.alignment = self.center_alignment
        
        # Escribir datos
        row = 6
        for employee_data in processed_data.values():
            employee_info = employee_data['employee_info']
            totals = employee_data['totals']
            compensations = employee_data['compensations']
            
            data_row = [
                employee_info.get('employeeInternalId', ''),
                employee_info.get('firstName', ''),
                employee_info.get('lastName', ''),
                totals['total_days_worked'],
                round(totals['total_hours_worked'], 2),
                round(totals['total_regular_hours'], 2),
                round(totals['total_extra_hours_50'], 2),
                round(totals['total_extra_hours_100'], 2),
                round(totals['total_night_hours'], 2),
                round(totals['total_pending_hours'], 2),
                round(compensations['compensated_with_50'], 2),
                round(compensations['compensated_with_100'], 2),
                round(compensations['net_extra_hours_50'], 2),
                round(compensations['net_extra_hours_100'], 2),
                round(compensations['remaining_pending_hours'], 2)
            ]
            
            for col, value in enumerate(data_row, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = self.thin_border
                
                # Aplicar colores según el tipo de hora
                if col == 6:  # Horas regulares
                    cell.fill = self.regular_fill
                elif col in [7, 13]:  # Horas extra 50%
                    cell.fill = self.extra_50_fill
                elif col in [8, 14]:  # Horas extra 100%
                    cell.fill = self.extra_100_fill
                elif col == 9:  # Horas nocturnas
                    cell.fill = self.night_fill
                elif col in [10, 15]:  # Horas pendientes
                    cell.fill = self.pending_fill
            
            row += 1
        
        # Ajustar ancho de columnas
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 15
        
        # Fila de totales
        total_row = row + 1
        ws.cell(row=total_row, column=1, value="TOTALES").font = Font(bold=True)
        
        # Calcular totales
        for col in range(4, len(headers) + 1):
            if col in [1, 2, 3]:  # Saltar columnas de texto
                continue
            
            total_formula = f"=SUM({get_column_letter(col)}6:{get_column_letter(col)}{row-1})"
            cell = ws.cell(row=total_row, column=col, value=total_formula)
            cell.font = Font(bold=True)
            cell.border = self.thin_border
    
    def _create_daily_sheet(self, wb: Workbook, processed_data: Dict, 
                          start_date: str, end_date: str):
        """Crea la hoja de detalle diario"""
        ws = wb.create_sheet("Detalle Diario")
        
        # Título
        ws['A1'] = f"DETALLE DIARIO DE ASISTENCIA"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Aplicar estilo al título
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        
        # Headers
        headers = [
            'ID Empleado', 'Nombre', 'Apellido', 'Fecha', 'Día', 'Horas Trabajadas',
            'Horas Regulares', 'Horas Extra 50%', 'Horas Extra 100%', 'Horas Nocturnas',
            'Horas Pendientes', 'Es Feriado', 'Nombre Feriado', 'Tiene Licencia',
            'Tipo Licencia', 'Tiene Ausencia', 'Observaciones'
        ]
        
        # Escribir headers
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=5, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.thin_border
            cell.alignment = self.center_alignment
        
        # Escribir datos diarios
        row = 6
        for employee_data in processed_data.values():
            employee_info = employee_data['employee_info']
            
            for daily_record in employee_data['daily_data']:
                # Generar observaciones
                observations = []
                if daily_record['is_holiday']:
                    observations.append(f"Feriado: {daily_record['holiday_name'] or 'N/A'}")
                if daily_record['has_time_off']:
                    observations.append(f"Licencia: {daily_record['time_off_name'] or 'N/A'}")
                if daily_record['has_absence']:
                    observations.append("Ausencia")
                if daily_record['pending_hours'] > 0:
                    observations.append(f"{daily_record['pending_hours']:.1f}h pendientes")
                if daily_record['day_of_week'] in ['Sábado', 'Domingo']:
                    observations.append("Fin de semana")
                
                data_row = [
                    employee_info.get('employeeInternalId', ''),
                    employee_info.get('firstName', ''),
                    employee_info.get('lastName', ''),
                    daily_record['date'],
                    daily_record['day_of_week'],
                    round(daily_record['hours_worked'], 2),
                    round(daily_record['regular_hours'], 2),
                    round(daily_record['extra_hours_50'], 2),
                    round(daily_record['extra_hours_100'], 2),
                    round(daily_record['night_hours'], 2),
                    round(daily_record['pending_hours'], 2),
                    'Sí' if daily_record['is_holiday'] else 'No',
                    daily_record['holiday_name'] or '',
                    'Sí' if daily_record['has_time_off'] else 'No',
                    daily_record['time_off_name'] or '',
                    'Sí' if daily_record['has_absence'] else 'No',
                    ', '.join(observations) if observations else ''
                ]
                
                for col, value in enumerate(data_row, 1):
                    cell = ws.cell(row=row, column=col, value=value)
                    cell.border = self.thin_border
                    
                    # Aplicar colores según el tipo de hora
                    if col == 7:  # Horas regulares
                        cell.fill = self.regular_fill
                    elif col == 8:  # Horas extra 50%
                        cell.fill = self.extra_50_fill
                    elif col == 9:  # Horas extra 100%
                        cell.fill = self.extra_100_fill
                    elif col == 10:  # Horas nocturnas
                        cell.fill = self.night_fill
                    elif col == 11:  # Horas pendientes
                        cell.fill = self.pending_fill
                    
                    # Resaltar fines de semana y feriados
                    if daily_record['day_of_week'] in ['Sábado', 'Domingo']:
                        cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                    elif daily_record['is_holiday']:
                        cell.fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
                
                row += 1
        
        # Ajustar ancho de columnas
        for col in range(1, len(headers) + 1):
            if col in [13, 15, 17]:  # Columnas de observaciones más anchas
                ws.column_dimensions[get_column_letter(col)].width = 20
            else:
                ws.column_dimensions[get_column_letter(col)].width = 12
    
    def _create_statistics_sheet(self, wb: Workbook, processed_data: Dict, 
                               start_date: str, end_date: str):
        """Crea la hoja de estadísticas con gráficos"""
        ws = wb.create_sheet("Estadísticas")
        
        # Título
        ws['A1'] = f"ESTADÍSTICAS Y GRÁFICOS"
        ws['A2'] = f"Período: {start_date} al {end_date}"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Aplicar estilo al título
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        
        # Calcular estadísticas generales
        total_employees = len(processed_data)
        total_hours = sum(emp['totals']['total_hours_worked'] for emp in processed_data.values())
        total_regular = sum(emp['totals']['total_regular_hours'] for emp in processed_data.values())
        total_extra_50 = sum(emp['totals']['total_extra_hours_50'] for emp in processed_data.values())
        total_extra_100 = sum(emp['totals']['total_extra_hours_100'] for emp in processed_data.values())
        total_night = sum(emp['totals']['total_night_hours'] for emp in processed_data.values())
        total_pending = sum(emp['compensations']['remaining_pending_hours'] for emp in processed_data.values())
        
        # Escribir estadísticas generales
        stats_data = [
            ['ESTADÍSTICAS GENERALES', ''],
            ['Total de empleados', total_employees],
            ['Total horas trabajadas', round(total_hours, 2)],
            ['Horas regulares', round(total_regular, 2)],
            ['Horas extra 50%', round(total_extra_50, 2)],
            ['Horas extra 100%', round(total_extra_100, 2)],
            ['Horas nocturnas', round(total_night, 2)],
            ['Horas pendientes finales', round(total_pending, 2)],
            ['Promedio horas por empleado', round(total_hours / total_employees if total_employees > 0 else 0, 2)]
        ]
        
        for row, (label, value) in enumerate(stats_data, 5):
            ws.cell(row=row, column=1, value=label).font = Font(bold=True)
            ws.cell(row=row, column=2, value=value)
        
        # Top 10 empleados por horas trabajadas
        top_employees = sorted(
            processed_data.values(),
            key=lambda x: x['totals']['total_hours_worked'],
            reverse=True
        )[:10]
        
        # Escribir top 10
        ws.cell(row=15, column=1, value="TOP 10 EMPLEADOS POR HORAS TRABAJADAS").font = Font(bold=True, size=12)
        
        headers_top = ['Empleado', 'Total Horas', 'Horas Regulares', 'Horas Extra', 'Horas Pendientes']
        for col, header in enumerate(headers_top, 1):
            cell = ws.cell(row=16, column=col, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.border = self.thin_border
        
        for row, emp_data in enumerate(top_employees, 17):
            emp_info = emp_data['employee_info']
            totals = emp_data['totals']
            compensations = emp_data['compensations']
            
            name = f"{emp_info.get('firstName', '')} {emp_info.get('lastName', '')}"
            total_extra = totals['total_extra_hours_50'] + totals['total_extra_hours_100']
            
            data_row = [
                name,
                round(totals['total_hours_worked'], 2),
                round(totals['total_regular_hours'], 2),
                round(total_extra, 2),
                round(compensations['remaining_pending_hours'], 2)
            ]
            
            for col, value in enumerate(data_row, 1):
                cell = ws.cell(row=row, column=col, value=value)
                cell.border = self.thin_border
        
        # Ajustar ancho de columnas
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 15
    
    def _create_config_sheet(self, wb: Workbook, start_date: str, end_date: str):
        """Crea la hoja de configuración y parámetros"""
        ws = wb.create_sheet("Configuración")
        
        # Título
        ws['A1'] = f"CONFIGURACIÓN DEL SISTEMA"
        ws['A2'] = f"Parámetros utilizados para el reporte"
        ws['A3'] = f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
        
        # Aplicar estilo al título
        for row in range(1, 4):
            ws[f'A{row}'].font = Font(bold=True, size=12)
        
        # Parámetros del reporte
        config_data = [
            ['PARÁMETROS DEL REPORTE', ''],
            ['Fecha de inicio', start_date],
            ['Fecha de fin', end_date],
            ['Jornada completa (horas)', DEFAULT_CONFIG['jornada_completa_horas']],
            ['Tolerancia (minutos)', DEFAULT_CONFIG['tolerancia_minutos']],
            ['Fragmento mínimo (minutos)', DEFAULT_CONFIG['fragmento_minutos']],
            ['Hora nocturna inicio', f"{DEFAULT_CONFIG['hora_nocturna_inicio']}:00"],
            ['Hora nocturna fin', f"{DEFAULT_CONFIG['hora_nocturna_fin']}:00"],
            ['Sábado límite regular', f"{DEFAULT_CONFIG['sabado_limite_hora']}:00"],
            ['Zona horaria', DEFAULT_CONFIG['timezone']],
            ['', ''],
            ['NORMATIVA APLICADA', ''],
            ['Art. 197 LCT', 'Jornada máxima de 8 horas diarias'],
            ['Art. 201 LCT', 'Horas extras limitadas a 2 horas diarias'],
            ['Art. 204 LCT', 'Recargo del 50% primeras 2 horas extras'],
            ['Art. 204 LCT', 'Recargo del 100% horas extras adicionales'],
            ['Art. 200 LCT', 'Trabajo nocturno (21:00 a 06:00)'],
            ['Art. 204 LCT', 'Sábados después de 13:00 = 100%'],
            ['Art. 204 LCT', 'Domingos y feriados = 100%']
        ]
        
        for row, (label, value) in enumerate(config_data, 5):
            ws.cell(row=row, column=1, value=label).font = Font(bold=True)
            ws.cell(row=row, column=2, value=value)
        
        # Ajustar ancho de columnas
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 30
