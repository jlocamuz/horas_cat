"""
Calculador de Horas según Normativa Argentina
Adaptado del proyecto Horas-cat-v3
"""

from datetime import datetime, timedelta
from typing import Dict, List, Optional
from config.default_config import DEFAULT_CONFIG


class ArgentineHoursCalculator:
    """Calculador de horas según normativa laboral argentina"""
    
    def __init__(self):
        self.jornada_completa = DEFAULT_CONFIG['jornada_completa_horas']
        self.hora_nocturna_inicio = DEFAULT_CONFIG['hora_nocturna_inicio']
        self.hora_nocturna_fin = DEFAULT_CONFIG['hora_nocturna_fin']
        self.sabado_limite = DEFAULT_CONFIG['sabado_limite_hora']
        self.tolerancia_minutos = DEFAULT_CONFIG['tolerancia_minutos']
        self.fragmento_minutos = DEFAULT_CONFIG['fragmento_minutos']
    
    def process_employee_data(self, day_summaries: List[Dict], employee_info: Dict, 
                            previous_pending_hours: float = 0, holidays: List[Dict] = None) -> Dict:
        """
        Procesa los datos de un empleado desde day summaries
        Args:
            day_summaries: Lista de resúmenes diarios
            employee_info: Información del empleado
            previous_pending_hours: Horas pendientes del período anterior
            holidays: Lista de feriados
        Returns:
            Datos procesados del empleado
        """
        daily_data = []
        totals = {
            'total_days_worked': 0,
            'total_hours_worked': 0,
            'total_regular_hours': 0,
            'total_extra_hours_50': 0,
            'total_extra_hours_100': 0,
            'total_night_hours': 0,
            'total_pending_hours': previous_pending_hours
        }
        
        # Procesar cada day summary
        for day_summary in day_summaries:
            date = datetime.strptime(day_summary.get('referenceDate', day_summary.get('date')), '%Y-%m-%d')
            
            # Extraer información del day summary
            hours_worked = day_summary.get('hours', {}).get('worked', 0) or day_summary.get('totalHours', 0)
            is_holiday = bool(day_summary.get('holidays') and len(day_summary.get('holidays', [])) > 0)
            has_time_off = bool(day_summary.get('timeOffRequests') and len(day_summary.get('timeOffRequests', [])) > 0)
            has_absence = bool(day_summary.get('incidences') and 'ABSENT' in day_summary.get('incidences', []))
            
            # Si no hay horas trabajadas y no hay licencia, saltar
            if hours_worked == 0 and not has_time_off:
                continue
            
            # Calcular horas categorizadas usando la lógica argentina
            day_hours = self.calculate_argentine_hours_from_summary(day_summary, date, holidays)
            
            if hours_worked > 0:
                totals['total_days_worked'] += 1
                totals['total_hours_worked'] += day_hours['hours_worked']
                totals['total_regular_hours'] += day_hours['regular_hours']
                totals['total_extra_hours_50'] += day_hours['extra_hours_50']
                totals['total_extra_hours_100'] += day_hours['extra_hours_100']
                totals['total_night_hours'] += day_hours['night_hours']
                
                # Solo agregar horas pendientes si no hay licencia
                if not has_time_off and not has_absence:
                    totals['total_pending_hours'] += day_hours['pending_hours']
            
            daily_data.append({
                'employee_id': employee_info.get('employeeInternalId'),
                'date': date.strftime('%Y-%m-%d'),
                'day_of_week': self.get_day_of_week_spanish(date),
                'hours_worked': day_hours['hours_worked'],
                'regular_hours': day_hours['regular_hours'],
                'extra_hours_50': day_hours['extra_hours_50'],
                'extra_hours_100': day_hours['extra_hours_100'],
                'night_hours': day_hours['night_hours'],
                'pending_hours': day_hours['pending_hours'] if not (has_time_off or has_absence) else 0,
                'is_holiday': is_holiday,
                'holiday_name': day_summary.get('holidays', [{}])[0].get('name') if is_holiday else None,
                'has_time_off': has_time_off,
                'time_off_name': day_summary.get('timeOffRequests', [{}])[0].get('name') if has_time_off else None,
                'has_absence': has_absence,
                'is_full_time': day_hours['hours_worked'] >= self.jornada_completa
            })
        
        # Calcular compensaciones
        compensations = self.calculate_compensations(
            totals['total_extra_hours_50'], 
            totals['total_extra_hours_100'], 
            totals['total_pending_hours']
        )
        
        return {
            'employee_info': employee_info,
            'daily_data': daily_data,
            'totals': totals,
            'compensations': compensations
        }
    
    def calculate_argentine_hours_from_summary(self, day_summary: Dict, date: datetime, 
                                             holidays: List[Dict] = None) -> Dict:
        """
        Calcula horas argentinas desde un day summary
        Args:
            day_summary: Day summary de la API
            date: Fecha del día
            holidays: Lista de feriados
        Returns:
            Horas categorizadas
        """
        hours_worked = day_summary.get('hours', {}).get('worked', 0) or day_summary.get('totalHours', 0)
        
        if hours_worked == 0:
            return {
                'hours_worked': 0,
                'regular_hours': 0,
                'extra_hours_50': 0,
                'extra_hours_100': 0,
                'night_hours': 0,
                'pending_hours': 0
            }
        
        # Determinar tipo de día
        day_type = self.get_day_type(date)
        
        # Verificar si es feriado (prioridad sobre tipo de día)
        is_holiday = day_summary.get('holidays') and len(day_summary.get('holidays', [])) > 0
        if is_holiday:
            day_type = 'HOLIDAY'
        
        # Calcular horas nocturnas si están disponibles
        night_hours = 0
        if day_summary.get('categorizedHours'):
            night_category = next(
                (cat for cat in day_summary['categorizedHours'] 
                 if cat.get('category', {}).get('name') in ['NIGHT', 'NOCTURNA']), 
                None
            )
            if night_category:
                night_hours = night_category.get('hours', 0)
        
        # Aplicar lógica de distribución de horas
        return self.calculate_hour_distribution(hours_worked, date, is_holiday, 
                                              day_summary.get('hasTimeOff', False), night_hours)
    
    def calculate_hour_distribution(self, hours_worked: float, date: datetime, 
                                  is_holiday: bool = False, has_time_off: bool = False,
                                  night_hours: float = 0) -> Dict:
        """
        Distribuye las horas trabajadas según la normativa argentina
        Args:
            hours_worked: Horas trabajadas en el día
            date: Fecha del día
            is_holiday: Si es feriado
            has_time_off: Si tiene licencia
            night_hours: Horas nocturnas calculadas
        Returns:
            Distribución de horas
        """
        if hours_worked == 0:
            return {
                'hours_worked': hours_worked,
                'regular_hours': 0,
                'extra_hours_50': 0,
                'extra_hours_100': 0,
                'night_hours': night_hours,
                'pending_hours': 0
            }
        
        day_of_week = date.weekday()  # 0 = Lunes, 6 = Domingo
        regular_hours = 0
        extra_hours_50 = 0
        extra_hours_100 = 0
        pending_hours = 0
        
        if is_holiday or day_of_week == 6:  # Feriados y domingos
            # Todas las horas al 100%
            extra_hours_100 = hours_worked
        elif day_of_week == 5:  # Sábados
            # Todas las horas al 50% (simplificado, en realidad depende del horario)
            extra_hours_50 = hours_worked
        else:  # Lunes a viernes (días laborables)
            if hours_worked <= self.jornada_completa:
                regular_hours = hours_worked
                # Calcular horas pendientes solo si no hay licencia
                if not has_time_off and hours_worked < self.jornada_completa:
                    pending_hours = self.jornada_completa - hours_worked
            else:
                regular_hours = self.jornada_completa
                extra_hours = hours_worked - self.jornada_completa
                
                # Primeras 2 horas extras al 50%, el resto al 100%
                if extra_hours <= 2:
                    extra_hours_50 = extra_hours
                else:
                    extra_hours_50 = 2
                    extra_hours_100 = extra_hours - 2
        
        return {
            'hours_worked': hours_worked,
            'regular_hours': regular_hours,
            'extra_hours_50': extra_hours_50,
            'extra_hours_100': extra_hours_100,
            'night_hours': night_hours,
            'pending_hours': pending_hours
        }
    
    def calculate_compensations(self, extra_hours_50: float, extra_hours_100: float, 
                              pending_hours: float) -> Dict:
        """
        Calcula las compensaciones de horas extras con horas pendientes
        Args:
            extra_hours_50: Horas extras al 50%
            extra_hours_100: Horas extras al 100%
            pending_hours: Horas pendientes
        Returns:
            Compensaciones calculadas
        """
        compensated_with_50 = 0
        compensated_with_100 = 0
        remaining_pending_hours = pending_hours
        
        # Primero compensar con horas al 50% (1:1)
        if remaining_pending_hours > 0 and extra_hours_50 > 0:
            compensated_with_50 = min(remaining_pending_hours, extra_hours_50)
            remaining_pending_hours -= compensated_with_50
        
        # Luego compensar con horas al 100% (1:1.5, es decir, 1 hora al 100% compensa 1.5 horas pendientes)
        if remaining_pending_hours > 0 and extra_hours_100 > 0:
            max_compensation_with_100 = extra_hours_100 * 1.5
            compensated_with_100 = min(remaining_pending_hours, max_compensation_with_100)
            remaining_pending_hours -= compensated_with_100
        
        # Calcular horas netas después de compensaciones
        net_extra_hours_50 = extra_hours_50 - compensated_with_50
        net_extra_hours_100 = extra_hours_100 - (compensated_with_100 / 1.5)
        
        return {
            'compensated_with_50': compensated_with_50,
            'compensated_with_100': compensated_with_100,
            'net_extra_hours_50': net_extra_hours_50,
            'net_extra_hours_100': net_extra_hours_100,
            'remaining_pending_hours': remaining_pending_hours
        }
    
    def get_day_type(self, date: datetime) -> str:
        """
        Obtiene el tipo de día para cálculos
        Args:
            date: Fecha a evaluar
        Returns:
            Tipo de día: WEEKDAY, SATURDAY, SUNDAY, HOLIDAY
        """
        day_of_week = date.weekday()
        if day_of_week == 6:  # Domingo
            return 'SUNDAY'
        elif day_of_week == 5:  # Sábado
            return 'SATURDAY'
        else:  # Lunes a viernes
            return 'WEEKDAY'
    
    def get_day_of_week_spanish(self, date: datetime) -> str:
        """
        Obtiene el nombre del día de la semana en español
        Args:
            date: Fecha
        Returns:
            Nombre del día en español
        """
        days = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
        return days[date.weekday()]
    
    def is_night_hour(self, hour: int) -> bool:
        """
        Determina si una hora es nocturna (21:00 a 06:00)
        Args:
            hour: Hora (0-23)
        Returns:
            True si es hora nocturna
        """
        return hour >= self.hora_nocturna_inicio or hour < self.hora_nocturna_fin
    
    def format_hours(self, hours: float) -> str:
        """
        Formatea horas para mostrar
        Args:
            hours: Horas en formato decimal
        Returns:
            Horas formateadas
        """
        if hours == 0:
            return '0.00'
        return f"{hours:.2f}"
    
    def format_hours_to_hhmm(self, hours: float) -> str:
        """
        Convierte horas decimales a formato HH:MM
        Args:
            hours: Horas en formato decimal
        Returns:
            Formato HH:MM
        """
        total_minutes = round(hours * 60)
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h:02d}:{m:02d}"
    
    def minutes_to_hours(self, minutes: int) -> float:
        """
        Convierte minutos a horas decimales
        Args:
            minutes: Minutos
        Returns:
            Horas decimales
        """
        return round(minutes / 60, 2)
    
    def round_to_fragment(self, minutes: int) -> int:
        """
        Redondea minutos al fragmento más cercano
        Args:
            minutes: Minutos a redondear
        Returns:
            Minutos redondeados
        """
        import math
        return math.ceil(minutes / self.fragmento_minutos) * self.fragmento_minutos


# Funciones de utilidad para compatibilidad con el proyecto anterior
def process_employee_data_from_day_summaries(day_summaries: List[Dict], employee_info: Dict, 
                                           previous_pending_hours: float = 0, 
                                           period_dates: Dict = None, holidays: List[Dict] = None) -> Dict:
    """Función de compatibilidad con el proyecto anterior"""
    calculator = ArgentineHoursCalculator()
    return calculator.process_employee_data(day_summaries, employee_info, previous_pending_hours, holidays)


def calculate_compensations(extra_hours_50: float, extra_hours_100: float, pending_hours: float) -> Dict:
    """Función de compatibilidad con el proyecto anterior"""
    calculator = ArgentineHoursCalculator()
    return calculator.calculate_compensations(extra_hours_50, extra_hours_100, pending_hours)
