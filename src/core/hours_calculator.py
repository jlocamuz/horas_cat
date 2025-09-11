"""
Calculador de Horas según Normativa Argentina
"""

from datetime import datetime
from typing import Dict, List, Optional, Set
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
        self.holiday_names = DEFAULT_CONFIG.get('holiday_names', {})

    # -------------------- Helpers --------------------

    def _get_ref_str(self, day_summary: Dict) -> str:
        ref = day_summary.get('referenceDate') or day_summary.get('date') or ''
        return ref[:10]

    def _get_night_hours_from_summary(self, day_summary: Dict) -> float:
        night_hours = 0.0
        for cat in (day_summary.get('categorizedHours') or []):
            name = (cat.get('category') or {}).get('name')
            if name in ('NIGHT', 'NOCTURNA'):
                try:
                    night_hours += float(cat.get('hours') or 0)
                except (TypeError, ValueError):
                    pass
        return night_hours

    def _get_start_end_iso(self, day_summary: Dict):
        start_iso = end_iso = None
        for e in (day_summary.get('entries') or []):
            if e.get('type') == 'START' and not start_iso:
                start_iso = (e.get('time') or e.get('date') or '')[:19]
            elif e.get('type') == 'END' and not end_iso:
                end_iso = (e.get('time') or e.get('date') or '')[:19]
        return start_iso, end_iso

    def _hhmm(self, iso: Optional[str]) -> Optional[str]:
        if not iso:
            return None
        return iso[11:16]

    def _crosses_into_holiday(self, day_summary: Dict, ref_str: str, holiday_dates: Set[str]) -> Optional[str]:
        """
        Si el END cae en un día distinto y ese día es feriado, devuelve esa fecha (YYYY-MM-DD).
        De lo contrario, devuelve None.
        """
        for e in (day_summary.get('entries') or []):
            if e.get('type') == 'END':
                end_iso = (e.get('time') or e.get('date') or '')[:10]
                if end_iso and end_iso != ref_str and end_iso in holiday_dates:
                    return end_iso
        return None

    def _get_holiday_name(self, date_str: str, day_summary: Dict) -> Optional[str]:
        # 1) si viene desde la API
        if day_summary.get('holidays'):
            name = (day_summary['holidays'][0] or {}).get('name')
            if name:
                return name
        # 2) si está en el config
        return self.holiday_names.get(date_str)

    # -------------------- Cálculo principal --------------------

# -------------------- Cálculo principal --------------------
    def process_employee_data(self, day_summaries: List[Dict], employee_info: Dict,
                            previous_pending_hours: float = 0,
                            holidays: Optional[Set[str]] = None) -> Dict:

        holiday_dates = set(holidays) if holidays else set(DEFAULT_CONFIG.get('holidays', []))

        daily_data = []
        totals = {
            'total_days_worked': 0.0,
            'total_hours_worked': 0.0,
            'total_regular_hours': 0.0,
            'total_extra_hours_50': 0.0,
            'total_extra_hours_100': 0.0,
            'total_night_hours': 0.0,
            'total_pending_hours': float(previous_pending_hours)
        }

        for day_summary in day_summaries:
            ref_str = self._get_ref_str(day_summary)
            if not ref_str:
                continue
            ref_dt = datetime.strptime(ref_str, '%Y-%m-%d')

            hours_worked = float(day_summary.get('hours', {}).get('worked', 0) or day_summary.get('totalHours', 0) or 0)
            is_holiday_api = bool(day_summary.get('holidays'))
            has_time_off = bool(day_summary.get('timeOffRequests'))
            has_absence = 'ABSENT' in (day_summary.get('incidences') or [])
            is_rest_day = not bool(day_summary.get('isWorkday', True))  # ← FRANCO

            if hours_worked == 0 and not has_time_off:
                continue

            # detecto cruce a feriado
            end_holiday_str = self._crosses_into_holiday(day_summary, ref_str, holiday_dates)
            is_ref_holiday_cfg = ref_str in holiday_dates
            is_out_holiday_cfg = bool(end_holiday_str)

            # ¿a qué fecha asigno la fila?
            if is_out_holiday_cfg:
                out_date_str = end_holiday_str          # 22→06 y el 06 es feriado ⇒ asigno al fin
            else:
                out_date_str = ref_str                   # franco nocturno ⇒ asigno al inicio

            # ¿es feriado para tasa 100%?
            is_holiday_output = is_holiday_api or is_ref_holiday_cfg or is_out_holiday_cfg
            holiday_name = None
            if is_holiday_output:
                holiday_name = self._get_holiday_name(out_date_str, day_summary) or \
                            self._get_holiday_name(ref_str, day_summary)

            # distribución de horas
            night_hours = self._get_night_hours_from_summary(day_summary)
            if is_holiday_output:
                # Feriado (del día o por cruce) ⇒ 100%
                day_hours = {
                    'hours_worked': hours_worked, 'regular_hours': 0.0,
                    'extra_hours_50': 0.0, 'extra_hours_100': hours_worked,
                    'night_hours': night_hours, 'pending_hours': 0.0
                }
            elif is_rest_day and hours_worked > 0:
                # ← NUEVO: FRANCO TRABAJADO ⇒ TODO AL 100%
                day_hours = {
                    'hours_worked': hours_worked, 'regular_hours': 0.0,
                    'extra_hours_50': 0.0, 'extra_hours_100': hours_worked,
                    'night_hours': night_hours, 'pending_hours': 0.0
                }
            else:
                # Día laborable normal
                day_hours = self.calculate_hour_distribution(
                    hours_worked=hours_worked,
                    date=ref_dt,  # distribución por el día de INICIO (franco nocturno)
                    is_holiday=False,
                    has_time_off=has_time_off,
                    night_hours=night_hours
                )

            # acumulo totales
            if hours_worked > 0:
                totals['total_days_worked'] += 1
                totals['total_hours_worked'] += day_hours['hours_worked']
                totals['total_regular_hours'] += day_hours['regular_hours']
                totals['total_extra_hours_50'] += day_hours['extra_hours_50']
                totals['total_extra_hours_100'] += day_hours['extra_hours_100']
                totals['total_night_hours'] += day_hours['night_hours']
                if not has_time_off and not has_absence:
                    totals['total_pending_hours'] += day_hours['pending_hours']

            # explicación para la hoja
            start_iso, end_iso = self._get_start_end_iso(day_summary)
            start_d = (start_iso[:10] if start_iso else ref_str)
            end_d = (end_iso[:10] if end_iso else ref_str)
            start_h = self._hhmm(start_iso) or ''
            end_h = self._hhmm(end_iso) or ''

            exp_parts = []
            if start_iso or end_iso:
                exp_parts.append(f"Inicio {start_d} {start_h} → Fin {end_d} {end_h}.")
            else:
                exp_parts.append("Sin marcajes, se usa resumen del día.")

            if is_out_holiday_cfg:
                exp_parts.append(
                    f"Fin cae en feriado {holiday_name or 'sin nombre'} ⇒ "
                    f"{hours_worked:.2f}h al 100% y fecha asignada {out_date_str}."
                )
            elif is_holiday_api or is_ref_holiday_cfg:
                exp_parts.append(
                    f"Feriado {holiday_name or 'sin nombre'} ⇒ {hours_worked:.2f}h al 100%."
                )
            elif is_rest_day:
                exp_parts.append(
                    f"Franco trabajado (día de descanso) ⇒ {hours_worked:.2f}h al 100%."
                )
            else:
                dow = ref_dt.weekday()
                if dow == 6:
                    exp_parts.append("Domingo ⇒ todo al 100%.")
                elif dow == 5:
                    exp_parts.append("Sábado ⇒ todo al 50% (simplificado).")
                else:
                    if day_hours['hours_worked'] <= self.jornada_completa:
                        if day_hours['pending_hours'] > 0:
                            exp_parts.append(
                                f"Lun–Vie: {day_hours['hours_worked']:.2f}h regulares + "
                                f"{day_hours['pending_hours']:.2f}h pendientes."
                            )
                        else:
                            exp_parts.append("Lun–Vie: horas dentro de la jornada regular.")
                    else:
                        extra = day_hours['hours_worked'] - self.jornada_completa
                        exp_parts.append(
                            f"Lun–Vie: 8h regulares + {min(extra,2):.2f}h 50% + {max(0, extra-2):.2f}h 100%."
                        )

            # nota de cálculo
            calc_note = " ".join(exp_parts)

            # fila diaria
            daily_data.append({
                'employee_id': employee_info.get('employeeInternalId'),
                'date': out_date_str,
                'day_of_week': self.get_day_of_week_spanish(datetime.strptime(out_date_str, '%Y-%m-%d')),
                'hours_worked': day_hours['hours_worked'],
                'regular_hours': day_hours['regular_hours'],
                'extra_hours_50': day_hours['extra_hours_50'],
                'extra_hours_100': day_hours['extra_hours_100'],
                'night_hours': day_hours['night_hours'],
                'pending_hours': day_hours['pending_hours'] if not (has_time_off or has_absence) else 0.0,
                'is_holiday': is_holiday_output,          # franco ≠ feriado ⇒ queda en "No"
                'holiday_name': holiday_name,
                'has_time_off': has_time_off,
                'time_off_name': (day_summary.get('timeOffRequests') or [{}])[0].get('name') if has_time_off else None,
                'has_absence': has_absence,
                'is_full_time': day_hours['hours_worked'] >= self.jornada_completa,
                'calc_note': calc_note,
            })

        # compensaciones
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

    # -------------------- Distribución estándar --------------------

    def calculate_hour_distribution(self, hours_worked: float, date: datetime,
                                    is_holiday: bool = False, has_time_off: bool = False,
                                    night_hours: float = 0.0) -> Dict:
        if hours_worked == 0:
            return {
                'hours_worked': 0.0,
                'regular_hours': 0.0,
                'extra_hours_50': 0.0,
                'extra_hours_100': 0.0,
                'night_hours': float(night_hours),
                'pending_hours': 0.0
            }

        day_of_week = date.weekday()  # 0=Lun … 6=Dom
        regular_hours = 0.0
        extra_hours_50 = 0.0
        extra_hours_100 = 0.0
        pending_hours = 0.0

        if day_of_week == 6:  # Domingo
            extra_hours_100 = hours_worked
        elif day_of_week == 5:  # Sábado (simplificado)
            extra_hours_50 = hours_worked
        else:  # Lun–Vie
            if hours_worked <= self.jornada_completa:
                regular_hours = hours_worked
                if not has_time_off and hours_worked < self.jornada_completa:
                    pending_hours = self.jornada_completa - hours_worked
            else:
                regular_hours = float(self.jornada_completa)
                extra = hours_worked - self.jornada_completa
                if extra <= 2:
                    extra_hours_50 = extra
                else:
                    extra_hours_50 = 2.0
                    extra_hours_100 = extra - 2.0

        return {
            'hours_worked': float(hours_worked),
            'regular_hours': float(regular_hours),
            'extra_hours_50': float(extra_hours_50),
            'extra_hours_100': float(extra_hours_100),
            'night_hours': float(night_hours),
            'pending_hours': float(pending_hours)
        }

    # -------------------- Otras utilidades --------------------

    def calculate_compensations(self, extra_hours_50: float, extra_hours_100: float, pending_hours: float) -> Dict:
        compensated_with_50 = 0.0
        compensated_with_100 = 0.0
        remaining = float(pending_hours)

        if remaining > 0 and extra_hours_50 > 0:
            compensated_with_50 = min(remaining, extra_hours_50)
            remaining -= compensated_with_50

        if remaining > 0 and extra_hours_100 > 0:
            max_comp_100 = extra_hours_100 * 1.5
            compensated_with_100 = min(remaining, max_comp_100)
            remaining -= compensated_with_100

        net50 = float(extra_hours_50) - compensated_with_50
        net100 = float(extra_hours_100) - (compensated_with_100 / 1.5 if compensated_with_100 else 0.0)

        return {
            'compensated_with_50': float(compensated_with_50),
            'compensated_with_100': float(compensated_with_100),
            'net_extra_hours_50': float(net50),
            'net_extra_hours_100': float(net100),
            'remaining_pending_hours': float(remaining)
        }

    def get_day_of_week_spanish(self, date: datetime) -> str:
        days = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
        return days[date.weekday()]

    def is_night_hour(self, hour: int) -> bool:
        return hour >= self.hora_nocturna_inicio or hour < self.hora_nocturna_fin

    def format_hours(self, hours: float) -> str:
        return '0.00' if hours == 0 else f"{hours:.2f}"

    def format_hours_to_hhmm(self, hours: float) -> str:
        total_minutes = round(hours * 60)
        h = total_minutes // 60
        m = total_minutes % 60
        return f"{h:02d}:{m:02d}"

    def minutes_to_hours(self, minutes: int) -> float:
        return round(minutes / 60, 2)

    def round_to_fragment(self, minutes: int) -> int:
        import math
        return math.ceil(minutes / self.fragmento_minutos) * self.fragmento_minutos


# Compatibilidad
def process_employee_data_from_day_summaries(day_summaries: List[Dict], employee_info: Dict,
                                             previous_pending_hours: float = 0,
                                             period_dates: Dict = None, holidays: Optional[Set[str]] = None) -> Dict:
    calc = ArgentineHoursCalculator()
    return calc.process_employee_data(day_summaries, employee_info, previous_pending_hours, holidays or set())


def calculate_compensations(extra_hours_50: float, extra_hours_100: float, pending_hours: float) -> Dict:
    calc = ArgentineHoursCalculator()
    return calc.calculate_compensations(extra_hours_50, extra_hours_100, pending_hours)
