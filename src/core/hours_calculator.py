"""
Calculador de Horas según Normativa Argentina
"""

from datetime import datetime, timedelta
from typing import Dict, List, Optional, Set, Tuple
from zoneinfo import ZoneInfo  # stdlib (Python >=3.9)
from config.default_config import DEFAULT_CONFIG


class ArgentineHoursCalculator:
    """Calculador de horas según normativa laboral argentina"""

    def __init__(self):
        self.jornada_completa     = DEFAULT_CONFIG['jornada_completa_horas']
        self.hora_nocturna_inicio = DEFAULT_CONFIG['hora_nocturna_inicio']  # ej. 21
        self.hora_nocturna_fin    = DEFAULT_CONFIG['hora_nocturna_fin']    # ej. 6
        self.sabado_limite        = DEFAULT_CONFIG.get('sabado_limite_hora', 13)
        self.tolerancia_minutos   = DEFAULT_CONFIG['tolerancia_minutos']
        self.fragmento_minutos    = DEFAULT_CONFIG['fragmento_minutos']
        self.holiday_names        = DEFAULT_CONFIG.get('holiday_names', {})
        self.local_tz             = ZoneInfo(DEFAULT_CONFIG.get('local_timezone',
                                                                'America/Argentina/Buenos_Aires'))
        self.extras_al_50         = DEFAULT_CONFIG.get("extras_al_50", 2)  # p.ej. 4 en ARM

    # -------------------- Helpers de parsing / fechas --------------------

    def _get_ref_str(self, day_summary: Dict) -> str:
        ref = day_summary.get('referenceDate') or day_summary.get('date') or ''
        return ref[:10]

    def _parse_iso_to_local(self, s: Optional[str]) -> Optional[datetime]:
        """
        Convierte ISO (con 'Z' u offset) a datetime **local** (naive).
        Si no trae tz, se asume local.
        """
        if not s:
            return None
        s = s.replace('Z', '+00:00')  # normalizo 'Z'
        try:
            dt = datetime.fromisoformat(s)
            if dt.tzinfo is None:
                return dt  # ya está en local
            return dt.astimezone(self.local_tz).replace(tzinfo=None)
        except Exception:
            return None

    def _first_entry_pair_local(self, day_summary: Dict) -> Tuple[Optional[datetime], Optional[datetime]]:
        """
        Toma el primer START y el primer END de entries, los convierte a **local** y
        devuelve (start_local, end_local). Maneja cruce de día si end <= start.
        """
        start_iso = end_iso = None
        for e in (day_summary.get('entries') or []):
            if e.get('type') == 'START' and not start_iso:
                start_iso = e.get('time') or e.get('date')
            elif e.get('type') == 'END' and not end_iso:
                end_iso = e.get('time') or e.get('date')

        s_dt = self._parse_iso_to_local(start_iso[:25] if start_iso else None)
        e_dt = self._parse_iso_to_local(end_iso[:25] if end_iso else None)
        if s_dt and e_dt and e_dt <= s_dt:
            e_dt += timedelta(days=1)  # cruza medianoche
        return s_dt, e_dt

    def _display_from_entries(self, day_summary: Dict) -> Tuple[str, str, str, str]:
        """
        Devuelve (start_date, start_hhmm, end_date, end_hhmm) usando ENTRIES en local.
        Si faltan, devuelve strings vacíos anclados al ref_str.
        """
        ref_str = self._get_ref_str(day_summary)
        s_dt, e_dt = self._first_entry_pair_local(day_summary)
        if not (s_dt and e_dt):
            return ref_str, "", ref_str, ""
        return (
            s_dt.strftime("%Y-%m-%d"),
            s_dt.strftime("%H:%M"),
            e_dt.strftime("%Y-%m-%d"),
            e_dt.strftime("%H:%M"),
        )

    def _get_intervals_from_entries(self, day_summary: Dict) -> List[Tuple[datetime, datetime]]:
        """
        Devuelve [(start_local, end_local)] usando entries.
        (Si quisieras soportar varios pares START/END, expandí acá).
        """
        s_dt, e_dt = self._first_entry_pair_local(day_summary)
        return [(s_dt, e_dt)] if (s_dt and e_dt) else []

    def _get_holiday_name(self, date_str: str, day_summary: Dict) -> Optional[str]:
        # 1) si viene desde la API
        if day_summary.get('holidays'):
            name = (day_summary['holidays'][0] or {}).get('name')
            if name:
                return name
        # 2) si está en el config
        return self.holiday_names.get(date_str)

    # -------------------- Intersecciones / nocturnas --------------------

    def _intersect_hours(self, a_start: datetime, a_end: datetime,
                         b_start: datetime, b_end: datetime) -> float:
        start = max(a_start, b_start)
        end = min(a_end, b_end)
        return max(0.0, (end - start).total_seconds() / 3600)

    def _compute_night_hours_from_intervals(self, intervals: List[Tuple[datetime, datetime]],
                                            ref_dt: datetime) -> float:
        """
        Ventana nocturna anclada al **día de inicio** (ref_dt): 21:00 → 06:00 del día siguiente.
        """
        n_start = ref_dt.replace(hour=self.hora_nocturna_inicio, minute=0, second=0, microsecond=0)
        n_end   = (ref_dt + timedelta(days=1)).replace(hour=self.hora_nocturna_fin, minute=0, second=0, microsecond=0)
        total = 0.0
        for s_dt, e_dt in intervals:
            total += self._intersect_hours(s_dt, e_dt, n_start, n_end)
        return round(total, 2)

    # -------------------- Feriado por FIN local --------------------

    def _crosses_into_holiday_local_end(self, day_summary: Dict,
                                        ref_str: str,
                                        holiday_dates: Set[str]) -> Optional[str]:
        """
        Si el END (en hora LOCAL) cae en un día distinto y ese día es feriado,
        devuelve esa fecha (YYYY-MM-DD). Si no, None.
        """
        _, e_dt = self._first_entry_pair_local(day_summary)
        if not e_dt:
            return None
        end_date_local = e_dt.strftime("%Y-%m-%d")
        if end_date_local != ref_str and end_date_local in holiday_dates:
            return end_date_local
        return None

    # -------------------- Distribución auxiliar (Lun–Vie) --------------------

    def _weekday_distribution(self, hours: float, has_time_off: bool) -> Dict:
        """
        Reparte horas como Lun–Vie: regulares hasta jornada, luego extras 50% hasta
        'extras_al_50' y el resto 100%.
        """
        if hours <= 0:
            return {'regular': 0.0, 'extra50': 0.0, 'extra100': 0.0, 'pending': 0.0}

        regular = min(hours, float(self.jornada_completa))
        extra = max(0.0, hours - float(self.jornada_completa))

        if extra <= self.extras_al_50:
            e50 = extra
            e100 = 0.0
        else:
            e50 = float(self.extras_al_50)
            e100 = extra - float(self.extras_al_50)

        pending = 0.0
        if not has_time_off and hours < self.jornada_completa:
            pending = float(self.jornada_completa) - hours

        return {'regular': regular, 'extra50': e50, 'extra100': e100, 'pending': pending}

    # -------------------- Cálculo principal --------------------

    def process_employee_data(self, day_summaries: List[Dict], employee_info: Dict,
                              previous_pending_hours: float = 0,
                              holidays: Optional[Set[str]] = None) -> Dict:

        holiday_dates = set(holidays) if holidays else set(DEFAULT_CONFIG.get('holidays', []))

        turno = ''

        for seg in employee_info.get("segmentations", []):
            if seg.get("group") == "Turno":
                turno = seg['item']
                employee_info["turno"] = turno
        daily_data: List[Dict] = []
        totals = {
            'total_days_worked': 0.0,
            'total_hours_worked': 0.0,
            'total_regular_hours': 0.0,
            'total_extra_hours_50': 0.0,
            'total_extra_hours_100': 0.0,
            'total_night_hours': 0.0,
            'total_holiday_hours': 0.0,          # NUEVO
            'total_pending_hours': float(previous_pending_hours)
        }

        for day_summary in day_summaries:
            
                      
            ref_str = self._get_ref_str(day_summary)
            if not ref_str:
                continue
            ref_dt = datetime.strptime(ref_str, '%Y-%m-%d')
            dow = ref_dt.weekday()  # 0=Lun … 6=Dom

            hours_worked = float(day_summary.get('hours', {}).get('worked', 0)
                                 or day_summary.get('totalHours', 0) or 0)
            is_holiday_api = bool(day_summary.get('holidays'))
            has_time_off   = bool(day_summary.get('timeOffRequests'))
            has_absence    = 'ABSENT' in (day_summary.get('incidences') or [])
            is_rest_day    = not bool(day_summary.get('isWorkday', True))  # FRANCO

            if hours_worked == 0 and not has_time_off:
                continue

            # Feriado por fin local
            end_holiday_str    = self._crosses_into_holiday_local_end(day_summary, ref_str, holiday_dates)
            is_ref_holiday_cfg = ref_str in holiday_dates
            is_out_holiday_cfg = bool(end_holiday_str)

            # ¿A qué fecha imputo?
            out_date_str = end_holiday_str if is_out_holiday_cfg else ref_str

            # ¿Es feriado para tasa 100%?
            is_holiday_output = is_holiday_api or is_ref_holiday_cfg or is_out_holiday_cfg
            holiday_name = None
            if is_holiday_output:
                holiday_name = self._get_holiday_name(out_date_str, day_summary) or \
                               self._get_holiday_name(ref_str, day_summary)

            # Intervalos (from ENTRIES en local) y nocturnas
            intervals = self._get_intervals_from_entries(day_summary)
            night_hours = self._compute_night_hours_from_intervals(intervals, ref_dt) \
                          if intervals else 0.0

            # Horas “feriado” (solo si es feriado, no domingo)
            holiday_hours = hours_worked if is_holiday_output else 0.0

            # ---------------- Lógica de distribución ----------------
            if is_holiday_output:
                # Feriado (del día o por cruce) ⇒ TODO 100%
                regular_hours = 0.0
                extra50 = 0.0
                extra100 = hours_worked
                pending = 0.0

            elif is_rest_day and hours_worked > 0:
                # Franco trabajado ⇒ TODO 100%
                regular_hours = 0.0
                extra50 = 0.0
                extra100 = hours_worked
                pending = 0.0

            elif dow == 6:
                # Domingo ⇒ TODO 100%
                regular_hours = 0.0
                extra50 = 0.0
                extra100 = hours_worked
                pending = 0.0

            elif dow == 5:
                # Sábado (NO feriado, NO franco):
                #   - Porción después de 13:00 del sábado ⇒ 100%
                #   - Porción en domingo (si el turno cruza) ⇒ 100%
                #   - El resto (antes de 13) ⇒ distribución Lun–Vie
                if intervals:
                    sat_13 = ref_dt.replace(hour=self.sabado_limite, minute=0, second=0, microsecond=0)
                    sat_24 = (ref_dt + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                    sat100 = 0.0
                    for s_dt, e_dt in intervals:
                        sat100 += self._intersect_hours(s_dt, e_dt, sat_13, sat_24)

                    # Si el día siguiente es domingo, también es 100% (cualquier tramo en ese domingo)
                    next_day = ref_dt + timedelta(days=1)
                    sun100 = 0.0
                    if next_day.weekday() == 6:  # domingo
                        sun_start = next_day.replace(hour=0, minute=0, second=0, microsecond=0)
                        sun_end   = (next_day + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                        for s_dt, e_dt in intervals:
                            sun100 += self._intersect_hours(s_dt, e_dt, sun_start, sun_end)
                    weekend_100 = round(sat100 + sun100, 2)
                else:
                    # Sin intervals (entries) no podemos cortar por 13:00 ⇒
                    # lo más seguro: tratamos todo como Lun–Vie (antes de 13) y NO marcamos 100.
                    weekend_100 = 0.0

                base_hours = max(0.0, hours_worked - weekend_100)
                dist = self._weekday_distribution(base_hours, has_time_off)
                regular_hours = dist['regular']
                extra50 = dist['extra50']
                extra100 = dist['extra100'] + weekend_100
                pending = dist['pending']

            else:
                # Lun–Vie normal
                dist = self._weekday_distribution(hours_worked, has_time_off)
                regular_hours = dist['regular']
                extra50 = dist['extra50']
                extra100 = dist['extra100']
                pending = dist['pending']

            # ---------------- Acumulo totales ----------------
            totals['total_days_worked']     += 1
            totals['total_hours_worked']    += hours_worked
            totals['total_regular_hours']   += regular_hours
            totals['total_extra_hours_50']  += extra50
            totals['total_extra_hours_100'] += extra100
            totals['total_night_hours']     += night_hours
            totals['total_holiday_hours']   += holiday_hours
            if not has_time_off and not has_absence:
                totals['total_pending_hours'] += pending

            # ---- Nota (horarios locales) ----
            disp_start_d, disp_start_h, disp_end_d, disp_end_h = self._display_from_entries(day_summary)
            exp_parts = []
            if disp_start_h or disp_end_h:
                exp_parts.append(f"Inicio {disp_start_d} {disp_start_h} → Fin {disp_end_d} {disp_end_h}.")
            else:
                exp_parts.append("Sin marcajes, se usa resumen del día.")
            exp_parts.append("Imputado al día de inicio del turno.")

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
                exp_parts.append(f"Franco trabajado ⇒ {hours_worked:.2f}h al 100%.")
            elif dow == 6:
                exp_parts.append("Domingo ⇒ todo al 100%.")
            elif dow == 5:
                if intervals:
                    exp_parts.append("Sábado: antes de 13:00 horario normal; desde 13:00 y domingo ⇒ 100%.")
                else:
                    exp_parts.append("Sábado sin marcajes (entries): se aplica horario normal.")
            else:
                if hours_worked <= self.jornada_completa:
                    if pending > 0:
                        exp_parts.append(
                            f"Lun–Vie: {hours_worked:.2f}h regulares + {pending:.2f}h pendientes."
                        )
                    else:
                        exp_parts.append("Lun–Vie: horas dentro de la jornada regular.")
                else:
                    extra = hours_worked - self.jornada_completa
                    e50_view = min(extra, float(self.extras_al_50))
                    e100_view = max(0.0, extra - float(self.extras_al_50))
                    exp_parts.append(
                        f"Lun–Vie: 8h regulares + {e50_view:.2f}h 50% + {e100_view:.2f}h 100%."
                    )

            calc_note = " ".join(exp_parts)

            # Fila diaria
            daily_data.append({
                'employee_id': employee_info.get('employeeInternalId'),
                'date': out_date_str,
                'day_of_week': self.get_day_of_week_spanish(datetime.strptime(out_date_str, '%Y-%m-%d')),
                'hours_worked': round(hours_worked, 2),
                'regular_hours': round(regular_hours, 2),
                'extra_hours_50': round(extra50, 2),
                'extra_hours_100': round(extra100, 2),
                'night_hours': round(night_hours, 2),
                'holiday_hours': round(holiday_hours, 2),               # NUEVO
                'pending_hours': round(pending, 2) if not (has_time_off or has_absence) else 0.0,
                'is_holiday': is_holiday_output,
                'holiday_name': holiday_name,
                'is_rest_day': bool(is_rest_day),                        # NUEVO
                'has_time_off': has_time_off,
                'time_off_name': (day_summary.get('timeOffRequests') or [{}])[0].get('name') if has_time_off else None,
                'has_absence': has_absence,
                'is_full_time': hours_worked >= self.jornada_completa,
                'shift_start': " ".join([disp_start_d, disp_start_h]).strip(),  # NUEVO
                'shift_end':   " ".join([disp_end_d,   disp_end_h]).strip(),    # NUEVO
                'turno': turno,
                'calc_note': calc_note,
            })

        # compensaciones (si las usás)
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

    # -------------------- Distribución estándar (compat / no usada en sábado especial) --------------------

    def calculate_hour_distribution(self, hours_worked: float, date: datetime,
                                    is_holiday: bool = False, has_time_off: bool = False,
                                    night_hours: float = 0.0) -> Dict:
        """
        Mantengo por compatibilidad, pero la rama de sábado ahora se maneja en process_employee_data
        con cortes por intervalo. Para Lun–Vie usa extras_al_50 de config.
        """
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
        if day_of_week == 6:
            return {
                'hours_worked': float(hours_worked),
                'regular_hours': 0.0,
                'extra_hours_50': 0.0,
                'extra_hours_100': float(hours_worked),
                'night_hours': float(night_hours),
                'pending_hours': 0.0
            }

        # Lun–Vie (y sábado simplificado antiguo eliminado)
        dist = self._weekday_distribution(float(hours_worked), has_time_off)
        return {
            'hours_worked': float(hours_worked),
            'regular_hours': float(dist['regular']),
            'extra_hours_50': float(dist['extra50']),
            'extra_hours_100': float(dist['extra100']),
            'night_hours': float(night_hours),
            'pending_hours': float(dist['pending'])
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


# Funciones de compatibilidad
def process_employee_data_from_day_summaries(day_summaries: List[Dict], employee_info: Dict,
                                             previous_pending_hours: float = 0,
                                             period_dates: Dict = None, holidays: Optional[Set[str]] = None) -> Dict:
    calc = ArgentineHoursCalculator()
    return calc.process_employee_data(day_summaries, employee_info, previous_pending_hours, holidays or set())


def calculate_compensations(extra_hours_50: float, extra_hours_100: float, pending_hours: float) -> Dict:
    calc = ArgentineHoursCalculator()
    return calc.calculate_compensations(extra_hours_50, extra_hours_100, pending_hours)
