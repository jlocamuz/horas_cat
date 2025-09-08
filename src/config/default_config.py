"""
Configuración por defecto del sistema
Valores heredados del proyecto anterior Horas-cat-v3
"""

DEFAULT_CONFIG = {
    # API Configuration
    'api_key': 'NTgyNTM5NTpuYzhJSXFQNEUzeXZNcndpNzVCR3ZJYm4wTkJ2aWpXTg==',
    'base_url': 'https://api-prod.humand.co/public/api/v1',
    
    # Configuración de jornada laboral
    'jornada_completa_horas': 8,
    'tolerancia_minutos': 20,
    'fragmento_minutos': 30,
    
    # Horarios especiales
    'hora_nocturna_inicio': 21,  # 21:00
    'hora_nocturna_fin': 6,      # 06:00
    'sabado_limite_hora': 13,    # 13:00 - después de esta hora son horas 100%
    
    # Configuración de zona horaria
    'timezone': 'America/Argentina/Buenos_Aires',
    
    # Configuración de API
    'max_retries': 3,
    'retry_delay': 1000,  # milisegundos
    'request_timeout': 30000,  # 30 segundos
    
    # Configuración de procesamiento paralelo
    'max_workers': 6,
    'batch_size_users': 10,
    'batch_size_dates': 7,
    'delay_between_retries': 1000,
    'delay_between_batches': 500,
    
    # Configuración de archivos
    'output_directory': '~/Downloads',
    'filename_format': 'reporte_{start_date}_{end_date}.xlsx',
    
    # Configuración de interfaz
    'window_width': 800,
    'window_height': 600,
    'theme': 'default'
}

# Headers para las llamadas a la API
def get_api_headers(api_key=None):
    """Obtiene los headers para las llamadas a la API"""
    if api_key is None:
        api_key = DEFAULT_CONFIG['api_key']
    
    return {
        'Authorization': f'Basic {api_key}',
        'Content-Type': 'application/json'
    }

# Endpoints de la API
API_ENDPOINTS = {
    'users': '/users',
    'time_tracking_entries': '/time-tracking/entries',
    'day_summaries': '/time-tracking/day-summaries'
}
