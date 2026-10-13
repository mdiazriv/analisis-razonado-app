# -*- coding: utf-8 -*-
"""
Traducción dinámica ES→EN vía la API de DeepL — complemento, no reemplazo, de las traducciones
fijas que ya existen en render_tables_v5.py (TABLES_EN) y assemble_en.py (FDEC_LABELS_EN).

Por qué hace falta esto: las traducciones fijas son una foto congelada del texto de UN trimestre
específico (ej. la descripción de un riesgo), verificada a mano contra el Word en inglés de
referencia de ese trimestre. Eso funciona perfecto mientras el texto en español no cambie de un
trimestre a otro, pero se desactualiza en cuanto el contenido real cambia (redacción editada, un
riesgo nuevo que reemplaza a uno antiguo, etc.) — ahí la app seguía mostrando la traducción vieja,
sin relación con lo que dice el Excel/Word en español de ESTE trimestre. Esto es lo que Melissa
reportó como "el texto en inglés no está quedando exactamente igual al español", y pidió
explícitamente resolver instalando una API de traducción externa (DeepL), aceptando el trade-off de
que ese texto puntual sale de este computador hacia los servidores de DeepL para traducirse.

Diseño importante:
- Por defecto (sin API key configurada) esto NO hace nada: ni un solo byte sale de este
  computador, y todo sigue funcionando exactamente igual que antes (con las traducciones fijas
  como respaldo). La traducción dinámica es estrictamente opt-in.
- Nunca lanza una excepción que tumbe la generación del informe: cualquier falla de red, cuota
  agotada, timeout, o API caída devuelve None, y quien llama debe tener un plan B (una traducción
  fija de respaldo, o dejar el texto en español).
- Cachea cada texto traducido dentro de la misma corrida, para no gastar cuota repitiendo el mismo
  texto dos veces (ej. una etiqueta que se usa en más de una tabla).
"""
import requests

_API_KEY = None
_CACHE = {}  # (texto, source_lang, target_lang) -> traducción

# _LAST_ERROR: motivo de la ÚLTIMA llamada fallida a la API (p.ej. key inválida, cuota agotada,
# problema de red). Antes translate_es_en() tragaba CUALQUIER excepción en silencio y devolvía
# None sin dejar rastro, así que si la key estaba mal escrita o la cuenta no estaba activada, el
# único síntoma visible para Melissa era "SPANISH — DeepL not configured" en el inglés generado —
# el MISMO mensaje que cuando nunca se configuró ninguna key, lo que la llevó a pensar que su key
# no se había guardado bien cuando en realidad el problema era otro (cuenta no verificada, cuota
# agotada, etc.). Con esto, assemble_en.py puede mostrar el motivo real en el log de generación.
_LAST_ERROR = None


def configure_deepl(api_key):
    """Guarda la API key para esta corrida de generación (ver app.py, que la lee de
    st.secrets['DEEPL_API_KEY'] y llama a esto antes de generar el Word en inglés).
    api_key=None o '' desactiva la traducción dinámica (comportamiento por defecto)."""
    global _API_KEY, _LAST_ERROR
    _API_KEY = (api_key or '').strip() or None
    _CACHE.clear()
    _LAST_ERROR = None


def is_configured():
    return _API_KEY is not None


def get_last_error():
    """Motivo de la última llamada fallida a la API en esta corrida (None si no ha fallado
    ninguna, o si DeepL nunca estuvo configurado). Ver nota de _LAST_ERROR arriba."""
    return _LAST_ERROR


def translate_es_en(text, source_lang='ES', target_lang='EN-US'):
    """Traduce texto libre de español a inglés vía la API de DeepL. Devuelve None (nunca lanza
    excepción) si no hay api_key configurada o si la llamada falla por cualquier motivo — quien
    llama debe seguir funcionando sin traducción dinámica en ese caso. El motivo de una falla
    queda disponible vía get_last_error() para mostrarlo en el log (ver assemble_en.py)."""
    global _LAST_ERROR
    if text is None:
        return None
    text = text.strip()
    if not text:
        return text
    if _API_KEY is None:
        return None
    cache_key = (text, source_lang, target_lang)
    if cache_key in _CACHE:
        return _CACHE[cache_key]
    # Las cuentas gratuitas de DeepL usan un endpoint distinto al de las cuentas de pago; la propia
    # API key lo indica (las de cuenta gratuita siempre terminan en ":fx").
    base = 'https://api-free.deepl.com' if _API_KEY.endswith(':fx') else 'https://api.deepl.com'
    try:
        resp = requests.post(
            f'{base}/v2/translate',
            data={
                'auth_key': _API_KEY,
                'text': text,
                'source_lang': source_lang,
                'target_lang': target_lang,
            },
            timeout=15,
        )
        if resp.status_code == 403:
            _LAST_ERROR = 'API key inválida o no autorizada (error 403 de DeepL)'
            return None
        if resp.status_code == 456:
            _LAST_ERROR = 'Se agotó la cuota de caracteres de la cuenta DeepL (error 456)'
            return None
        if resp.status_code == 429:
            _LAST_ERROR = 'Demasiadas solicitudes a DeepL a la vez (error 429) — reintentar más tarde'
            return None
        resp.raise_for_status()
        translated = resp.json()['translations'][0]['text']
        _CACHE[cache_key] = translated
        return translated
    except requests.exceptions.Timeout:
        _LAST_ERROR = 'Tiempo de espera agotado contactando a DeepL (posible problema de red)'
        return None
    except requests.exceptions.ConnectionError:
        _LAST_ERROR = 'No se pudo conectar con DeepL (posible problema de red o firewall)'
        return None
    except Exception as e:
        _LAST_ERROR = f'Error inesperado llamando a DeepL: {e}'
        return None
