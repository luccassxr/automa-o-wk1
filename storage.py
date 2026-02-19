import json
import os
from typing import Any, Dict


SETTINGS_FILE = "app_settings.json"

_DEFAULT_SETTINGS: Dict[str, Any] = {
    "goodcard_fallback_url": "",
}


def _get_base_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _get_settings_path() -> str:
    base = _get_base_dir()
    return os.path.join(base, SETTINGS_FILE)


def load_settings() -> Dict[str, Any]:
    """
    Lê o arquivo de configurações da aplicação.
    Se não existir, retorna um dicionário com valores padrão.
    """
    path = _get_settings_path()
    if not os.path.exists(path):
        return dict(_DEFAULT_SETTINGS)

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return dict(_DEFAULT_SETTINGS)

    out = dict(_DEFAULT_SETTINGS)
    out.update(data or {})
    return out


def save_settings(data: Dict[str, Any]) -> None:
    """
    Salva o dicionário de configurações em disco.
    """
    path = _get_settings_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        # Não quebrar a aplicação se não conseguir salvar; apenas falhar silenciosamente.
        pass

