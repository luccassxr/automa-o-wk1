import os
import sys
import json
import time
from collections import Counter
from datetime import datetime
from typing import Callable, Dict, List, Optional, Tuple

import pyautogui

from robo_cartoes_emsys_v3 import (
    CDP_URL,
    CONFIG_FILE,
    CAPTURES_DIR,
    VALE_DESP_FILE,
    GOODCARD_FALLBACK_URL,
    JS_EXTRACT_DATETIME_AND_BRUTO,
    ensure_dir,
    normalize_brl,
    normalize_dt,
    brl_to_float,
    float_to_brl,
    date_range_from_rows,
    save_capture_txt,
    read_all_captures,
    copy_current_row_text,
    extract_rs_original_from_row,
    extract_titulo_from_row,
    valecard_capture_from_pdf,
    valecard_somar_despesas_pdf,
    redefrota_capture_from_pdf,
)

import storage


def get_base_dir() -> str:
    """
    Retorna o diretório base do aplicativo (compatível com PyInstaller).
    Todos os arquivos (capturas, configs, etc.) são salvos aqui.
    """
    if getattr(sys, "frozen", False):
        # Executável gerado pelo PyInstaller: usa o diretório do .exe
        base_path = os.path.dirname(sys.executable)
    else:
        # Execução via Python normal: usa o diretório deste arquivo
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path


def chdir_to_base():
    """
    Garante que o diretório de trabalho atual seja o diretório do app.
    Assim, arquivos são sempre criados ao lado do .py/.exe.
    """
    base = get_base_dir()
    os.chdir(base)
    ensure_dir(CAPTURES_DIR)


def get_goodcard_fallback_url() -> str:
    """
    Retorna a URL de fallback configurada para o portal Good Card.
    Se não houver configuração, usa o valor padrão do core.
    """
    settings = storage.load_settings()
    url = settings.get("goodcard_fallback_url") or GOODCARD_FALLBACK_URL
    return url


def set_goodcard_fallback_url(url: str):
    settings = storage.load_settings()
    settings["goodcard_fallback_url"] = url.strip()
    storage.save_settings(settings)


def goodcard_check_cdp() -> Tuple[bool, Optional[str]]:
    """
    Verifica se o Chrome está acessível via CDP na porta 9222.
    Retorna (ok, mensagem_erro).
    """
    import urllib.request
    import urllib.error

    try:
        with urllib.request.urlopen(CDP_URL, timeout=2) as resp:
            # Se respondeu algo, consideramos OK
            _ = resp.read(1)
        return True, None
    except urllib.error.URLError as e:
        return False, str(e)
    except Exception as e:  # pragma: no cover - segurança extra
        return False, str(e)


def goodcard_list_tabs() -> List[Dict[str, str]]:
    """
    Lista abas abertas no Chrome acessível via CDP (Playwright).
    Retorna lista de dicts com title e url.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Playwright não está instalado. Rode: pip install playwright && python -m playwright install"
        )

    tabs: List[Dict[str, str]] = []

    p = sync_playwright().start()
    browser = None
    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        pages = context.pages

        for pg in pages:
            try:
                title = (pg.title() or "").strip()
            except Exception:
                title = ""
            try:
                url = (pg.url or "").strip()
            except Exception:
                url = ""
            if not title and not url:
                continue
            tabs.append({"title": title, "url": url})

    finally:
        try:
            if browser:
                browser.close()
        except Exception:
            pass
        p.stop()

    return tabs


def goodcard_open_portal_tab():
    """
    Abre uma nova aba do portal Good Card na URL de fallback configurada.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Playwright não está instalado. Rode: pip install playwright && python -m playwright install"
        )

    fallback_url = get_goodcard_fallback_url()

    p = sync_playwright().start()
    browser = None
    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.new_page()
        page.goto(fallback_url, wait_until="domcontentloaded")
    finally:
        try:
            if browser:
                browser.close()
        except Exception:
            pass
        p.stop()


def goodcard_capture_from_url(page_url: str) -> List[Dict[str, str]]:
    """
    Captura vendas do Good Card na aba com a URL indicada.
    Retorna lista de dicts {dt, bruto, id}.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError(
            "Playwright não está instalado. Rode: pip install playwright && python -m playwright install"
        )

    p = sync_playwright().start()
    browser = None
    try:
        browser = p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        pages = context.pages

        target_page = None
        for pg in pages:
            try:
                if (pg.url or "").strip() == page_url.strip():
                    target_page = pg
                    break
            except Exception:
                continue

        if target_page is None:
            raise RuntimeError("Não encontrei a aba selecionada. Atualize a lista de abas e tente novamente.")

        target_page.bring_to_front()
        target_page.wait_for_timeout(800)

        rows = []
        try:
            main_rows = target_page.evaluate(JS_EXTRACT_DATETIME_AND_BRUTO)
            if main_rows:
                rows.extend(main_rows)
        except Exception:
            pass

        if not rows:
            for fr in target_page.frames:
                try:
                    fr_rows = fr.evaluate(JS_EXTRACT_DATETIME_AND_BRUTO)
                    if fr_rows:
                        rows.extend(fr_rows)
                except Exception:
                    continue

        out = []
        seen = set()
        for r in rows:
            dt = normalize_dt(r.get("dt", ""))
            bruto = normalize_brl(r.get("bruto", ""))
            if not dt or not bruto:
                continue
            key = (dt, bruto)
            if key in seen:
                continue
            seen.add(key)
            out.append({"dt": dt, "bruto": bruto, "id": ""})

        return out

    finally:
        try:
            if browser:
                browser.close()
        except Exception:
            pass
        p.stop()


def summarize_unified_captures():
    """
    Lê todas as capturas e retorna um resumo:
    - items: lista unificada
    - total: quantidade
    - soma: valor bruto total
    - dmin/dmax: datas mais antiga/recente (datetime ou None)
    """
    items = read_all_captures()
    total = len(items)
    soma = sum(brl_to_float(i["bruto"]) for i in items)
    dmin, dmax = date_range_from_rows(items)
    return {
        "items": items,
        "total": total,
        "soma": soma,
        "dmin": dmin,
        "dmax": dmax,
    }


def load_valecard_despesas() -> Optional[Dict]:
    """
    Lê o arquivo de despesas do Vale Card, se existir.
    Retorna dict com campos *_abs já tratados, ou None.
    """
    if not os.path.exists(VALE_DESP_FILE):
        return None
    try:
        with open(VALE_DESP_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        return None

    # Garantir campos numéricos
    total_abs = float(d.get("total_despesas_abs", 0.0))
    taxa_abs = float(d.get("taxa_adm_abs", 0.0))
    outras_abs = float(d.get("outras_abs", 0.0))
    d["total_despesas_abs"] = abs(total_abs)
    d["taxa_adm_abs"] = abs(taxa_abs)
    d["outras_abs"] = abs(outras_abs)
    return d


def clear_captures() -> int:
    """
    Remove todos os arquivos captura_*.txt em CAPTURES_DIR.
    Retorna a quantidade de arquivos removidos.
    """
    ensure_dir(CAPTURES_DIR)
    removed = 0
    for fn in os.listdir(CAPTURES_DIR):
        if fn.startswith("captura_") and fn.endswith(".txt"):
            try:
                os.remove(os.path.join(CAPTURES_DIR, fn))
                removed += 1
            except Exception:
                pass
    return removed


def export_unified_to_csv(csv_path: str) -> Dict[str, int]:
    """
    Exporta as capturas unificadas para CSV simples.
    """
    import csv

    summary = summarize_unified_captures()
    items = summary["items"]

    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(["data_hora", "valor_bruto", "origem", "id_opcional"])
        for it in items:
            writer.writerow([it["dt"], it["bruto"], it.get("origem", ""), it.get("id", "")])

    return {"total": len(items)}


def save_emsys_config_from_gui(grid_cell: Dict[str, int]):
    """
    Salva o arquivo de configuração do EMSYS usando o ponto capturado via GUI.
    Mantém a mesma estrutura usada pelo script original.
    """
    cfg = {
        "grid_cell": {"x": int(grid_cell["x"]), "y": int(grid_cell["y"])},
        "max_steps": 25000,
        "same_row_limit": 25,
        "delay_apos_copiar": 0.15,
        "delay_entre_linhas": 0.06,
    }
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


ProgressCallback = Callable[[Dict], None]


def run_emsys_marking_with_progress(
    unified_rows: List[Dict[str, str]],
    progress_cb: Optional[ProgressCallback] = None,
    cancel_event=None,
):
    """
    Versão de run_emsys_marking com callback de progresso para a GUI.
    Mantém a mesma lógica do core, mas em vez de depender só de prints,
    envia eventos estruturados para a callback.
    """

    def emit(event_type: str, **data):
        if progress_cb:
            payload = {"type": event_type}
            payload.update(data)
            try:
                progress_cb(payload)
            except Exception:
                # Não deixar a automação quebrar por causa da GUI
                pass

    if not os.path.exists(CONFIG_FILE):
        emit("error", message=f"Não achei {CONFIG_FILE}. Rode a calibração primeiro.")
        return

    try:
        cfg = json.load(open(CONFIG_FILE, "r", encoding="utf-8"))
    except Exception as e:
        emit("error", message=f"Erro ao ler {CONFIG_FILE}: {e}")
        return

    delay_apos_copiar = float(cfg.get("delay_apos_copiar", 0.15))
    delay_entre_linhas = float(cfg.get("delay_entre_linhas", 0.06))

    portal_values = [r["bruto"] for r in unified_rows]
    target_counts = Counter(portal_values)
    total_portal = sum(target_counts.values())

    emit(
        "start",
        total_portal=total_portal,
        valores_unicos=len(target_counts),
    )

    time.sleep(1.0)

    pyautogui.FAILSAFE = True

    try:
        # Clique inicial no grid
        grid_cell = cfg["grid_cell"]
        pyautogui.click(grid_cell["x"], grid_cell["y"])
        time.sleep(0.2)
    except Exception as e:
        emit("error", message=f"Erro ao clicar no grid do EMSYS: {e}")
        return

    found: List[str] = []
    last_row_text = None
    same_row_count = 0
    last_titulo = None
    same_titulo_count = 0
    same_row_limit = int(cfg.get("same_row_limit", 25))

    try:
        for _ in range(int(cfg.get("max_steps", 25000))):
            # Permite cancelamento gracioso a partir da GUI
            if cancel_event is not None and getattr(cancel_event, "is_set", lambda: False)():
                emit("log", message="Marcação interrompida pelo usuário (botão Parar).")
                break

            if len(found) >= total_portal:
                emit(
                    "log",
                    message="Todas as vendas do portal foram marcadas no EMSYS. Encerrando.",
                )
                break

            row = copy_current_row_text()
            time.sleep(delay_apos_copiar)
            row_norm = (row or "").strip()

            if last_row_text is not None and row_norm == last_row_text:
                same_row_count += 1
            else:
                same_row_count = 0
                last_row_text = row_norm

            titulo = extract_titulo_from_row(row)
            if titulo and last_titulo is not None and titulo == last_titulo:
                same_titulo_count += 1
            else:
                same_titulo_count = 0
                if titulo:
                    last_titulo = titulo

            if same_row_count >= same_row_limit or same_titulo_count >= same_row_limit:
                emit("log", message="Cheguei ao final do grid. Encerrando.")
                break

            rs_original = extract_rs_original_from_row(row)

            if rs_original and target_counts.get(rs_original, 0) > 0:
                pyautogui.press("enter")
                time.sleep(0.08)
                target_counts[rs_original] -= 1
                found.append(rs_original)
                emit(
                    "progress",
                    marcado=len(found),
                    total=total_portal,
                    valor=rs_original,
                )
                continue

            pyautogui.press("down")
            time.sleep(delay_entre_linhas)

    except pyautogui.FailSafeException:
        emit("log", message="Automação interrompida pelo FAILSAFE do mouse (canto superior esquerdo).")
    except Exception as e:
        emit("error", message=f"Erro durante a marcação: {e}")

    missing: List[str] = []
    for val, cnt in target_counts.items():
        if cnt > 0:
            missing.extend([val] * cnt)

    with open("encontrados.txt", "w", encoding="utf-8") as f:
        for v in found:
            f.write(v + "\n")

    with open("nao_encontrados.txt", "w", encoding="utf-8") as f:
        for v in missing:
            f.write(v + "\n")

    soma_encontrados = sum(brl_to_float(v) for v in found)
    soma_nao_encontrados = sum(brl_to_float(v) for v in missing)

    with open("resumo.txt", "w", encoding="utf-8") as f:
        f.write("Resumo Portal x EMSYS\n")
        f.write("---------------------\n")
        f.write(f"Total portal (unificado): {total_portal}\n")
        f.write(f"Marcados EMSYS: {len(found)}\n")
        f.write(f"Não encontrados: {len(missing)}\n\n")
        f.write(f"Soma marcados: R$ {float_to_brl(soma_encontrados)}\n")
        f.write(f"Soma não encontrados: R$ {float_to_brl(soma_nao_encontrados)}\n\n")
        f.write(f"Pasta de capturas: {CAPTURES_DIR}\\\n")

        if os.path.exists(VALE_DESP_FILE):
            try:
                d = json.load(open(VALE_DESP_FILE, "r", encoding="utf-8"))
                f.write("\nVale Card - Despesas (do último PDF lido)\n")
                f.write(f"Total despesas: R$ {float_to_brl(d.get('total_despesas_abs', 0.0))}\n")
                f.write(f"Taxa administrativa: R$ {float_to_brl(d.get('taxa_adm_abs', 0.0))}\n")
                f.write(f"Outras despesas: R$ {float_to_brl(d.get('outras_abs', 0.0))}\n")
            except Exception:
                pass

    emit(
        "end",
        total_portal=total_portal,
        marcados=len(found),
        nao_encontrados=len(missing),
        soma_marcados=soma_encontrados,
        soma_nao_encontrados=soma_nao_encontrados,
    )

