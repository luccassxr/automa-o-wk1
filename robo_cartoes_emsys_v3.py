# robo_cartoes_emsys.py
# Automação: capturar vendas (Good Card via Chrome/CDP, Vale Card via PDF, Rede Frota via PDF)
# -> salvar em capturas_portal -> unificar -> marcar no EMSYS (grid) por valor bruto
#
# Requisitos (uma vez):
#   pip install pyautogui pyperclip openpyxl pdfplumber playwright
#   python -m playwright install
#
# Para Good Card (CDP): abrir Chrome com:
#   "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir=C:\chrome-automacao
#
# O que este script faz:
# - Captura valores brutos + data/hora (para deduplicar corretamente) e salva em capturas_portal\captura_XXX.txt
# - Unifica todas as capturas e mostra intervalo de datas (mais velha e mais recente)
# - No EMSYS: varre o grid, copia a linha (Ctrl+C), lê o "R$ Original" e marca (Enter) quando encontrar valores do portal.
# - Para automaticamente quando:
#     (a) chega no fim do grid (mesma linha repetindo), OU
#     (b) já marcou TODAS as vendas do portal (novo!)
# - Vale Card: além de capturar vendas (positivas) para EMSYS, soma despesas (valores negativos) e separa Taxa Administrativa.
#
# Ajuste opcional:
# - GOODCARD_FALLBACK_URL: se não achar a guia do Good Card, o script abre uma nova aba nesse endereço.

import os
import re
import json
import time
import traceback
from datetime import datetime
from collections import Counter

import pyautogui
import pyperclip

# =====================
# CONFIG
# =====================
CDP_URL = "http://127.0.0.1:9222"
CONFIG_FILE = "config_emsys_grid.json"
CAPTURES_DIR = "capturas_portal"
VALE_DESP_FILE = "valecard_despesas.json"

GOODCARD_FALLBACK_URL = "about:blank"

pyautogui.FAILSAFE = True

BRL_NUM_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{1,2}")
BRL_SIGNED_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{1,2}")
TITULO_RE = re.compile(r"\b\d+/\d+\b")

# =====================
# Helpers: BRL / Date
# =====================
def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def normalize_brl(s: str) -> str:
    s = (s or "").strip().replace("R$", "").strip()
    m = BRL_NUM_RE.search(s)
    if not m:
        return ""
    num = m.group(0)
    if re.match(r".*,\d$", num):
        num += "0"
    return num

def brl_to_float(brl_num: str) -> float:
    brl_num = (brl_num or "").strip()
    if not brl_num:
        return 0.0
    if re.match(r".*,\d$", brl_num):
        brl_num += "0"
    return float(brl_num.replace(".", "").replace(",", "."))

def brl_to_float_signed(s: str) -> float:
    s = (s or "").strip().replace("R$", "").strip()
    m = BRL_SIGNED_RE.search(s)
    if not m:
        return 0.0
    num = m.group(0)
    if re.match(r".*,\d$", num):
        num += "0"
    return float(num.replace(".", "").replace(",", "."))

def float_to_brl(v: float) -> str:
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

def normalize_dt(dt_str: str) -> str:
    dt_str = (dt_str or "").strip()
    if not dt_str:
        return ""
    if re.match(r"^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}$", dt_str):
        dt_str += ":00"
    if re.match(r"^\d{2}/\d{2}/\d{4}$", dt_str):
        dt_str += " 00:00:00"
    try:
        datetime.strptime(dt_str, "%d/%m/%Y %H:%M:%S")
        return dt_str
    except:
        return ""

def dt_to_obj(dt_str: str):
    dt_str = normalize_dt(dt_str)
    if not dt_str:
        return None
    try:
        return datetime.strptime(dt_str, "%d/%m/%Y %H:%M:%S")
    except:
        return None

def date_range_from_rows(rows):
    dts = [dt_to_obj(r.get("dt","")) for r in rows]
    dts = [d for d in dts if d is not None]
    if not dts:
        return (None, None)
    return (min(dts), max(dts))

def next_capture_filename() -> str:
    ensure_dir(CAPTURES_DIR)
    existing = []
    for fn in os.listdir(CAPTURES_DIR):
        if fn.startswith("captura_") and fn.endswith(".txt"):
            try:
                n = int(fn.replace("captura_", "").replace(".txt", ""))
                existing.append(n)
            except:
                pass
    nxt = (max(existing) + 1) if existing else 1
    return os.path.join(CAPTURES_DIR, f"captura_{nxt:03d}.txt")

def save_capture_txt(rows, origem: str):
    fn = next_capture_filename()
    with open(fn, "w", encoding="utf-8") as f:
        f.write("data_hora;valor_bruto;origem;id_opcional\n")
        for r in rows:
            dt = normalize_dt(r.get("dt", ""))
            bruto = normalize_brl(r.get("bruto", ""))
            if not dt or not bruto:
                continue
            id_opt = str(r.get("id", "") or "").strip()
            f.write(f"{dt};{bruto};{origem};{id_opt}\n")
    return fn

def read_all_captures():
    ensure_dir(CAPTURES_DIR)
    items = []
    for fn in sorted(os.listdir(CAPTURES_DIR)):
        if not (fn.startswith("captura_") and fn.endswith(".txt")):
            continue
        path = os.path.join(CAPTURES_DIR, fn)
        with open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.lower().startswith("data_hora"):
                    continue
                parts = line.split(";")
                if len(parts) < 3:
                    continue
                dt = normalize_dt(parts[0].strip())
                bruto = normalize_brl(parts[1].strip())
                origem = parts[2].strip()
                id_opt = parts[3].strip() if len(parts) >= 4 else ""
                if dt and bruto:
                    items.append({"dt": dt, "bruto": bruto, "origem": origem, "id": id_opt})

    seen = set()
    out = []
    for it in items:
        key = (it["dt"], it["bruto"], it.get("origem", ""), it.get("id", ""))
        if key not in seen:
            seen.add(key)
            out.append(it)
    return out

# =====================
# EMSYS helpers
# =====================
def capture_point(name):
    print(f"\n👉 Posicione o mouse em: {name}")
    input("   Quando estiver em cima, pressione ENTER aqui...")
    x, y = pyautogui.position()
    print(f"   OK: {name} = ({x}, {y})")
    return {"x": x, "y": y}

def click(p):
    pyautogui.click(p["x"], p["y"])

def copy_current_row_text() -> str:
    pyperclip.copy("")
    pyautogui.hotkey("ctrl", "c")
    time.sleep(0.15)
    return pyperclip.paste()

def extract_rs_original_from_row(row_text: str) -> str:
    if not row_text:
        return ""
    row_text = row_text.replace("\r", "")
    if "\n" in row_text:
        lines = [l for l in row_text.split("\n") if l.strip()]
        row_line = lines[-1] if lines else ""
    else:
        row_line = row_text

    parts = row_line.split("\t")
    if len(parts) >= 8:
        return normalize_brl(parts[6])

    nums = BRL_NUM_RE.findall(row_line)
    return normalize_brl(nums[0]) if nums else ""

def extract_titulo_from_row(row_text: str) -> str:
    if not row_text:
        return ""
    row_text = row_text.replace("\r", "")
    if "\n" in row_text:
        lines = [l for l in row_text.split("\n") if l.strip()]
        row_line = lines[-1] if lines else ""
    else:
        row_line = row_text
    m = TITULO_RE.search(row_line)
    return m.group(0) if m else ""

def calibrate_emsys():
    print("=== CALIBRAÇÃO EMSYS (GRID) ===")
    print("1) Abra o EMSYS na tela 'Geração de Fatura'")
    print("2) Clique em uma linha do GRID para selecioná-la\n")

    cfg = {}
    cfg["grid_cell"] = capture_point("UMA CÉLULA QUALQUER do GRID (na linha selecionada)")
    cfg["max_steps"] = 25000
    cfg["same_row_limit"] = 25
    cfg["delay_apos_copiar"] = 0.15
    cfg["delay_entre_linhas"] = 0.06

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Calibração salva em {CONFIG_FILE}")

def run_emsys_marking(unified_rows):
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ Não achei {CONFIG_FILE}. Rode a calibração primeiro.")
        return

    cfg = json.load(open(CONFIG_FILE, "r", encoding="utf-8"))
    delay_apos_copiar = float(cfg.get("delay_apos_copiar", 0.15))
    delay_entre_linhas = float(cfg.get("delay_entre_linhas", 0.06))

    portal_values = [r["bruto"] for r in unified_rows]
    target_counts = Counter(portal_values)
    total_portal = sum(target_counts.values())

    print("\n=== EMSYS: MARCAÇÃO ===")
    print(f"Total transações unificadas: {total_portal}")
    print(f"Valores únicos: {len(target_counts)}")
    print("⚠️ Não mexa no mouse/teclado durante a execução.")
    print("FAILSAFE: mova o mouse pro canto superior esquerdo para parar.\n")
    time.sleep(1.0)

    click(cfg["grid_cell"])
    time.sleep(0.2)

    found = []
    last_row_text = None
    same_row_count = 0
    last_titulo = None
    same_titulo_count = 0
    same_row_limit = int(cfg.get("same_row_limit", 25))

    for _ in range(int(cfg.get("max_steps", 25000))):
        # ✅ novo: se já marcou tudo, para
        if len(found) >= total_portal:
            print("\n✅ Todas as vendas do portal foram marcadas no EMSYS. Encerrando.")
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
            print("\n🔚 Cheguei ao final do grid. Encerrando.")
            break

        rs_original = extract_rs_original_from_row(row)

        if rs_original and target_counts.get(rs_original, 0) > 0:
            pyautogui.press("enter")
            time.sleep(0.08)
            target_counts[rs_original] -= 1
            found.append(rs_original)
            print(f"✅ Marcado: {rs_original} ({len(found)}/{total_portal})")
            continue

        pyautogui.press("down")
        time.sleep(delay_entre_linhas)

    missing = []
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
            except:
                pass

    print("\n✅ Finalizado.")
    print(f"- encontrados.txt: {len(found)} | soma: R$ {float_to_brl(soma_encontrados)}")
    print(f"- nao_encontrados.txt: {len(missing)} | soma: R$ {float_to_brl(soma_nao_encontrados)}")
    print("- resumo.txt gerado.")

# =====================
# GOOD CARD: Chrome/CDP
# =====================
JS_EXTRACT_DATETIME_AND_BRUTO = r"""
() => {
  const headerDatetime = "Data / Hora";
  const headerBruto = "Valor Bruto da Transação";

  const rsRegex = /R\$\s*\d{1,3}(?:\.\d{3})*,\d{1,2}/;
  const dtRegex = /^\d{2}\/\d{2}\/\d{4}(\s+\d{2}:\d{2}:\d{2})?$/;

  function norm(s) { return (s || "").replace(/\s+/g, " ").trim().toLowerCase(); }
  function getText(el) { return (el && (el.innerText || el.textContent) || "").trim(); }

  function findHeaderIndex(table, headerName) {
    const target = norm(headerName);
    const headerRows = Array.from(table.querySelectorAll("thead tr"));
    if (!headerRows.length) return -1;

    for (const tr of headerRows) {
      const cells = Array.from(tr.querySelectorAll("th, td"));
      const texts = cells.map(c => norm(getText(c)));
      const idx = texts.findIndex(t => t.includes(target));
      if (idx !== -1) return idx;
    }
    return -1;
  }

  function safeGetCellText(tds, idx) {
    if (idx < 0 || idx >= tds.length) return "";
    return getText(tds[idx]);
  }

  const out = [];
  const tables = Array.from(document.querySelectorAll("table"));

  for (const table of tables) {
    const dtIdx = findHeaderIndex(table, headerDatetime);
    const brutoIdx0 = findHeaderIndex(table, headerBruto);
    if (dtIdx === -1 || brutoIdx0 === -1) continue;

    const rows = Array.from(table.querySelectorAll("tbody tr"));
    if (!rows.length) continue;

    for (const tr of rows) {
      const tds = Array.from(tr.querySelectorAll("td"));
      if (tds.length <= Math.max(dtIdx, brutoIdx0)) continue;

      let dt = safeGetCellText(tds, dtIdx);
      let bruto = safeGetCellText(tds, brutoIdx0);

      if (!rsRegex.test(bruto)) {
        for (let delta = 1; delta <= 4; delta++) {
          const cand1 = safeGetCellText(tds, brutoIdx0 + delta);
          if (rsRegex.test(cand1)) { bruto = cand1; break; }
          const cand2 = safeGetCellText(tds, brutoIdx0 - delta);
          if (rsRegex.test(cand2)) { bruto = cand2; break; }
        }
      }

      if (!dtRegex.test(dt)) continue;
      if (!rsRegex.test(bruto)) continue;

      bruto = bruto.replace("R$", "").trim();
      const m = bruto.match(/\d{1,3}(?:\.\d{3})*,\d{1,2}/);
      if (!m) continue;
      let val = m[0];
      if (/,\d$/.test(val)) val = val + "0";
      out.push({ dt, bruto: val });
    }
  }

  return out;
}
"""

def goodcard_capture_via_cdp():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("❌ Falta instalar Playwright. Rode: pip install playwright && python -m playwright install")
        return []

    p = sync_playwright().start()
    try:
        try:
            browser = p.chromium.connect_over_cdp(CDP_URL)
        except Exception:
            print("❌ Não consegui conectar ao Chrome via CDP.")
            print("Abra o Chrome com --remote-debugging-port=9222 e tente novamente.")
            return []

        context = browser.contexts[0]
        pages = context.pages

        auto_idx = None
        for i, pg in enumerate(pages):
            try:
                t = (pg.title() or "").lower()
                u = (pg.url or "").lower()
                if ("good" in t and "card" in t) or ("good" in u and "card" in u):
                    auto_idx = i
                    break
            except:
                continue

        if auto_idx is None:
            try:
                newp = context.new_page()
                newp.goto(GOODCARD_FALLBACK_URL, wait_until="domcontentloaded")
                pages = context.pages
                auto_idx = len(pages) - 1
                print("\n🟦 Não achei uma guia do Good Card. Abri uma nova aba automaticamente.")
                print("   Faça login/navegue até a tabela e depois escolha essa aba.\n")
            except:
                pass

        print("\nAbas encontradas no Chrome (CDP):")
        for i, pg in enumerate(pages):
            try:
                print(f"{i}: {pg.title()} | {pg.url}")
            except:
                print(f"{i}: (sem título) | {pg.url}")

        prompt = "\nDigite o número da aba do portal Good Card e pressione ENTER"
        if auto_idx is not None:
            prompt += f" (sugestão: {auto_idx})"
        prompt += ": "
        idx = input(prompt).strip()
        if not idx.isdigit():
            print("Número inválido.")
            return []
        page = pages[int(idx)]
        page.bring_to_front()
        page.wait_for_timeout(800)

        rows = []
        try:
            main_rows = page.evaluate(JS_EXTRACT_DATETIME_AND_BRUTO)
            if main_rows:
                rows.extend(main_rows)
        except:
            pass

        if not rows:
            for fr in page.frames:
                try:
                    fr_rows = fr.evaluate(JS_EXTRACT_DATETIME_AND_BRUTO)
                    if fr_rows:
                        rows.extend(fr_rows)
                except:
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
            browser.close()
        except:
            pass
        p.stop()

# =====================
# VALE CARD: PDF + DESPESAS
# =====================
def valecard_capture_from_pdf(pdf_path: str):
    try:
        import pdfplumber
    except ImportError:
        print("❌ Falta instalar pdfplumber. Rode: pip install pdfplumber")
        return []

    rows = []
    line_rx = re.compile(
        r"(?P<data>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<tipo>[VT])\s+"
        r"(?P<cod>\d{4,})\s+"
        r".*?"
        r"(?P<valor>\d{1,3}(?:\.\d{3})*,\d{1,2})"
        r"$"
    )

    any_text = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                any_text = True

            tables = []
            try:
                tables = page.extract_tables() or []
            except:
                tables = []

            got_any = False

            for tb in tables:
                if not tb or len(tb) < 2:
                    continue
                for row in tb[1:]:
                    if not row:
                        continue
                    row_join = " ".join([str(c or "").strip() for c in row if str(c or "").strip()])
                    m = line_rx.search(row_join)
                    if not m:
                        continue
                    data = m.group("data")
                    tipo = m.group("tipo")
                    cod = m.group("cod")
                    valor = m.group("valor")
                    if tipo != "V":
                        continue
                    bruto = normalize_brl(valor)
                    if not bruto:
                        continue
                    if brl_to_float(bruto) < 0:
                        continue
                    dt = normalize_dt(data)
                    rows.append({"dt": dt, "bruto": bruto, "id": cod})
                    got_any = True

            if not got_any and text.strip():
                for raw_line in text.splitlines():
                    line = raw_line.strip()
                    if not line:
                        continue
                    m = line_rx.search(line)
                    if not m:
                        continue
                    data = m.group("data")
                    tipo = m.group("tipo")
                    cod = m.group("cod")
                    valor = m.group("valor")
                    if tipo != "V":
                        continue
                    bruto = normalize_brl(valor)
                    if not bruto:
                        continue
                    if brl_to_float(bruto) < 0:
                        continue
                    dt = normalize_dt(data)
                    rows.append({"dt": dt, "bruto": bruto, "id": cod})

    if not any_text and not rows:
        print("❌ Esse PDF parece ser escaneado (imagem). Não dá pra ler sem OCR.")
        return []

    out = []
    seen = set()
    for r in rows:
        key = (r["dt"], r["bruto"], r.get("id", ""))
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out

def valecard_somar_despesas_pdf(pdf_path: str) -> dict:
    try:
        import pdfplumber
    except ImportError:
        print("❌ Falta instalar pdfplumber. Rode: pip install pdfplumber")
        return {"total_despesas": 0.0, "total_taxa_adm": 0.0, "total_outras": 0.0}

    total_despesas = 0.0
    total_taxa_adm = 0.0
    total_outras = 0.0

    def is_taxa_adm(line: str) -> bool:
        t = (line or "").lower()
        return ("taxa" in t) and (("adm" in t) or ("administr" in t))

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if not text.strip():
                continue
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue

                # ⛔ Ignora linhas de TOTAL / SUBTOTAL do rodapé (ex.: "Total Taxa Administração")
                low = line.lower()
                if re.search(r"\bsub[-\s]?total\b", low) or re.search(r"\btotal\b", low) or "valor total" in low:
                    continue

                vals = BRL_SIGNED_RE.findall(line.replace("R$", ""))
                if not vals:
                    continue
                v = brl_to_float_signed(vals[-1])
                if v < 0:
                    total_despesas += v
                    if is_taxa_adm(line):
                        total_taxa_adm += v
                    else:
                        total_outras += v

    return {"total_despesas": total_despesas, "total_taxa_adm": total_taxa_adm, "total_outras": total_outras}

# =====================
# REDE FROTA: PDF
# =====================
def redefrota_capture_from_pdf(pdf_path: str):
    try:
        import pdfplumber
    except ImportError:
        print("❌ Falta instalar pdfplumber. Rode: pip install pdfplumber")
        return []

    line_rx = re.compile(
        r"(?P<id>\d{6,})\s+"
        r"(?P<desc>\S+)\s+"
        r"(?P<data>\d{2}/\d{2}/\d{4})\s+"
        r"(?P<hora>\d{2}:\d{2}:\d{2})\s+"
        r"(?P<bruto>\d{1,3}(?:\.\d{3})*,\d{1,2})"
    )

    rows = []
    any_text = False

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                any_text = True
            if not text.strip():
                continue

            in_resumo = False
            for raw in text.splitlines():
                line = raw.strip()
                if not line:
                    continue
                if line.upper() == "RESUMO":
                    in_resumo = True
                    continue
                if not in_resumo:
                    continue
                m = line_rx.search(line)
                if not m:
                    continue
                tid = m.group("id").strip()
                data = m.group("data").strip()
                hora = m.group("hora").strip()
                bruto = m.group("bruto").strip()

                dt = normalize_dt(f"{data} {hora}")
                bruto_n = normalize_brl(bruto)
                if not dt or not bruto_n:
                    continue
                rows.append({"dt": dt, "bruto": bruto_n, "id": tid})

    if not any_text and not rows:
        print("❌ Esse PDF parece ser escaneado (imagem). Não dá pra ler sem OCR.")
        return []

    out = []
    seen = set()
    for r in rows:
        key = (r["dt"], r["bruto"], r.get("id", ""))
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    return out

# =====================
# Menu
# =====================
def print_header():
    print("\n======================================")
    print(" Automação Cartões -> Captura -> EMSYS ")
    print("======================================\n")

def explain_chrome_cdp():
    print("\nPara usar Good Card (CDP), abra o Chrome assim:")
    print(r'"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir=C:\chrome-automacao')
    print("Depois faça login no portal e deixe a tabela visível.\n")

def print_capture_range(rows, label="Captura"):
    dmin, dmax = date_range_from_rows(rows)
    if not dmin or not dmax:
        print(f"📅 {label}: não consegui calcular intervalo de datas.")
        return
    print(f"📅 {label} - mais velha: {dmin.strftime('%d/%m/%Y %H:%M:%S')} | mais recente: {dmax.strftime('%d/%m/%Y %H:%M:%S')}")

def show_vale_desp_if_exists():
    if not os.path.exists(VALE_DESP_FILE):
        return
    try:
        d = json.load(open(VALE_DESP_FILE, "r", encoding="utf-8"))
        print("\n📌 Vale Card - Despesas (último PDF lido)")
        print(f"Arquivo: {d.get('arquivo','')}")
        print(f"Atualizado em: {d.get('atualizado_em','')}")
        print(f"Total despesas: R$ {float_to_brl(float(d.get('total_despesas_abs',0.0)))}")
        print(f"Taxa administrativa: R$ {float_to_brl(float(d.get('taxa_adm_abs',0.0)))}")
        print(f"Outras despesas: R$ {float_to_brl(float(d.get('outras_abs',0.0)))}")
    except:
        pass

def menu_capturar():
    print("\n=== CAPTURAR ===")
    print("1) Good Card (Chrome/CDP)")
    print("2) Vale Card (PDF) + despesas")
    print("3) Rede Frota (PDF)")
    print("4) Voltar")
    op = input("Escolha: ").strip()

    if op == "1":
        explain_chrome_cdp()
        input("Se necessário, deixe o portal aberto e pressione ENTER...")
        rows = goodcard_capture_via_cdp()
        if not rows:
            print("❌ Não consegui capturar nada do Good Card.")
            return
        fn = save_capture_txt(rows, "GoodCard")
        print(f"✅ Capturado {len(rows)} transações. Salvo em: {fn}")
        print_capture_range(rows, "Good Card")

    elif op == "2":
        pdf_path = input("Caminho do PDF Vale Card: ").strip().strip('"')
        if not os.path.exists(pdf_path):
            print("❌ Arquivo não existe.")
            return

        rows = valecard_capture_from_pdf(pdf_path)
        if rows:
            fn = save_capture_txt(rows, "ValeCard")
            print(f"✅ Capturado {len(rows)} VENDAS. Salvo em: {fn}")
            print_capture_range(rows, "Vale Card (vendas)")
        else:
            print("⚠️ Não consegui extrair VENDAS (positivos). Vou calcular despesas mesmo assim.")

        desp = valecard_somar_despesas_pdf(pdf_path)
        total_abs = abs(desp["total_despesas"])
        taxa_abs = abs(desp["total_taxa_adm"])
        outras_abs = abs(desp["total_outras"])

        print("\n📌 VALE CARD - DESPESAS (valores negativos no PDF)")
        print(f"Total despesas: R$ {float_to_brl(total_abs)}")
        print(f"Taxa administrativa: R$ {float_to_brl(taxa_abs)}")
        print(f"Outras despesas: R$ {float_to_brl(outras_abs)}")

        with open(VALE_DESP_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "total_despesas_abs": total_abs,
                "taxa_adm_abs": taxa_abs,
                "outras_abs": outras_abs,
                "arquivo": pdf_path,
                "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            }, f, ensure_ascii=False, indent=2)
        print(f"✅ {VALE_DESP_FILE} atualizado (para o resumo).")

    elif op == "3":
        pdf_path = input("Caminho do PDF Rede Frota: ").strip().strip('"')
        if not os.path.exists(pdf_path):
            print("❌ Arquivo não existe.")
            return
        rows = redefrota_capture_from_pdf(pdf_path)
        if not rows:
            print("❌ Não consegui extrair transações desse PDF Rede Frota.")
            return
        fn = save_capture_txt(rows, "RedeFrota")
        print(f"✅ Capturado {len(rows)} transações. Salvo em: {fn}")
        print_capture_range(rows, "Rede Frota")

    else:
        return

def menu_principal():
    ensure_dir(CAPTURES_DIR)
    while True:
        print_header()
        print("1) Calibrar EMSYS (primeira vez)")
        print("2) Capturar (escolher cartão)")
        print("3) Ver total unificado (ler capturas) + intervalo de datas + despesas Vale Card")
        print("4) Rodar EMSYS (marcar grid)")
        print("5) Limpar capturas")
        print("6) Sair")
        op = input("Escolha: ").strip()

        if op == "1":
            calibrate_emsys()

        elif op == "2":
            menu_capturar()

        elif op == "3":
            items = read_all_captures()
            total = len(items)
            soma = sum(brl_to_float(i["bruto"]) for i in items)
            print(f"\n📦 Capturas unificadas: {total}")
            print(f"💰 Soma bruta (apenas referência): R$ {float_to_brl(soma)}")
            print(f"📁 Pasta: {CAPTURES_DIR}\\")
            dmin, dmax = date_range_from_rows(items)
            if dmin and dmax:
                print(f"📅 Intervalo para filtrar no portal: {dmin.strftime('%d/%m/%Y %H:%M:%S')}  até  {dmax.strftime('%d/%m/%Y %H:%M:%S')}")
            else:
                print("📅 Intervalo: não consegui calcular (faltam datas nas capturas).")
            show_vale_desp_if_exists()
            input("\nENTER para voltar...")

        elif op == "4":
            items = read_all_captures()
            if not items:
                print("❌ Não há capturas. Capture pelo menos uma vez.")
                time.sleep(1)
                continue
            input("\nDeixe o EMSYS aberto em 'Geração de Fatura', clique numa linha do grid e pressione ENTER...")
            run_emsys_marking(items)
            input("\nENTER para voltar...")

        elif op == "5":
            confirm = input("Apagar TODOS os arquivos em capturas_portal? (s/N): ").strip().lower()
            if confirm == "s":
                for fn in os.listdir(CAPTURES_DIR):
                    if fn.startswith("captura_") and fn.endswith(".txt"):
                        try:
                            os.remove(os.path.join(CAPTURES_DIR, fn))
                        except:
                            pass
                print("✅ Capturas apagadas.")
                time.sleep(1)

        elif op == "6":
            print("Saindo...")
            break

        else:
            print("Opção inválida.")
            time.sleep(1)

if __name__ == "__main__":
    try:
        menu_principal()
    except KeyboardInterrupt:
        print("\nInterrompido.")
    except Exception:
        print("\n❌ Erro inesperado:")
        print(traceback.format_exc())
