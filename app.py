import os
import sys
import threading
import queue
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
from typing import Dict, Any, List, Optional

try:
    from ttkbootstrap import Style as BootstrapStyle  # type: ignore[import]

    HAS_BOOTSTRAP = True
except Exception:
    BootstrapStyle = None  # type: ignore[assignment]
    HAS_BOOTSTRAP = False

import core
import storage
from ui_components import create_card, setup_styles


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Automação de Cartões")

        # Garante diretório base
        core.chdir_to_base()

        # Ícone
        self._setup_icon()

        # Fila para logs/progresso (threads -> UI)
        self.event_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

        # Estado em memória
        self.goodcard_tabs: List[Dict[str, str]] = []
        self.goodcard_selected_url: Optional[str] = None
        self.unified_items: List[Dict[str, Any]] = []
        self.emsys_thread: Optional[threading.Thread] = None
        self._emsys_cancel_event: Optional[threading.Event] = None

        # Variáveis de status
        self.status_goodcard = tk.StringVar(value="Desconectado do Chrome (CDP).")
        self.status_goodcard_capture = tk.StringVar(value="Nenhuma captura Good Card ainda.")
        self.status_valecard = tk.StringVar(value="Nenhum PDF Vale Card processado.")
        self.status_redefrota = tk.StringVar(value="Nenhum PDF Rede Frota processado.")
        self.status_unificado = tk.StringVar(value="Capturas ainda não unificadas.")
        self.status_emsys_calibracao = tk.StringVar(value="Calibração não realizada.")
        self.status_emsys_exec = tk.StringVar(value="EMSYS aguardando execução.")

        self.goodcard_tabs_var = tk.StringVar(value="")
        self.emsys_progress_var = tk.StringVar(value="Marcado: 0/0")
        self.emsys_last_value_var = tk.StringVar(value="Último valor marcado: -")

        # Splash opcional
        self._show_splash_then_build_ui()

        # Agendar processamento de eventos de thread
        self.root.after(200, self._process_event_queue)

    # --------------------------------------------------------------------- UI base
    def _setup_icon(self):
        try:
            # Usa o diretório de trabalho atual (ajustado em core.chdir_to_base())
            base = os.getcwd()
            icon_path = os.path.join(base, "icon.ico")
            if not os.path.exists(icon_path):
                # Fallback: ícone gerado na pasta de assets do Cursor (ambiente de desenvolvimento)
                icon_path = r"C:\Users\Usuário\.cursor\projects\c-Users-Usu-rio-Desktop-teste-bot\assets\icon.ico"
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception:
            # Fallback silencioso se o ícone não existir ou falhar
            pass

    def _show_splash_then_build_ui(self):
        # Criar splash screen simples (opcional)
        splash = tk.Toplevel(self.root)
        splash.overrideredirect(True)

        # Centralizar splash
        w, h = 360, 220
        sw = splash.winfo_screenwidth()
        sh = splash.winfo_screenheight()
        x = int((sw - w) / 2)
        y = int((sh - h) / 2)
        splash.geometry(f"{w}x{h}+{x}+{y}")

        frame = ttk.Frame(splash, padding=20)
        frame.pack(expand=True, fill="both")

        # Tentar carregar icon.png para exibir
        icon_label = ttk.Label(frame)
        icon_label.pack(pady=(0, 10))
        try:
            from PIL import Image, ImageTk  # type: ignore[import]

            # Em build com PyInstaller, os arquivos devem estar na raiz (diretório atual)
            base_dir = os.getcwd()
            png_path = os.path.join(base_dir, "icon.png")
            if not os.path.exists(png_path):
                # Fallback: ícone gerado na pasta de assets do Cursor (ambiente de desenvolvimento)
                png_path = r"C:\Users\Usuário\.cursor\projects\c-Users-Usu-rio-Desktop-teste-bot\assets\icon.png"
            if os.path.exists(png_path):
                img = Image.open(png_path)
                img = img.resize((96, 96), Image.LANCZOS)
                self._splash_img = ImageTk.PhotoImage(img)
                icon_label.configure(image=self._splash_img)
        except Exception:
            icon_label.configure(text="LOGO", font=("Segoe UI", 24, "bold"))  # apenas visual

        ttk.Label(frame, text="Automação de Cartões", font=("Segoe UI", 14, "bold")).pack(pady=(0, 4))
        ttk.Label(frame, text="by: lucas", font=("Segoe UI", 9, "italic")).pack(pady=(0, 12))
        ttk.Label(
            frame,
            text="Carregando...",
            font=("Segoe UI", 10),
        ).pack()

        self.root.withdraw()

        def finish_splash():
            try:
                splash.destroy()
            except Exception:
                pass
            self._build_main_ui()
            self.root.deiconify()

        self.root.after(2200, finish_splash)

    def _build_main_ui(self):
        setup_styles(self.root)

        container = ttk.Frame(self.root, padding=12)
        container.pack(expand=True, fill="both")

        # Notebook (abas)
        notebook = ttk.Notebook(container)
        notebook.pack(expand=True, fill="both")

        self.tab_inicio = ttk.Frame(notebook)
        self.tab_captura = ttk.Frame(notebook)
        self.tab_emsys = ttk.Frame(notebook)
        self.tab_relatorios = ttk.Frame(notebook)
        self.tab_config = ttk.Frame(notebook)

        notebook.add(self.tab_inicio, text="Início")
        notebook.add(self.tab_captura, text="Captura")
        notebook.add(self.tab_emsys, text="EMSYS")
        notebook.add(self.tab_relatorios, text="Relatórios")
        notebook.add(self.tab_config, text="Configuração/Ajuda")

        # Construir conteúdo de cada aba
        self._build_tab_inicio(notebook)
        self._build_tab_captura()
        self._build_tab_emsys()
        self._build_tab_relatorios()
        self._build_tab_config()

    # --------------------------------------------------------------------- Abas
    def _build_tab_inicio(self, notebook: ttk.Notebook):
        frame = self.tab_inicio
        frame.columnconfigure(0, weight=1)

        titulo = ttk.Label(
            frame,
            text="Automação de Cartões → Captura → EMSYS",
            font=("Segoe UI", 14, "bold"),
        )
        titulo.grid(row=0, column=0, sticky="w", pady=(4, 12))

        cards_container = ttk.Frame(frame)
        cards_container.grid(row=1, column=0, sticky="nsew")
        cards_container.columnconfigure(0, weight=1)
        cards_container.columnconfigure(1, weight=1)

        # Card 1: Capturar Good Card
        card1, btns1 = create_card(
            cards_container,
            "1) Capturar Good Card",
            "Conectar ao Chrome (CDP) e capturar vendas diretamente da aba do portal Good Card.",
            status_var=self.status_goodcard_capture,
        )
        card1.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns1, text="Ir para Captura", command=lambda: notebook.select(self.tab_captura)).pack(
            side="left"
        )

        # Card 2: Capturar Vale Card
        card2, btns2 = create_card(
            cards_container,
            "2) Capturar Vale Card (PDF)",
            "Ler PDF do Vale Card, extrair vendas positivas e calcular bloco de despesas.",
            status_var=self.status_valecard,
        )
        card2.grid(row=0, column=1, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns2, text="Ir para Captura", command=lambda: notebook.select(self.tab_captura)).pack(
            side="left"
        )

        # Card 3: Capturar Rede Frota
        card3, btns3 = create_card(
            cards_container,
            "3) Capturar Rede Frota (PDF)",
            "Ler PDF da Rede Frota e extrair as transações para o EMSYS.",
            status_var=self.status_redefrota,
        )
        card3.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns3, text="Ir para Captura", command=lambda: notebook.select(self.tab_captura)).pack(
            side="left"
        )

        # Card 4: Unificar / Resumo
        card4, btns4 = create_card(
            cards_container,
            "4) Unificar / Ver Resumo",
            "Ler todos os arquivos em capturas_portal, unificar e mostrar resumo para filtragem no portal.",
            status_var=self.status_unificado,
        )
        card4.grid(row=1, column=1, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns4, text="Unificar capturas agora", command=self._action_unificar).pack(side="left")

        # Card 5: Calibrar EMSYS
        card5, btns5 = create_card(
            cards_container,
            "5) Calibrar EMSYS",
            "Definir o ponto de clique do grid no EMSYS para que a automação possa navegar corretamente.",
            status_var=self.status_emsys_calibracao,
        )
        card5.grid(row=2, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns5, text="Ir para EMSYS", command=lambda: notebook.select(self.tab_emsys)).pack(side="left")

        # Card 6: Rodar EMSYS
        card6, btns6 = create_card(
            cards_container,
            "6) Rodar EMSYS",
            "Marcar automaticamente o grid do EMSYS com base nas capturas unificadas.",
            status_var=self.status_emsys_exec,
        )
        card6.grid(row=2, column=1, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns6, text="Ir para EMSYS", command=lambda: notebook.select(self.tab_emsys)).pack(side="left")

        # Card 7: Relatórios
        card7, btns7 = create_card(
            cards_container,
            "7) Abrir Relatórios",
            "Abrir os arquivos de resumo e as capturas para conferência.",
        )
        card7.grid(row=3, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns7, text="Ir para Relatórios", command=lambda: notebook.select(self.tab_relatorios)).pack(
            side="left"
        )

        # Rodapé com "by: lucas"
        footer = ttk.Label(frame, text="by: lucas", font=("Segoe UI", 8, "italic"))
        footer.grid(row=2, column=0, sticky="e", pady=(8, 0))

    def _build_tab_captura(self):
        frame = self.tab_captura
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        # Card Good Card
        card_gc, btns_gc = create_card(
            frame,
            "Good Card (Chrome/CDP)",
            "Conecte ao Chrome com porta 9222 aberta, escolha a aba do portal Good Card e capture as vendas.",
            status_var=self.status_goodcard,
        )
        card_gc.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")

        ttk.Button(
            btns_gc,
            text="Conectar ao Chrome (CDP)",
            command=self._action_goodcard_check_cdp,
        ).pack(side="left", padx=(0, 4))

        ttk.Button(
            btns_gc,
            text="Listar abas",
            command=self._action_goodcard_list_tabs,
        ).pack(side="left", padx=(0, 4))

        ttk.Button(
            btns_gc,
            text="Abrir nova aba do portal",
            command=self._action_goodcard_open_portal,
        ).pack(side="left")

        # Combobox para abas
        combo_frame = ttk.Frame(card_gc)
        combo_frame.grid(row=4, column=0, sticky="we", pady=(6, 2))
        ttk.Label(combo_frame, text="Aba do portal Good Card:").pack(side="left")
        self.combo_tabs = ttk.Combobox(
            combo_frame,
            textvariable=self.goodcard_tabs_var,
            state="readonly",
            width=50,
        )
        self.combo_tabs.pack(side="left", padx=(4, 0))

        ttk.Button(
            card_gc,
            text="Capturar da aba selecionada",
            command=self._action_goodcard_capture_selected,
        ).grid(row=5, column=0, sticky="w", pady=(4, 0))

        # Card Vale Card
        card_vc, btns_vc = create_card(
            frame,
            "Vale Card (PDF)",
            "Selecione o PDF do Vale Card para capturar vendas positivas e somar despesas (Taxa Administrativa e outras).",
            status_var=self.status_valecard,
        )
        card_vc.grid(row=0, column=1, padx=6, pady=6, sticky="nsew")

        ttk.Button(
            btns_vc,
            text="Selecionar PDF Vale Card",
            command=self._action_valecard_pdf,
        ).pack(side="left", padx=(0, 4))

        ttk.Button(
            btns_vc,
            text="Selecionar vários PDFs Vale Card",
            command=self._action_valecard_multiple_pdfs,
        ).pack(side="left")

        # Card Rede Frota
        card_rf, btns_rf = create_card(
            frame,
            "Rede Frota (PDF)",
            "Selecione o PDF da Rede Frota para capturar as transações do resumo.",
            status_var=self.status_redefrota,
        )
        card_rf.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")

        ttk.Button(
            btns_rf,
            text="Selecionar PDF Rede Frota",
            command=self._action_redefrota_pdf,
        ).pack(side="left")

        # Card Unificar
        card_unif, btns_unif = create_card(
            frame,
            "Unificar / Resumo",
            "Leia todos os arquivos em capturas_portal, unifique, calcule soma bruta e intervalo de datas.",
            status_var=self.status_unificado,
        )
        card_unif.grid(row=1, column=1, padx=6, pady=6, sticky="nsew")

        ttk.Button(btns_unif, text="Unificar capturas", command=self._action_unificar).pack(side="left", padx=(0, 4))
        ttk.Button(btns_unif, text="Limpar capturas", command=self._action_limpar_capturas).pack(side="left")

    def _build_tab_emsys(self):
        frame = self.tab_emsys
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        # Card Calibrar EMSYS
        card_cal, btns_cal = create_card(
            frame,
            "Calibrar EMSYS",
            "1) Abra o EMSYS em 'Geração de Fatura'.\n"
            "2) Clique em uma linha do grid.\n"
            "3) Clique em 'Capturar ponto do grid agora' e posicione o mouse sobre uma célula do grid.",
            status_var=self.status_emsys_calibracao,
        )
        card_cal.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")

        ttk.Button(
            btns_cal,
            text="Capturar ponto do grid agora",
            command=self._action_capturar_ponto_grid,
        ).pack(side="left", padx=(0, 4))

        ttk.Button(
            btns_cal,
            text="Salvar calibração",
            command=self._action_salvar_calibracao,
        ).pack(side="left")

        # Card Rodar EMSYS
        card_run, btns_run = create_card(
            frame,
            "Rodar EMSYS",
            "Deixe o EMSYS aberto em 'Geração de Fatura', clique em uma linha do grid e clique em 'Iniciar marcação'.\n"
            "Não mexa no mouse/teclado durante a execução (FAILSAFE ativo: canto superior esquerdo interrompe).",
            status_var=self.status_emsys_exec,
        )
        card_run.grid(row=0, column=1, padx=6, pady=6, sticky="nsew")

        self.btn_emsys_start = ttk.Button(
            btns_run,
            text="Iniciar marcação",
            command=self._action_rodar_emsys,
        )
        self.btn_emsys_start.pack(side="left", padx=(0, 4))

        self.btn_emsys_stop = ttk.Button(
            btns_run,
            text="Parar",
            command=self._action_parar_emsys,
            state="disabled",
        )
        self.btn_emsys_stop.pack(side="left")

        # Progresso
        prog_frame = ttk.Frame(card_run)
        prog_frame.grid(row=4, column=0, sticky="we", pady=(6, 4))
        ttk.Label(prog_frame, textvariable=self.emsys_progress_var).pack(anchor="w")
        ttk.Label(prog_frame, textvariable=self.emsys_last_value_var).pack(anchor="w")

        # Log rolável
        log_label = ttk.Label(card_run, text="Log da execução:")
        log_label.grid(row=5, column=0, sticky="w")
        self.text_log = ScrolledText(card_run, height=10, width=70, state="disabled")
        self.text_log.grid(row=6, column=0, sticky="nsew", pady=(4, 0))
        card_run.rowconfigure(6, weight=1)

    def _build_tab_relatorios(self):
        frame = self.tab_relatorios
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        # Card encontrados
        card_e, btns_e = create_card(
            frame,
            "Relatório: encontrados.txt",
            "Valores que foram encontrados e marcados no EMSYS.",
        )
        card_e.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns_e, text="Abrir encontrados.txt", command=lambda: self._open_file("encontrados.txt")).pack(
            side="left"
        )

        # Card nao_encontrados
        card_ne, btns_ne = create_card(
            frame,
            "Relatório: nao_encontrados.txt",
            "Valores que não foram localizados no grid do EMSYS.",
        )
        card_ne.grid(row=0, column=1, padx=6, pady=6, sticky="nsew")
        ttk.Button(
            btns_ne,
            text="Abrir nao_encontrados.txt",
            command=lambda: self._open_file("nao_encontrados.txt"),
        ).pack(side="left")

        # Card resumo
        card_r, btns_r = create_card(
            frame,
            "Relatório: resumo.txt",
            "Resumo geral de portal x EMSYS, incluindo bloco de despesas do Vale Card.",
        )
        card_r.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns_r, text="Abrir resumo.txt", command=lambda: self._open_file("resumo.txt")).pack(side="left")

        # Card capturas_portal
        card_cp, btns_cp = create_card(
            frame,
            "Pasta de capturas",
            "Abrir a pasta capturas_portal com todos os arquivos de captura gerados.",
        )
        card_cp.grid(row=1, column=1, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns_cp, text="Abrir pasta capturas_portal", command=self._open_capturas_dir).pack(side="left")

        # Card CSV
        card_csv, btns_csv = create_card(
            frame,
            "Exportar CSV",
            "Exporta a lista unificada de capturas para um arquivo CSV (capturas_unificadas.csv).",
        )
        card_csv.grid(row=2, column=0, padx=6, pady=6, sticky="nsew")
        ttk.Button(btns_csv, text="Exportar CSV", command=self._action_export_csv).pack(side="left")

    def _build_tab_config(self):
        frame = self.tab_config
        frame.columnconfigure(0, weight=1)

        # Card Good Card - Fallback URL
        card_cfg, btns_cfg = create_card(
            frame,
            "Good Card - URL do portal (fallback)",
            "URL usada para abrir uma nova aba do portal Good Card caso nenhuma aba seja encontrada.",
        )
        card_cfg.grid(row=0, column=0, padx=6, pady=6, sticky="nsew")

        settings = storage.load_settings()
        self.goodcard_url_var = tk.StringVar(value=settings.get("goodcard_fallback_url") or core.get_goodcard_fallback_url())

        entry = ttk.Entry(card_cfg, textvariable=self.goodcard_url_var, width=60)
        entry.grid(row=4, column=0, sticky="w", pady=(4, 4))

        ttk.Button(btns_cfg, text="Salvar URL", command=self._action_salvar_goodcard_url).pack(side="left")

        # Card Ajuda
        card_help, _ = create_card(
            frame,
            "Ajuda rápida",
            "Para Good Card (CDP), abra o Chrome com o comando:\n"
            r'"C:\Program Files\Google\Chrome\Application\chrome.exe" '
            r'--remote-debugging-port=9222 --user-data-dir=C:\chrome-automacao' "\n\n"
            "Certifique-se também de ter executado:\n"
            " - pip install -r requirements.txt\n"
            " - python -m playwright install\n",
        )
        card_help.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")

    # --------------------------------------------------------------------- Ações (handlers)
    def _run_in_thread(self, target, *args, **kwargs):
        t = threading.Thread(target=target, args=args, kwargs=kwargs, daemon=True)
        t.start()
        return t

    # -------- Good Card
    def _action_goodcard_check_cdp(self):
        def worker():
            ok, err = core.goodcard_check_cdp()
            if ok:
                self.event_queue.put(
                    {"type": "ui", "action": "set_status_goodcard", "text": "Chrome (CDP) acessível em 127.0.0.1:9222."}
                )
            else:
                msg = (
                    "Não consegui conectar ao Chrome via CDP em 127.0.0.1:9222.\n\n"
                    "Abra o Chrome com o comando:\n"
                    r'"C:\Program Files\Google\Chrome\Application\chrome.exe" '
                    r'--remote-debugging-port=9222 --user-data-dir=C:\chrome-automacao'
                    f"\n\nDetalhes técnicos: {err}"
                )
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Chrome (CDP) não encontrado",
                        "message": msg,
                    }
                )

        self._run_in_thread(worker)

    def _action_goodcard_list_tabs(self):
        def worker():
            try:
                tabs = core.goodcard_list_tabs()
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro ao listar abas",
                        "message": str(e),
                    }
                )
                return

            self.event_queue.put({"type": "ui", "action": "update_goodcard_tabs", "tabs": tabs})

        self._run_in_thread(worker)

    def _action_goodcard_open_portal(self):
        def worker():
            try:
                core.goodcard_open_portal_tab()
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "info_message",
                        "title": "Nova aba aberta",
                        "message": "Foi aberta uma nova aba do portal Good Card no Chrome.\n"
                        "Se necessário, faça login e volte para listar as abas.",
                    }
                )
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro ao abrir aba do portal",
                        "message": str(e),
                    }
                )

        self._run_in_thread(worker)

    def _action_goodcard_capture_selected(self):
        selection = self.goodcard_tabs_var.get()
        if not selection:
            messagebox.showwarning("Good Card", "Selecione primeiro uma aba na lista.")
            return

        # A string da combo é "Título | URL"
        url = None
        for tab in self.goodcard_tabs:
            label = f"{tab.get('title','') or '(sem título)'} | {tab.get('url','')}"
            if label == selection:
                url = tab.get("url")
                break

        if not url:
            messagebox.showerror("Good Card", "Não consegui identificar a URL da aba selecionada.")
            return

        def worker():
            try:
                rows = core.goodcard_capture_from_url(url)
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro na captura Good Card",
                        "message": str(e),
                    }
                )
                return

            if not rows:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "info_message",
                        "title": "Captura Good Card",
                        "message": "Nenhuma transação foi encontrada na aba selecionada.",
                    }
                )
                return

            fn = save_path = core.save_capture_txt(rows, "GoodCard")  # type: ignore[attr-defined]

            dmin, dmax = core.date_range_from_rows(rows)  # type: ignore[attr-defined]
            if dmin and dmax:
                intervalo = f"{dmin.strftime('%d/%m/%Y %H:%M:%S')}  até  {dmax.strftime('%d/%m/%Y %H:%M:%S')}"
            else:
                intervalo = "N/D"

            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "goodcard_captured",
                    "count": len(rows),
                    "file": save_path,
                    "intervalo": intervalo,
                }
            )

        self._run_in_thread(worker)

    # -------- Vale Card
    def _action_valecard_pdf(self):
        pdf_path = filedialog.askopenfilename(
            title="Selecione o PDF do Vale Card",
            filetypes=[("PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if not pdf_path:
            return

        def worker():
            try:
                rows = core.valecard_capture_from_pdf(pdf_path)  # type: ignore[attr-defined]
                desp = core.valecard_somar_despesas_pdf(pdf_path)  # type: ignore[attr-defined]
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro Vale Card",
                        "message": str(e),
                    }
                )
                return

            saved_file = None
            if rows:
                saved_file = core.save_capture_txt(rows, "ValeCard")  # type: ignore[attr-defined]

            # Atualizar arquivo de despesas (mantendo compatibilidade)
            total_abs = abs(desp.get("total_despesas", 0.0))
            taxa_abs = abs(desp.get("total_taxa_adm", 0.0))
            outras_abs = abs(desp.get("total_outras", 0.0))
            try:
                with open(core.VALE_DESP_FILE, "w", encoding="utf-8") as f:  # type: ignore[attr-defined]
                    json.dump(
                        {
                            "total_despesas_abs": total_abs,
                            "taxa_adm_abs": taxa_abs,
                            "outras_abs": outras_abs,
                            "arquivo": pdf_path,
                            "atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                        },
                        f,
                        ensure_ascii=False,
                        indent=2,
                    )
            except Exception:
                pass

            dmin, dmax = core.date_range_from_rows(rows) if rows else (None, None)  # type: ignore[attr-defined]
            if dmin and dmax:
                intervalo = f"{dmin.strftime('%d/%m/%Y %H:%M:%S')}  até  {dmax.strftime('%d/%m/%Y %H:%M:%S')}"
            else:
                intervalo = "N/D"

            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "valecard_processed",
                    "count": len(rows),
                    "file": saved_file,
                    "intervalo": intervalo,
                    "total_despesas": total_abs,
                    "taxa_adm": taxa_abs,
                    "outras": outras_abs,
                }
            )

        # Imports locais para evitar ciclos no topo do arquivo
        import json  # type: ignore
        from datetime import datetime  # type: ignore

        self._run_in_thread(worker)

    def _action_valecard_multiple_pdfs(self):
        pdf_paths = filedialog.askopenfilenames(
            title="Selecione os PDFs do Vale Card (Ctrl+clique para vários)",
            filetypes=[("PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if not pdf_paths:
            return

        def worker():
            import json as _json
            from datetime import datetime as _dt

            all_rows = []
            seen = set()
            total_despesas = 0.0
            total_taxa_adm = 0.0
            total_outras = 0.0
            errors = []

            for pdf_path in pdf_paths:
                try:
                    rows = core.valecard_capture_from_pdf(pdf_path)
                    desp = core.valecard_somar_despesas_pdf(pdf_path)
                except Exception as e:
                    errors.append(f"{os.path.basename(pdf_path)}: {e}")
                    continue

                for r in rows:
                    key = (r.get("dt", ""), r.get("bruto", ""), r.get("id", ""))
                    if key not in seen:
                        seen.add(key)
                        all_rows.append(r)

                total_despesas += desp.get("total_despesas", 0.0)
                total_taxa_adm += desp.get("total_taxa_adm", 0.0)
                total_outras += desp.get("total_outras", 0.0)

            if errors:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "warning_message",
                        "title": "Alguns PDFs falharam",
                        "message": "Os seguintes arquivos apresentaram erro:\n\n" + "\n".join(errors),
                    }
                )

            if not all_rows and not (total_despesas or total_taxa_adm or total_outras):
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "info_message",
                        "title": "Vale Card - Vários PDFs",
                        "message": "Nenhuma venda ou despesa encontrada nos PDFs selecionados.",
                    }
                )
                return

            saved_file = None
            if all_rows:
                saved_file = core.save_capture_txt(all_rows, "ValeCard")

            soma_vendas = sum(core.brl_to_float(r.get("bruto", "")) for r in all_rows)
            dmin, dmax = core.date_range_from_rows(all_rows) if all_rows else (None, None)
            if dmin and dmax:
                intervalo = f"{dmin.strftime('%d/%m/%Y %H:%M:%S')}  até  {dmax.strftime('%d/%m/%Y %H:%M:%S')}"
            else:
                intervalo = "N/D"

            total_desp_abs = abs(total_despesas)
            taxa_abs = abs(total_taxa_adm)
            outras_abs = abs(total_outras)

            try:
                with open(core.VALE_DESP_FILE, "w", encoding="utf-8") as f:
                    _json.dump(
                        {
                            "total_despesas_abs": total_desp_abs,
                            "taxa_adm_abs": taxa_abs,
                            "outras_abs": outras_abs,
                            "arquivo": f"{len(pdf_paths)} PDF(s) agregados",
                            "atualizado_em": _dt.now().strftime("%d/%m/%Y %H:%M:%S"),
                        },
                        f,
                        ensure_ascii=False,
                        indent=2,
                    )
            except Exception:
                pass

            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "valecard_processed",
                    "count": len(all_rows),
                    "file": saved_file,
                    "intervalo": intervalo,
                    "total_despesas": total_desp_abs,
                    "taxa_adm": taxa_abs,
                    "outras": outras_abs,
                    "soma_vendas": soma_vendas,
                    "num_pdfs": len(pdf_paths),
                }
            )

        self._run_in_thread(worker)

    # -------- Rede Frota
    def _action_redefrota_pdf(self):
        pdf_path = filedialog.askopenfilename(
            title="Selecione o PDF da Rede Frota",
            filetypes=[("PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if not pdf_path:
            return

        def worker():
            try:
                rows = core.redefrota_capture_from_pdf(pdf_path)  # type: ignore[attr-defined]
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro Rede Frota",
                        "message": str(e),
                    }
                )
                return

            if not rows:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "info_message",
                        "title": "Captura Rede Frota",
                        "message": "Nenhuma transação foi encontrada neste PDF.",
                    }
                )
                return

            save_path = core.save_capture_txt(rows, "RedeFrota")  # type: ignore[attr-defined]
            dmin, dmax = core.date_range_from_rows(rows)  # type: ignore[attr-defined]
            if dmin and dmax:
                intervalo = f"{dmin.strftime('%d/%m/%Y %H:%M:%S')}  até  {dmax.strftime('%d/%m/%Y %H:%M:%S')}"
            else:
                intervalo = "N/D"

            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "redefrota_processed",
                    "count": len(rows),
                    "file": save_path,
                    "intervalo": intervalo,
                }
            )

        self._run_in_thread(worker)

    # -------- Unificar / Limpar capturas / CSV
    def _action_unificar(self):
        def worker():
            try:
                summary = core.summarize_unified_captures()
                vale = core.load_valecard_despesas()
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro ao unificar capturas",
                        "message": str(e),
                    }
                )
                return

            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "unified",
                    "summary": summary,
                    "vale": vale,
                }
            )

        self._run_in_thread(worker)

    def _action_limpar_capturas(self):
        if not messagebox.askyesno(
            "Limpar capturas",
            "Tem certeza que deseja apagar TODOS os arquivos em capturas_portal?",
        ):
            return

        def worker():
            count = core.clear_captures()
            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "info_message",
                    "title": "Capturas limpas",
                    "message": f"Foram removidos {count} arquivo(s) de captura.",
                }
            )

        self._run_in_thread(worker)

    def _action_export_csv(self):
        csv_path = os.path.join(os.getcwd(), "capturas_unificadas.csv")

        def worker():
            try:
                info = core.export_unified_to_csv(csv_path)
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro ao exportar CSV",
                        "message": str(e),
                    }
                )
                return

            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "info_message",
                    "title": "CSV exportado",
                    "message": f"Arquivo gerado: {csv_path}\nTotal de linhas: {info.get('total', 0)}",
                }
            )

        self._run_in_thread(worker)

    # -------- EMSYS
    def _action_capturar_ponto_grid(self):
        messagebox.showinfo(
            "Calibrar EMSYS",
            "Em 3 segundos, o sistema vai capturar a posição atual do mouse.\n"
            "Posicione o mouse sobre uma célula do grid do EMSYS.",
        )

        def capture():
            import pyautogui  # type: ignore[import]

            x, y = pyautogui.position()
            self._last_grid_point = {"x": x, "y": y}
            self.status_emsys_calibracao.set(f"Ponto capturado: ({x}, {y}). Clique em 'Salvar calibração'.")

        self.root.after(3000, capture)

    def _action_salvar_calibracao(self):
        point = getattr(self, "_last_grid_point", None)
        if not point:
            messagebox.showwarning(
                "Calibrar EMSYS",
                "Primeiro use o botão 'Capturar ponto do grid agora' com o mouse sobre o grid.",
            )
            return

        try:
            core.save_emsys_config_from_gui(point)
        except Exception as e:
            messagebox.showerror("Calibrar EMSYS", f"Erro ao salvar calibração: {e}")
            return

        self.status_emsys_calibracao.set("Calibração salva com sucesso.")
        messagebox.showinfo("Calibrar EMSYS", "Calibração salva com sucesso.")

    def _action_rodar_emsys(self):
        # Garante que temos capturas unificadas na memória
        def worker_prepare_and_run():
            try:
                summary = core.summarize_unified_captures()
            except Exception as e:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Erro ao ler capturas",
                        "message": str(e),
                    }
                )
                return

            items = summary.get("items") or []
            if not items:
                self.event_queue.put(
                    {
                        "type": "ui",
                        "action": "error_message",
                        "title": "Sem capturas",
                        "message": "Não há capturas unificadas. Faça pelo menos uma captura antes de rodar o EMSYS.",
                    }
                )
                return

            # Pergunta rápida de confirmação
            self.event_queue.put(
                {
                    "type": "ui",
                    "action": "confirm_rodar_emsys",
                    "items": items,
                }
            )

        self._run_in_thread(worker_prepare_and_run)

    def _action_parar_emsys(self):
        """
        Solicita parada graciosa da marcação no EMSYS.
        A thread atual termina ao fim do ciclo em andamento.
        """
        ev = self._emsys_cancel_event
        if ev is not None:
            ev.set()
            self._append_log("Parada solicitada pelo usuário. Aguardando finalizar a linha atual...")
            self.status_emsys_exec.set("Parada solicitada. Aguardando finalizar a linha atual...")
            # Evita múltiplos cliques
            if hasattr(self, "btn_emsys_stop"):
                self.btn_emsys_stop.configure(state="disabled")

    # --------------------------------------------------------------------- Utilitários de UI
    def _append_log(self, text: str):
        self.text_log.configure(state="normal")
        self.text_log.insert("end", text + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")

    def _open_file(self, filename: str):
        path = os.path.join(os.getcwd(), filename)
        if not os.path.exists(path):
            messagebox.showwarning("Abrir arquivo", f"O arquivo '{filename}' ainda não existe.")
            return
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Abrir arquivo", f"Não foi possível abrir o arquivo:\n{e}")

    def _open_capturas_dir(self):
        from robo_cartoes_emsys_v3 import CAPTURES_DIR as CORE_CAPTURES_DIR  # type: ignore[import]

        path = os.path.join(os.getcwd(), CORE_CAPTURES_DIR)
        if not os.path.exists(path):
            os.makedirs(path, exist_ok=True)
        try:
            os.startfile(path)  # type: ignore[attr-defined]
        except Exception as e:
            messagebox.showerror("Abrir pasta", f"Não foi possível abrir a pasta de capturas:\n{e}")

    def _action_salvar_goodcard_url(self):
        url = self.goodcard_url_var.get().strip()
        storage.save_settings({"goodcard_fallback_url": url})
        messagebox.showinfo("Configuração", "URL de fallback do Good Card salva com sucesso.")

    # --------------------------------------------------------------------- Processamento da fila de eventos
    def _process_event_queue(self):
        while True:
            try:
                ev = self.event_queue.get_nowait()
            except queue.Empty:
                break

            if ev.get("type") == "ui":
                self._handle_ui_event(ev)
            elif ev.get("type") in ("start", "progress", "log", "end", "error"):
                # Eventos vindos do core.run_emsys_marking_with_progress
                self._handle_emsys_event(ev)

        self.root.after(200, self._process_event_queue)

    def _handle_ui_event(self, ev: Dict[str, Any]):
        action = ev.get("action")

        if action == "set_status_goodcard":
            self.status_goodcard.set(ev.get("text", ""))

        elif action == "error_message":
            messagebox.showerror(ev.get("title", "Erro"), ev.get("message", ""))

        elif action == "info_message":
            messagebox.showinfo(ev.get("title", "Informação"), ev.get("message", ""))

        elif action == "warning_message":
            messagebox.showwarning(ev.get("title", "Atenção"), ev.get("message", ""))

        elif action == "update_goodcard_tabs":
            tabs = ev.get("tabs", [])
            self.goodcard_tabs = tabs
            if not tabs:
                self.status_goodcard.set("Nenhuma aba encontrada via CDP. Verifique se o Chrome está aberto.")
                self.combo_tabs["values"] = ()
                self.goodcard_tabs_var.set("")
                return
            labels = [
                f"{t.get('title','') or '(sem título)'} | {t.get('url','')}"  # type: ignore[dict-item]
                for t in tabs
            ]
            self.combo_tabs["values"] = labels
            self.goodcard_tabs_var.set(labels[0])
            self.status_goodcard.set(f"{len(tabs)} aba(s) detectadas via CDP.")

        elif action == "goodcard_captured":
            count = ev.get("count", 0)
            file = ev.get("file")
            intervalo = ev.get("intervalo", "")
            self.status_goodcard_capture.set(
                f"Capturadas {count} transações do Good Card. Arquivo: {file}. Intervalo: {intervalo}"
            )
            messagebox.showinfo(
                "Captura Good Card",
                f"Capturadas {count} transações.\nArquivo salvo em:\n{file}\n\nIntervalo: {intervalo}",
            )

        elif action == "valecard_processed":
            count = ev.get("count", 0)
            file = ev.get("file")
            intervalo = ev.get("intervalo", "")
            total_desp = ev.get("total_despesas", 0.0)
            taxa_adm = ev.get("taxa_adm", 0.0)
            outras = ev.get("outras", 0.0)
            soma_vendas = ev.get("soma_vendas", 0.0)
            num_pdfs = ev.get("num_pdfs")
            if num_pdfs is not None and num_pdfs > 1:
                status_text = (
                    f"{num_pdfs} PDFs | Vendas: {count} transações | Total vendas: R$ {soma_vendas:.2f} | "
                    f"Menor/maior data: {intervalo} | Despesas: total={total_desp:.2f}"
                )
                self.status_valecard.set(status_text)
                msg = (
                    f"Vale Card - {num_pdfs} PDFs processados\n\n"
                    f"Total de vendas: R$ {soma_vendas:.2f} ({count} transações)\n\n"
                    f"Intervalo de datas:\n"
                    f"  Menor data: {intervalo.split('  até  ')[0] if '  até  ' in intervalo else intervalo}\n"
                    f"  Maior data: {intervalo.split('  até  ')[-1] if '  até  ' in intervalo else intervalo}\n\n"
                    f"Despesas agregadas:\n"
                    f"  Total: R$ {total_desp:.2f} | Taxa adm: R$ {taxa_adm:.2f} | Outras: R$ {outras:.2f}\n\n"
                    f"Arquivo salvo: {file or 'N/D'}"
                )
                messagebox.showinfo("Vale Card - Vários PDFs", msg)
            else:
                self.status_valecard.set(
                    f"Vendas: {count} (arquivo: {file or 'N/D'}) | Intervalo: {intervalo} | "
                    f"Despesas: total={total_desp:.2f}, taxa adm={taxa_adm:.2f}, outras={outras:.2f}"
                )

        elif action == "redefrota_processed":
            count = ev.get("count", 0)
            file = ev.get("file")
            intervalo = ev.get("intervalo", "")
            self.status_redefrota.set(
                f"Capturadas {count} transações Rede Frota. Arquivo: {file}. Intervalo: {intervalo}"
            )

        elif action == "unified":
            summary = ev.get("summary") or {}
            vale = ev.get("vale")
            self.unified_items = summary.get("items") or []
            total = summary.get("total", 0)
            soma = summary.get("soma", 0.0)
            dmin = summary.get("dmin")
            dmax = summary.get("dmax")
            if dmin and dmax:
                intervalo = f"{dmin.strftime('%d/%m/%Y %H:%M:%S')}  até  {dmax.strftime('%d/%m/%Y %H:%M:%S')}"
            else:
                intervalo = "N/D"

            msg = (
                f"Capturas unificadas: {total}\n"
                f"Soma bruta (referência): R$ {soma:.2f}\n"
                f"Intervalo de datas: {intervalo}\n"
            )
            if vale:
                msg += (
                    "\nBloco de despesas Vale Card (último PDF lido):\n"
                    f" - Total despesas: R$ {vale.get('total_despesas_abs', 0.0):.2f}\n"
                    f" - Taxa administrativa: R$ {vale.get('taxa_adm_abs', 0.0):.2f}\n"
                    f" - Outras despesas: R$ {vale.get('outras_abs', 0.0):.2f}\n"
                )

            self.status_unificado.set(
                f"Total unificado: {total} | Soma bruta: R$ {soma:.2f} | Intervalo: {intervalo}"
            )
            messagebox.showinfo("Unificação de capturas", msg)

        elif action == "confirm_rodar_emsys":
            items = ev.get("items") or []
            if not items:
                return
            if not messagebox.askyesno(
                "Rodar EMSYS",
                "Confirme que o EMSYS está aberto na tela 'Geração de Fatura' e uma linha do grid está selecionada.\n\n"
                "Deseja iniciar a marcação agora?",
            ):
                return

            # Inicia thread de execução do EMSYS com callback de progresso
            self._emsys_cancel_event = threading.Event()
            self.emsys_progress_var.set(f"Marcado: 0/{len(items)}")
            self.emsys_last_value_var.set("Último valor marcado: -")
            self.text_log.configure(state="normal")
            self.text_log.delete("1.0", "end")
            self.text_log.configure(state="disabled")
            self.status_emsys_exec.set("EMSYS em execução...")

            # Ajusta botões
            if hasattr(self, "btn_emsys_start"):
                self.btn_emsys_start.configure(state="disabled")
            if hasattr(self, "btn_emsys_stop"):
                self.btn_emsys_stop.configure(state="normal")

            def worker_run():
                core.run_emsys_marking_with_progress(items, self.event_queue.put, cancel_event=self._emsys_cancel_event)

            self.emsys_thread = self._run_in_thread(worker_run)

    def _handle_emsys_event(self, ev: Dict[str, Any]):
        etype = ev.get("type")

        if etype == "start":
            total = ev.get("total_portal", 0)
            self.emsys_progress_var.set(f"Marcado: 0/{total}")
            self._append_log(f"Iniciando marcação no EMSYS. Total de transações: {total}")

        elif etype == "progress":
            marcado = ev.get("marcado", 0)
            total = ev.get("total", 0)
            valor = ev.get("valor", "")
            self.emsys_progress_var.set(f"Marcado: {marcado}/{total}")
            self.emsys_last_value_var.set(f"Último valor marcado: {valor}")
            self._append_log(f"✅ Marcado: {valor} ({marcado}/{total})")

        elif etype == "log":
            msg = ev.get("message", "")
            if msg:
                self._append_log(msg)

        elif etype == "error":
            msg = ev.get("message", "Erro desconhecido durante a marcação.")
            self._append_log(f"Erro: {msg}")
            messagebox.showerror("EMSYS - erro", msg)
            self.status_emsys_exec.set("EMSYS - erro na execução.")
            # Reabilita botões
            if hasattr(self, "btn_emsys_start"):
                self.btn_emsys_start.configure(state="normal")
            if hasattr(self, "btn_emsys_stop"):
                self.btn_emsys_stop.configure(state="disabled")

        elif etype == "end":
            total = ev.get("total_portal", 0)
            marcados = ev.get("marcados", 0)
            nao_encontrados = ev.get("nao_encontrados", 0)
            soma_marcados = ev.get("soma_marcados", 0.0)
            soma_nao = ev.get("soma_nao_encontrados", 0.0)
            self.status_emsys_exec.set("Execução EMSYS finalizada.")
            self._append_log("Execução EMSYS finalizada.")
            # Reabilita botões
            if hasattr(self, "btn_emsys_start"):
                self.btn_emsys_start.configure(state="normal")
            if hasattr(self, "btn_emsys_stop"):
                self.btn_emsys_stop.configure(state="disabled")

            msg = (
                "EMSYS finalizado.\n\n"
                f"Total portal (unificado): {total}\n"
                f"Marcados EMSYS: {marcados}\n"
                f"Não encontrados: {nao_encontrados}\n\n"
                f"Soma marcados: R$ {soma_marcados:.2f}\n"
                f"Soma não encontrados: R$ {soma_nao:.2f}\n\n"
                "Arquivos gerados: encontrados.txt, nao_encontrados.txt, resumo.txt"
            )
            if messagebox.askyesno("EMSYS finalizado", msg + "\n\nDeseja abrir a pasta do aplicativo agora?"):
                try:
                    os.startfile(os.getcwd())  # type: ignore[attr-defined]
                except Exception:
                    pass


def main():
    # Se ttkbootstrap estiver disponível, usamos o tema para dar cara mais moderna
    if HAS_BOOTSTRAP:
        style = BootstrapStyle(theme="flatly")  # type: ignore[call-arg]
        root = style.master  # type: ignore[assignment]
    else:
        root = tk.Tk()

    app = App(root)
    root.minsize(980, 640)
    root.mainloop()


if __name__ == "__main__":
    main()

