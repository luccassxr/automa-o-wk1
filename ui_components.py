import tkinter as tk
from tkinter import ttk
from typing import Tuple


def create_card(
    parent,
    title: str,
    description: str = "",
    status_var: tk.StringVar | None = None,
) -> Tuple[ttk.Frame, ttk.Frame]:
    """
    Cria um "card" visual (painel) com título, descrição, área para botões
    e, opcionalmente, um label de status.

    Retorna (frame_card, frame_buttons) para o chamador adicionar botões.
    """
    card = ttk.Frame(parent, padding=(16, 12))

    title_lbl = ttk.Label(card, text=title, style="CardTitle.TLabel")
    title_lbl.grid(row=0, column=0, sticky="w")

    if description:
        desc_lbl = ttk.Label(card, text=description, style="CardDesc.TLabel", wraplength=480, justify="left")
        desc_lbl.grid(row=1, column=0, sticky="w", pady=(4, 4))

    btn_frame = ttk.Frame(card)
    btn_frame.grid(row=2, column=0, sticky="w", pady=(4, 4))

    if status_var is not None:
        status_lbl = ttk.Label(card, textvariable=status_var, style="CardStatus.TLabel")
        status_lbl.grid(row=3, column=0, sticky="w", pady=(4, 0))

    # Borda visual de card
    card.configure(style="Card.TFrame")

    return card, btn_frame


def setup_styles(root: tk.Misc):
    """
    Define estilos básicos dos cards e fontes.
    Se ttkbootstrap estiver disponível, assume que o tema principal já está aplicado.
    """
    style = ttk.Style(root)

    # Caso o tema atual não tenha Card.TFrame, definimos algo simples.
    style.configure(
        "Card.TFrame",
        relief="raised",
        borderwidth=1,
        padding=10,
    )
    style.configure(
        "CardTitle.TLabel",
        font=("Segoe UI", 11, "bold"),
    )
    style.configure(
        "CardDesc.TLabel",
        font=("Segoe UI", 9),
        foreground="#444444",
    )
    style.configure(
        "CardStatus.TLabel",
        font=("Segoe UI", 9, "italic"),
        foreground="#006400",
    )

