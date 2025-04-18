import dash
from dash import Dash, dcc, html, dash_table, Input, Output, State
import plotly.express as px
import pandas as pd

# --- Lettura dei dati ---
df_iniziative = pd.read_excel("Partecipate Brescia copia.xlsx", sheet_name="Attività generali")
df_composizione = pd.read_excel("Partecipate Brescia copia.xlsx", sheet_name="Genere nelle aziende ")

# Pulizia minima: rimuove spazi indesiderati
df_iniziative["Nome azienda"] = df_iniziative["Nome azienda"].str.strip()
df_composizione["Nome azienda"] = df_composizione["Nome azienda"].str.strip()

# --- Funzione per creare le opzioni dei dropdown ---
def crea_opzioni(colonna, df):
    valori = sorted(df[colonna].dropna().unique())
    return [{"label": "Tutti", "value": "all"}] + [{"label": str(val), "value": val} for val in valori]

# Dropdown comuni per le sezioni relative a df_iniziative
opzioni_anno = crea_opzioni("Anno", df_iniziative)
opzioni_azienda = crea_opzioni("Nome azienda", df_iniziative)
opzioni_area = crea_opzioni("Area Prassi", df_iniziative)
opzioni_categoria = crea_opzioni("Categoria di diversità", df_iniziative)

# Dropdown per la sezione Composizione di Genere (basati su df_composizione)
opzioni_anno_genere = crea_opzioni("Anno", df_composizione)
opzioni_azienda_genere = crea_opzioni("Nome azienda", df_composizione)
opzioni_posizione_genere = crea_opzioni("Posizione", df_composizione)

# Palette personalizzata
PALETTE = ["#c0c0c0", "#c10a13"]

# --- Funzione per calcolare il D&I Index per ciascuna azienda ---
# --- Nuova funzione per calcolare i 3 indici e quello finale ---
def calcola_tre_indici(df_iniziative, df_genere):
    # -------- indice iniziative ----------
    n_iniziative = df_iniziative.groupby("Nome azienda")["Titolo dell'attività"].count()
    min_iniz, max_iniz = n_iniziative.min(), n_iniziative.max()
    indice_iniziative = (
        ((n_iniziative - min_iniz) / (max_iniz - min_iniz)) * 100
        if max_iniz != min_iniz else 0
    )

    # -------- indice categorie ----------
    categorie_rilevanti = ["Genere", "Età", "Disabilità", "Cultura", "LGBTQI+"]
    def _idx_cat(gruppo):
    # quante delle categorie_rilevanti compaiono almeno una volta?
        presenti = (
         gruppo.loc[gruppo["Categoria di diversità"].isin(categorie_rilevanti),
                    "Categoria di diversità"]
            .unique()
      )
        return (len(presenti) / len(categorie_rilevanti)) * 100
    indice_categorie = df_iniziative.groupby("Nome azienda").apply(_idx_cat)


    # -------- indice parità di genere ----------
    df_board = df_genere[df_genere["Posizione"] == "Board"].copy()      # <‑‑ usa df_genere!
    # conversione robusta: funziona sia se sono stringhe "42%" sia se sono già numeri 42/0.42
    df_board["Percentuale donne"] = (
        pd.to_numeric(
            df_board["Percentuale donne"].astype(str).str.replace('%', ''),
            errors="coerce"
        )
    )
    # se Excel restituisce 0–1 anziché 0–100 trasformo
    if df_board["Percentuale donne"].max() <= 1:
        df_board["Percentuale donne"] *= 100

    def _idx_parita(gruppo):
        perc = gruppo["Percentuale donne"].mean()
        return 100 - abs(50 - perc) * 2        # 50 → 100, 0/100 → 0

    indice_parita_genere = df_board.groupby("Nome azienda").apply(_idx_parita)

    # -------- unisco tutto ----------
    risultati = (
        pd.DataFrame({
            "Indice iniziative": indice_iniziative,
            "Indice categorie": indice_categorie,
            "Indice parità genere": indice_parita_genere,
        })
        .fillna(0)
        .reset_index(names="Nome azienda")      # così il nome della colonna è esplicito
    )

    risultati["Indice diversità finale"] = risultati[
        ["Indice iniziative", "Indice categorie", "Indice parità genere"]
    ].mean(axis=1)

    return risultati



def crea_grafici_indici(risultati):
    fig_iniziative = px.bar(
        risultati,
        x="Nome azienda",
        y="Indice iniziative",
        title="Indice Iniziative (0–100)",
        color="Indice iniziative",
        color_continuous_scale=PALETTE
    )
    fig_iniziative.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")

    fig_categorie = px.bar(
        risultati,
        x="Nome azienda",
        y="Indice categorie",
        title="Indice Categorie (0–100)",
        color="Indice categorie",
        color_continuous_scale=PALETTE
    )
    fig_categorie.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")

    fig_parita = px.bar(
        risultati,
        x="Nome azienda",
        y="Indice parità genere",
        title="Indice Parità di Genere (0–100)",
        color="Indice parità genere",
        color_continuous_scale=PALETTE
    )
    fig_parita.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")

    fig_media = px.bar(
        risultati,
        x="Nome azienda",
        y="Indice diversità finale",
        title="Indice Diversità Finale (Media dei tre indici)",
        color="Indice diversità finale",
        color_continuous_scale=PALETTE
    )
    fig_media.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")

    return fig_iniziative, fig_categorie, fig_parita, fig_media

# --- Creazione dell'app Dash ---
app = Dash(__name__)
server = app.server

app.layout = html.Div([
    dcc.Tabs(id="tabs", value="Introduzione", children=[
        # Tab Panoramica
        # ---------- TAB INTRODUZIONE (sostituisci tutta la tua versione con questa) ----------
dcc.Tab(
    label="Introduzione",
    value="tab-introduzione",
    children=[
        html.Div(
            [
                # ---------- intestazione rossa ----------
                html.Div(
                    html.H1(
                        "Dashboard D&I - Metodo e Fonti",
                        style={"margin": "0", "padding": "20px", "color": "#ffffff"},
                    ),
                    style={"backgroundColor": "#c10b13", "textAlign": "center"},
                ),

                # ---------- contenuto bianco ----------
                html.Div(
                    [
                        html.H2("Metodo", style={"color": "#080808"}),

                        html.P(
                            "Questa Dashboard è stata realizzata con l'obiettivo di analizzare e "
                            "visualizzare le iniziative in ambito Diversity & Inclusion promosse "
                            "dalle aziende partecipate del Comune di Brescia. I dati sono stati "
                            "raccolti, aggregati e analizzati a partire da fonti ufficiali, con "
                            "particolare attenzione alla presenza di attività concrete, pratiche "
                            "inclusive e alla composizione di genere nei ruoli aziendali.",
                            style={
                                "fontSize": "16px",
                                "lineHeight": "1.6",
                                "color": "#080808",
                            },
                        ),

                        html.P(
                            "La costruzione degli indicatori sintetici (D&I Index) si basa su tre "
                            "dimensioni principali:",
                            style={
                                "fontSize": "16px",
                                "lineHeight": "1.6",
                                "color": "#080808",
                            },
                        ),

                        # elenco puntato delle tre componenti
                        html.Ul(
                            [
                                html.Li(
                                    [
                                        html.B("Indice Iniziative – "),
                                        "numero totale di iniziative promosse da ogni azienda.",
                                    ]
                                ),
                                html.Li(
                                    [
                                        html.B("Indice Categorie – "),
                                        "copertura delle sette categorie di diversità considerate "
                                        "(Genere, Età, Disabilità, LGBTQI+, Religione e credo, Etnia, e infine attività generali (Cultura)); l’azienda ottiene "
                                        "il 100 % solo se ha almeno un’iniziativa per ciascuna categoria.",
                                    ]
                                ),
                                html.Li(
                                    [
                                        html.B("Indice Parità di Genere – "),
                                        "distanza dall’equilibrio 50 / 50 nella composizione del Board "
                                        "aziendale, rimappata su una scala 0‑100.",
                                    ]
                                ),
                            ],
                            style={
                                "fontSize": "16px",
                                "lineHeight": "1.6",
                                "color": "#080808",
                                "marginLeft": "20px",
                            },
                        ),

                        html.P(
                            "La media dei tre punteggi, normalizzati su scala 0‑100, restituisce "
                            "il valore finale del D&I Index, che permette il confronto omogeneo tra aziende.",
                            style={
                                "fontSize": "16px",
                                "lineHeight": "1.6",
                                "color": "#080808",
                            },
                        ),

                        # ---------- fonti dei dati ----------
                        html.H2("Fonti dei Dati", style={"color": "#080808", "marginTop": "30px"}),
                        html.Ul(
                            [
                                html.Li("Bilanci di sostenibilità delle aziende partecipate"),
                                html.Li("Siti istituzionali e documentazione pubblica delle aziende"),
                                html.Li("Database Orbis (Bureau van Dijk) per informazioni anagrafiche e di governance"),
                            ],
                            style={"fontSize": "16px", "color": "#080808"},
                        ),

                        html.P(
                            "I dati sono stati raccolti ed elaborati manualmente con un processo di "
                            "verifica incrociata, per garantire la coerenza e l'affidabilità delle "
                            "informazioni visualizzate.",
                            style={
                                "fontSize": "16px",
                                "lineHeight": "1.6",
                                "color": "#080808",
                            },
                        ),
                    ],
                    style={"padding": "30px"},
                ),
            ],
            style={
                "fontFamily": "Arial, sans-serif",
                "backgroundColor": "#f7f7f7",
                "padding": "20px",
            },
        )
    ],
),

        dcc.Tab(label="Panoramica", value="tab-overview", children=[
            html.Div([
                html.Div([
                    html.H1("Dashboard D&I - Sezione Panoramica", style={"margin": "0", "padding": "20px", "color": "#ffffff"})
                ], style={"backgroundColor": "#c10b13", "textAlign": "center"}),
                html.Div([
                    html.Div([
                        html.Label("Azienda", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-azienda-overview", options=opzioni_azienda, value="all")
                    ], style={"width": "24%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Area Prassi", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-area-overview", options=opzioni_area, value="all")
                    ], style={"width": "24%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Categoria di diversità", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-categoria-overview", options=opzioni_categoria, value="all")
                    ], style={"width": "24%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Anno", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-anno-overview", options=opzioni_anno, value="all")
                    ], style={"width": "24%", "display": "inline-block"})
                ], style={"margin": "20px 0"}),
                html.Div([
                    html.Div([
                        html.H3("Numero totale di aziende", style={"color": "#080808"}),
                        html.H4(id="kpi-aziende-overview", style={"color": "#c30c13"})
                    ], style={"width": "30%", "display": "inline-block", "textAlign": "center",
                              "border": "2px solid #c30c13", "padding": "10px", "borderRadius": "5px", "margin": "10px"}),
                    html.Div([
                        html.H3("Numero totale di iniziative", style={"color": "#080808"}),
                        html.H4(id="kpi-iniziative-overview", style={"color": "#ab0404"})
                    ], style={"width": "30%", "display": "inline-block", "textAlign": "center",
                              "border": "2px solid #ab0404", "padding": "10px", "borderRadius": "5px", "margin": "10px"}),
                    html.Div([
                        html.H3("% aziende con linguaggio inclusivo", style={"color": "#080808"}),
                        html.H4(id="kpi-inclusive-overview", style={"color": "#c2222c"})
                    ], style={"width": "30%", "display": "inline-block", "textAlign": "center",
                              "border": "2px solid #c2222c", "padding": "10px", "borderRadius": "5px", "margin": "10px"})
                ], style={"margin-bottom": "50px"}),
                dcc.Graph(id="graph-distribuzione-overview")
            ], style={"fontFamily": "Arial, sans-serif", "backgroundColor": "#f7f7f7", "padding": "20px"})
        ]),
        # Tab Iniziative D&I
        dcc.Tab(label="Iniziative D&I", value="tab-initiatives", children=[
            html.Div([
                html.Div([
                    html.H1("Dashboard D&I - Sezione Iniziative D&I", style={"margin": "0", "padding": "20px", "color": "#ffffff"})
                ], style={"backgroundColor": "#c10b13", "textAlign": "center"}),
                html.Div([
                    html.Div([
                        html.Label("Azienda", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-azienda-table", options=opzioni_azienda, value="all")
                    ], style={"width": "24%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Area Prassi", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-area-table", options=opzioni_area, value="all")
                    ], style={"width": "24%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Categoria di diversità", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-categoria-table", options=opzioni_categoria, value="all")
                    ], style={"width": "24%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Anno", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-anno-table", options=opzioni_anno, value="all")
                    ], style={"width": "24%", "display": "inline-block"})
                ], style={"margin": "20px 0"}),
                dcc.Tabs(id="tabs-initiatives", value="tab-table", children=[
                    dcc.Tab(label="Tabella Iniziative", value="tab-table", children=[
                        html.Div([
                            dash_table.DataTable(
                                id="table-iniziative",
                                columns=[{"name": col, "id": col} for col in df_iniziative.columns],
                                data=df_iniziative.to_dict("records"),
                                page_size=10,
                                style_table={"overflowX": "auto"},
                                style_header={"backgroundColor": "#c10b13", "color": "white", "fontWeight": "bold"},
                                style_cell={"textAlign": "left", "padding": "5px"}
                            )
                        ], style={"margin": "20px"})
                    ]),
                    dcc.Tab(label="Grafici Iniziative", value="tab-graphs", children=[
                        html.Div([
                            dcc.Graph(id="graph-area-prassi"),
                            dcc.Graph(id="graph-categoria-diversita"),
                            dcc.Graph(id="graph-evoluzione")
                        ], style={"margin": "20px"})
                    ])
                ])
            ], style={"fontFamily": "Arial, sans-serif", "backgroundColor": "#f7f7f7", "padding": "20px"})
        ]),
        # Tab Composizione di Genere (versione funzionante di prima)
        dcc.Tab(label="Composizione di Genere", value="tab-genere", children=[
            html.Div([
                html.Div([
                    html.H1("Dashboard D&I - Composizione di Genere", style={"margin": "0", "padding": "20px", "color": "#ffffff"})
                ], style={"backgroundColor": "#c10b13", "textAlign": "center"}),
                html.Div([
                    html.Div([
                        html.Label("Azienda", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-azienda-genere", options=opzioni_azienda_genere, value="all", multi=True)
                    ], style={"width": "33%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Anno", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-anno-genere", options=opzioni_anno_genere, value="all")
                    ], style={"width": "33%", "display": "inline-block", "margin-right": "1%"}),
                    html.Div([
                        html.Label("Posizione", style={"color": "#080808", "fontWeight": "bold"}),
                        dcc.Dropdown(id="dropdown-posizione-genere", options=opzioni_posizione_genere, value="all")
                    ], style={"width": "33%", "display": "inline-block"})
                ], style={"margin": "20px 0"}),
                html.Div([
                    dash_table.DataTable(
                        id="table-genere",
                        columns=[{"name": col, "id": col} for col in df_composizione.columns if col != "Linguaggio inclusivo"],
                        data=df_composizione.to_dict("records"),
                        page_size=10,
                        style_table={"overflowX": "auto"},
                        style_header={"backgroundColor": "#c10b13", "color": "white", "fontWeight": "bold"},
                        style_cell={"textAlign": "left", "padding": "5px"}
                    )
                ], style={"margin": "20px"}),
                html.Div([
                    dcc.Graph(id="graph-bar-genere"),
                    dcc.Graph(id="graph-pie-genere")
                ], style={"margin": "20px"})
            ], style={"fontFamily": "Arial, sans-serif", "backgroundColor": "#f7f7f7", "padding": "20px"})
        ]),
        # Tab Indicatori sintetici (D&I Index) - ultima tab a destra
        dcc.Tab(label="Indicatori sintetici", value="tab-index", children=[
    html.Div([
        html.Div([
            html.H1("Dashboard D&I - Indicatori sintetici (D&I Index)", style={"margin": "0", "padding": "20px", "color": "#ffffff"})
        ], style={"backgroundColor": "#c10b13", "textAlign": "center"}),

        html.Div([
            html.Label("Modalità visualizzazione:", style={"color": "#080808", "fontWeight": "bold"}),
            dcc.Dropdown(
                id="dropdown-index-mode",
                options=[
                    {"label": "Aggregato (tutti gli anni)", "value": "aggregato"},
                    {"label": "Per anno", "value": "anno"}
                ],
                value="aggregato"
            ),

            html.Div(id="div-dropdown-index-year", children=[
                html.Label("Seleziona Anno:", style={"color": "#080808", "fontWeight": "bold"}),
                dcc.Dropdown(id="dropdown-index-year", options=opzioni_anno, value="all")
            ], style={"display": "none", "marginTop": "10px"}),

            html.Div([
                html.Label("Seleziona Azienda:", style={"color": "#080808", "fontWeight": "bold"}),
                dcc.Dropdown(id="dropdown-index-azienda", options=opzioni_azienda, value="all")
            ], style={"marginTop": "10px"})
        ], style={"width": "60%", "margin": "20px auto"}),

        html.Div([
            dcc.Graph(id="graph-index"),

            dash_table.DataTable(
                id="table-index",
                columns=[
                    {"name": "Nome azienda", "id": "Nome azienda"},
                    {"name": "Indice Diversità Finale", "id": "Indice diversità finale"}
                ],
                data=[],
                sort_action="native",
                style_table={"overflowX": "auto"},
                style_header={"backgroundColor": "#c10b13", "color": "white", "fontWeight": "bold"},
                style_cell={"textAlign": "left", "padding": "5px"}
            ),

            html.Br(),
            html.H3("Indice Iniziative", style={"color": "#080808"}),
            dcc.Graph(id="graph-indice-iniziative"),

            html.H3("Indice Categorie di diversità", style={"color": "#080808"}),
            dcc.Graph(id="graph-indice-categorie"),

            html.H3("Indice Parità di genere (Board)", style={"color": "#080808"}),
            dcc.Graph(id="graph-indice-genere")
        ], style={"margin": "20px"})
    ], style={"fontFamily": "Arial, sans-serif", "backgroundColor": "#f7f7f7", "padding": "20px"})
])
        ])
    ])

# --- Callback per aggiornare la sezione Panoramica ---
@app.callback(
    [Output("kpi-aziende-overview", "children"),
     Output("kpi-iniziative-overview", "children"),
     Output("kpi-inclusive-overview", "children"),
     Output("graph-distribuzione-overview", "figure")],
    [Input("dropdown-azienda-overview", "value"),
     Input("dropdown-area-overview", "value"),
     Input("dropdown-categoria-overview", "value"),
     Input("dropdown-anno-overview", "value")]
)
def update_overview(azienda, area, categoria, anno):
    df_filtered = df_iniziative.copy()
    if azienda != "all":
        df_filtered = df_filtered[df_filtered["Nome azienda"] == azienda]
    if area != "all":
        df_filtered = df_filtered[df_filtered["Area Prassi"] == area]
    if categoria != "all":
        df_filtered = df_filtered[df_filtered["Categoria di diversità"] == categoria]
    if anno != "all":
        df_filtered = df_filtered[df_filtered["Anno"] == anno]
    
    num_aziende_filtered = df_filtered["Nome azienda"].nunique()
    num_iniziative_filtered = df_filtered["Titolo dell'attività"].count()
    
    aziende_filtered = df_filtered["Nome azienda"].unique()
    df_comp_filtered = df_composizione[df_composizione["Nome azienda"].isin(aziende_filtered)]
    if anno != "all":
        df_comp_filtered = df_comp_filtered[df_comp_filtered["Anno"] == anno]
    df_inclusive_filtered = df_comp_filtered[df_comp_filtered["Linguaggio inclusivo"] == "Sì"]
    aziende_inclusive = df_inclusive_filtered["Nome azienda"].unique()
    num_inclusive = sum(azienda in aziende_inclusive for azienda in aziende_filtered)
    perc_inclusive_filtered = (num_inclusive / len(aziende_filtered)) * 100 if len(aziende_filtered) > 0 else 0

    if df_filtered.empty:
        fig_overview = px.bar(title="Nessun dato disponibile per i filtri selezionati")
    else:
        df_dist = df_filtered.groupby("Anno")["Titolo dell'attività"].count().reset_index(name="Numero iniziative")
        fig_overview = px.bar(
            df_dist,
            x="Anno",
            y="Numero iniziative",
            title="Distribuzione delle iniziative per Anno",
            color="Numero iniziative",
            color_continuous_scale=PALETTE
        )
        fig_overview.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")
        
    return (f"{num_aziende_filtered}",
            f"{num_iniziative_filtered}",
            f"{perc_inclusive_filtered:.2f}%",
            fig_overview)

# --- Callback per aggiornare la sezione Iniziative D&I ---
@app.callback(
    [Output("table-iniziative", "data"),
     Output("graph-area-prassi", "figure"),
     Output("graph-categoria-diversita", "figure"),
     Output("graph-evoluzione", "figure")],
    [Input("dropdown-azienda-table", "value"),
     Input("dropdown-area-table", "value"),
     Input("dropdown-categoria-table", "value"),
     Input("dropdown-anno-table", "value")]
)
def update_initiatives(azienda, area, categoria, anno):
    df_filtered = df_iniziative.copy()
    if azienda != "all":
        df_filtered = df_filtered[df_filtered["Nome azienda"] == azienda]
    if area != "all":
        df_filtered = df_filtered[df_filtered["Area Prassi"] == area]
    if categoria != "all":
        df_filtered = df_filtered[df_filtered["Categoria di diversità"] == categoria]
    if anno != "all":
        df_filtered = df_filtered[df_filtered["Anno"] == anno]
    
    table_data = df_filtered.to_dict("records")
    
    if df_filtered.empty:
        fig_area = px.bar(title="Nessun dato disponibile")
    else:
        df_area = df_filtered.groupby("Area Prassi").size().reset_index(name="Count")
        fig_area = px.bar(
            df_area,
            x="Area Prassi",
            y="Count",
            title="Frequenza iniziative per Area Prassi",
            color="Count", color_continuous_scale=PALETTE
        )
        fig_area.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")
    
    if df_filtered.empty:
        fig_cat = px.pie(title="Nessun dato disponibile")
    else:
        df_cat = df_filtered.groupby("Categoria di diversità").size().reset_index(name="Count")
        fig_cat = px.pie(
            df_cat,
            names="Categoria di diversità",
            values="Count",
            title="Frequenza per Categoria di diversità",
            color_discrete_sequence=PALETTE
        )
        fig_cat.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")
        
    
    if df_filtered.empty:
        fig_evo = px.line(title="Nessun dato disponibile")
    else:
        df_evo = df_filtered.groupby("Anno").size().reset_index(name="Count")
        fig_evo = px.line(
            df_evo,
            x="Anno",
            y="Count",
            title="Evoluzione delle iniziative nel tempo",
            markers=True
        )
        fig_evo.update_traces(line_color="#c10a13", marker_color="#c10a13")
        fig_evo.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")
    
    return table_data, fig_area, fig_cat, fig_evo

# --- Callback per aggiornare la sezione Composizione di Genere ---
@app.callback(
    [Output("table-genere", "data"),
     Output("graph-bar-genere", "figure"),
     Output("graph-pie-genere", "figure")],
    [Input("dropdown-azienda-genere", "value"),
     Input("dropdown-anno-genere", "value"),
     Input("dropdown-posizione-genere", "value")]
)
def update_genere(aziende, anno, posizione):
    df_genere = df_composizione.copy()
    if aziende != "all":
        if isinstance(aziende, str):
            aziende = [aziende]
        df_genere = df_genere[df_genere["Nome azienda"].isin(aziende)]
    if anno != "all":
        df_genere = df_genere[df_genere["Anno"] == anno]
    if posizione != "all":
        df_genere = df_genere[df_genere["Posizione"] == posizione]
    
    table_data = df_genere.drop(columns=["Linguaggio inclusivo"], errors="ignore").to_dict("records")
    
    df_genere = df_genere.copy()
    try:
        df_genere["Percentuale donne"] = df_genere["Percentuale donne"].str.rstrip('%').astype(float)
        df_genere["Percentuale uomini"] = df_genere["Percentuale uomini"].str.rstrip('%').astype(float)
    except Exception:
        df_genere["Percentuale donne"] = df_genere["Percentuale donne"]
        df_genere["Percentuale uomini"] = df_genere["Percentuale uomini"]
    
    if df_genere.empty:
        fig_bar = px.bar(title="Nessun dato disponibile")
    else:
        df_bar = df_genere.groupby(["Nome azienda", "Posizione"]).agg({
            "Percentuale donne": "mean",
            "Percentuale uomini": "mean"
        }).reset_index()
        color_discrete_map = {
            "Percentuale donne": "#c30c13",
            "Percentuale uomini": "#c3c3c3"
        }
        fig_bar = px.bar(
            df_bar,
            x="Nome azienda",
            y=["Percentuale donne", "Percentuale uomini"],
            barmode="group",
            title="Confronto percentuale donne e uomini per Azienda",
            labels={"value": "Percentuale", "variable": "Genere"},
            color_discrete_map=color_discrete_map
        )
        fig_bar.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")
    
    if df_genere.empty:
        fig_pie = px.pie(title="Nessun dato disponibile")
    else:
        total_donne = df_genere["Numero donne"].sum()
        total_uomini = df_genere["Numero uomini"].sum()
        df_pie = pd.DataFrame({
            "Genere": ["Donne", "Uomini"],
            "Count": [total_donne, total_uomini]
        })
        color_discrete_map_pie = {
            "Donne": "#c30c13",
            "Uomini": "#c3c3c3"
        }
        fig_pie = px.pie(
            df_pie,
            names="Genere",
            values="Count",
            title="Composizione complessiva: Donne vs Uomini",
            hole=0.4,
            color="Genere",
            color_discrete_map=color_discrete_map_pie
        )
        fig_pie.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")
    
    return table_data, fig_bar, fig_pie

# --- Callback per aggiornare la sezione Indicatori sintetici (D&I Index) ---
# AGGIORNA LA CALLBACK update_index QUI SOTTO:

@app.callback(
    [
        Output("graph-index", "figure"),
        Output("table-index", "data"),
        Output("graph-indice-iniziative", "figure"),
        Output("graph-indice-categorie", "figure"),
        Output("graph-indice-genere", "figure")
    ],
    [
        Input("dropdown-index-mode", "value"),
        Input("dropdown-index-year", "value"),
        Input("dropdown-index-azienda", "value")  # NUOVO INPUT
    ]
)
def update_index(mode, year, azienda):
    # Filtro per anno
    if mode == "aggregato" or year == "all":
        df_index = df_iniziative.copy()
        df_genere = df_composizione.copy()
    else:
        df_index = df_iniziative[df_iniziative["Anno"] == year].copy()
        df_genere = df_composizione[df_composizione["Anno"] == year].copy()

    # Filtro per azienda
    if azienda != "all":
        df_index = df_index[df_index["Nome azienda"] == azienda]
        df_genere = df_genere[df_genere["Nome azienda"] == azienda]

    # Calcolo indici
    risultati = calcola_tre_indici(df_index, df_genere)
    risultati = risultati.sort_values("Indice diversità finale", ascending=False)

    # Grafico indice finale
    fig_index = px.bar(
        risultati,
        x="Nome azienda",
        y="Indice diversità finale",
        title="Indice Diversità Finale (0-100)",
        color="Indice diversità finale",
        color_continuous_scale=PALETTE
    )
    fig_index.update_layout(plot_bgcolor="#f7f7f7", paper_bgcolor="#f7f7f7", font_color="#080808")

    # Grafici singoli
    fig1, fig2, fig3, _ = crea_grafici_indici(risultati)

    # Tabella dati
    table_data = risultati[["Nome azienda", "Indice diversità finale"]].to_dict("records")

    return fig_index, table_data, fig1, fig2, fig3


# --- Callback per mostrare/nascondere il dropdown per l'anno nella sezione Indicatori sintetici ---
@app.callback(
    Output("div-dropdown-index-year", "style"),
    [Input("dropdown-index-mode", "value")]
)
def show_hide_year_dropdown(mode):

    if mode == "anno":
        return {"display": "block", "marginTop": "10px"}
    else:
        return {"display": "none"}

if __name__ == "__main__":
    app.run(debug=True)
