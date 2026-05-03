# ============================================================
# Oman Vision 2040 – GCC Index Dashboard
# ============================================================

library(shiny)
library(shinydashboard)
library(ggplot2)
library(dplyr)
library(tidyr)
library(plotly)
library(readxl)
library(scales)
library(stringr)
library(base64enc)
library(htmlwidgets)

`%||%` <- function(a, b) if (!is.null(a)) a else b

# ============================================================
# GLOBAL SETTINGS
# ============================================================

COUNTRY_COLORS <- c(
  "Oman"                 = "#E63946",
  "United Arab Emirates" = "#FFB703",
  "Saudi Arabia"         = "#2DC653",
  "Kuwait"               = "#1B4FBB",
  "Bahrain"              = "#9B5DE5",
  "Qatar"                = "#48CAE4"
)

COUNTRY_SYMBOLS <- c(
  "Oman"                 = "circle",
  "United Arab Emirates" = "square",
  "Saudi Arabia"         = "diamond",
  "Kuwait"               = "x",
  "Bahrain"              = "star",
  "Qatar"                = "triangle-up"
)

COUNTRY_SHAPE_LABELS <- c(
  "Oman"                 = "●",
  "United Arab Emirates" = "■",
  "Saudi Arabia"         = "◆",
  "Kuwait"               = "✕",
  "Bahrain"              = "★",
  "Qatar"                = "▲"
)

COUNTRY_ISO3 <- c(
  "Oman"                 = "OMN",
  "United Arab Emirates" = "ARE",
  "Saudi Arabia"         = "SAU",
  "Kuwait"               = "KWT",
  "Bahrain"              = "BHR",
  "Qatar"                = "QAT"
)

COUNTRY_LON <- c(
  "Oman"                 = 57.5,
  "United Arab Emirates" = 54.4,
  "Saudi Arabia"         = 45.0,
  "Kuwait"               = 47.5,
  "Bahrain"              = 50.6,
  "Qatar"                = 51.2
)

COUNTRY_LAT <- c(
  "Oman"                 = 21.5,
  "United Arab Emirates" = 24.4,
  "Saudi Arabia"         = 24.0,
  "Kuwait"               = 29.3,
  "Bahrain"              = 26.1,
  "Qatar"                = 25.3
)

COUNTRY_ALPHA_FULL  <- 1.0
COUNTRY_ALPHA_FADED <- 0.18
NO_RANK_INDICES <- c("Inflation Rate (CPI)")
TARGET_DISPLAY_MAX_YEAR <- 2026

logo_file <- if (file.exists("logo-oman2040.png")) {
  "logo-oman2040.png"
} else if (file.exists("vision2040_logo.png")) {
  "vision2040_logo.png"
} else {
  NULL
}

logo_b64 <- if (!is.null(logo_file)) {
  base64enc::dataURI(file = logo_file, mime = "image/png")
} else {
  NULL
}

logo_ui <- function(height = "38px") {
  if (!is.null(logo_b64)) {
    tags$img(src = logo_b64, height = height, style = "margin-right:8px; vertical-align:middle;")
  } else {
    tags$span("2040", style = "font-weight:800; margin-right:8px; color:#E63946;")
  }
}


resolve_file_name <- function(base_file) {
  # Prefer the newest uploaded Excel file with "(1)" when it exists.
  updated_file <- str_replace(base_file, "\\.xlsx$", "(1).xlsx")
  if (file.exists(updated_file)) return(updated_file)
  if (file.exists(base_file)) return(base_file)
  stop(paste0("Excel file not found: ", base_file, " or ", updated_file))
}


resolve_sheet_name <- function(file, requested_sheet) {
  available <- excel_sheets(file)

  # 1) exact match
  if (requested_sheet %in% available) return(requested_sheet)

  # 2) match after removing extra spaces at the beginning/end
  trimmed_available <- str_squish(available)
  trimmed_requested <- str_squish(requested_sheet)
  match_i <- which(tolower(trimmed_available) == tolower(trimmed_requested))[1]
  if (!is.na(match_i)) return(available[match_i])

  # 3) partial match for long Excel sheet names that were cut by Excel
  match_i <- which(str_detect(tolower(trimmed_available), fixed(tolower(trimmed_requested), ignore_case = TRUE)) |
                     str_detect(tolower(trimmed_requested), fixed(tolower(trimmed_available), ignore_case = TRUE)))[1]
  if (!is.na(match_i)) return(available[match_i])

  stop(paste0(
    "Sheet '", requested_sheet, "' not found in ", file,
    ". Available sheets are: ", paste(available, collapse = ", ")
  ))
}

# ============================================================
# ROBUST DATA LOADING
# This function works even when a sheet has a title row before the header.
# It also supports both Rank/Score sheets and Score-only sheets such as CPI.
# ============================================================

find_header_row <- function(file, sheet) {
  sheet <- resolve_sheet_name(file, sheet)
  raw <- read_excel(file, sheet = sheet, col_names = FALSE, .name_repair = "minimal")
  header_words <- c("country", "country name", "country / territory", "country/territory")
  row_has_country <- apply(raw, 1, function(row) {
    vals <- tolower(str_squish(as.character(row)))
    any(vals %in% header_words, na.rm = TRUE)
  })
  header_row <- which(row_has_country)[1]
  if (is.na(header_row)) header_row <- 1
  header_row
}

read_index_sheet <- function(file, sheet, index_name, pillar_name) {
  sheet <- resolve_sheet_name(file, sheet)
  header_row <- find_header_row(file, sheet)

  df <- read_excel(file, sheet = sheet, skip = header_row - 1, .name_repair = "unique")
  names(df) <- str_squish(names(df))

  country_col <- names(df)[tolower(names(df)) %in% c("country", "country name", "country / territory", "country/territory")][1]
  if (is.na(country_col)) stop(paste("Country column not found in", sheet))

  df <- df %>%
    rename(Country = all_of(country_col)) %>%
    mutate(Country = str_squish(as.character(Country))) %>%
    filter(!is.na(Country), Country != "") %>%
    filter(!str_detect(Country, "^https?://")) %>%
    filter(Country %in% names(COUNTRY_COLORS))

  # Keep only real year/rank/score columns before pivoting.
  # This avoids errors from text columns such as "2030 Target" or notes columns.
  data_cols <- names(df)[names(df) != "Country"]
  valid_metric_cols <- data_cols[
    str_detect(data_cols, regex("rank|score", ignore_case = TRUE)) |
      str_detect(str_squish(data_cols), "^\\d{4}$")
  ]

  if (length(valid_metric_cols) == 0) {
    return(tibble(Pillar = character(), Index = character(), Country = character(),
                  Year = integer(), Rank = numeric(), Score = numeric()))
  }

  long <- df %>%
    select(Country, all_of(valid_metric_cols)) %>%
    pivot_longer(
      cols = -Country,
      names_to = "Metric",
      values_to = "Value",
      values_transform = list(Value = as.character)
    ) %>%
    mutate(
      Metric = str_squish(Metric),
      Year = as.integer(str_extract(Metric, "\\d{4}")),
      Type = case_when(
        str_detect(Metric, regex("rank", ignore_case = TRUE)) ~ "Rank",
        str_detect(Metric, regex("score", ignore_case = TRUE)) ~ "Score",
        str_detect(Metric, "^\\d{4}$") ~ "Score",  # for Inflation CPI columns: 2020, 2021, etc.
        TRUE ~ NA_character_
      ),
      Value = str_replace_all(Value, ",", ""),
      Value = suppressWarnings(as.numeric(Value))
    ) %>%
    filter(!is.na(Year), !is.na(Type), !is.na(Value)) %>%
    select(Country, Year, Type, Value) %>%
    pivot_wider(names_from = Type, values_from = Value) %>%
    mutate(
      Rank = if ("Rank" %in% names(.)) Rank else NA_real_,
      Score = if ("Score" %in% names(.)) Score else NA_real_,
      Index = index_name,
      Pillar = pillar_name
    ) %>%
    select(Pillar, Index, Country, Year, Rank, Score)

  long
}

parse_target_rank <- function(target_value) {
  x <- str_squish(as.character(target_value))
  if (is.na(x) || x == "" || tolower(x) == "none") return(NA_real_)
  if (str_detect(x, regex("top", ignore_case = TRUE))) {
    val <- str_extract(x, "\\d+\\.?\\d*")
    return(suppressWarnings(as.numeric(val)))
  }
  NA_real_
}

parse_target_score_low <- function(target_value) {
  x <- str_squish(as.character(target_value))
  if (is.na(x) || x == "" || tolower(x) == "none") return(NA_real_)
  if (str_detect(x, regex("top", ignore_case = TRUE))) return(NA_real_)
  nums <- suppressWarnings(as.numeric(str_extract_all(x, "\\d+\\.?\\d*")[[1]]))
  if (length(nums) == 0) return(NA_real_)
  min(nums, na.rm = TRUE)
}

parse_target_score_high <- function(target_value) {
  x <- str_squish(as.character(target_value))
  if (is.na(x) || x == "" || tolower(x) == "none") return(NA_real_)
  if (str_detect(x, regex("top", ignore_case = TRUE))) return(NA_real_)
  nums <- suppressWarnings(as.numeric(str_extract_all(x, "\\d+\\.?\\d*")[[1]]))
  if (length(nums) == 0) return(NA_real_)
  max(nums, na.rm = TRUE)
}

read_target_sheet <- function(file, sheet, index_name, pillar_name) {
  sheet <- resolve_sheet_name(file, sheet)
  header_row <- find_header_row(file, sheet)
  df <- read_excel(file, sheet = sheet, skip = header_row - 1, .name_repair = "unique")
  names(df) <- str_squish(names(df))

  country_col <- names(df)[tolower(names(df)) %in% c("country", "country name", "country / territory", "country/territory")][1]
  if (is.na(country_col)) {
    return(tibble(Pillar = character(), Index = character(), Country = character(),
                  TargetYear = integer(), TargetText = character(), TargetRank = numeric(),
                  TargetScoreLow = numeric(), TargetScoreHigh = numeric()))
  }

  df <- df %>%
    rename(Country = all_of(country_col)) %>%
    mutate(Country = str_squish(as.character(Country))) %>%
    filter(!is.na(Country), Country != "", Country %in% names(COUNTRY_COLORS))

  target_cols <- names(df)[str_detect(names(df), regex("2030\\s*target|2040\\s*target", ignore_case = TRUE))]
  if (length(target_cols) == 0) {
    return(tibble(Pillar = character(), Index = character(), Country = character(),
                  TargetYear = integer(), TargetText = character(), TargetRank = numeric(),
                  TargetScoreLow = numeric(), TargetScoreHigh = numeric()))
  }

  df %>%
    select(Country, all_of(target_cols)) %>%
    pivot_longer(-Country, names_to = "TargetColumn", values_to = "TargetText", values_transform = list(TargetText = as.character)) %>%
    mutate(
      TargetText = str_squish(TargetText),
      TargetYear = as.integer(str_extract(TargetColumn, "\\d{4}")),
      TargetRank = sapply(TargetText, parse_target_rank),
      TargetScoreLow = sapply(TargetText, parse_target_score_low),
      TargetScoreHigh = sapply(TargetText, parse_target_score_high),
      Index = index_name,
      Pillar = pillar_name
    ) %>%
    filter(!is.na(TargetYear), !is.na(TargetText), TargetText != "") %>%
    select(Pillar, Index, Country, TargetYear, TargetText, TargetRank, TargetScoreLow, TargetScoreHigh)
}

load_targets <- function() {
  bind_rows(
    read_target_sheet(resolve_file_name("Vision_ECONOMY & DEVELOPMENT PILLAR.xlsx"), "Inflation Rate (CPI)", "Inflation Rate (CPI)", "Economy & Development"),
    read_target_sheet(resolve_file_name("Vision_ECONOMY & DEVELOPMENT PILLAR.xlsx"), "Economic Complexity Index", "Economic Complexity Index", "Economy & Development"),
    read_target_sheet(resolve_file_name("Vision_ECONOMY & DEVELOPMENT PILLAR.xlsx"), "Economic Freedom Index ", "Economic Freedom Index", "Economy & Development"),
    read_target_sheet(resolve_file_name("Vision_ECONOMY & DEVELOPMENT PILLAR.xlsx"), "Legatum Prosperity Index (Overa", "Legatum Prosperity Index (Overall)", "Economy & Development"),
    read_target_sheet(resolve_file_name("Vision_GOVERNANCE & INSTITUTIONAL PERFORMANCE PILLAR.xlsx"), "E-Government Development Index ", "E-Government Development Index", "Governance & Institutional Performance"),
    read_target_sheet(resolve_file_name("Vision_GOVERNANCE & INSTITUTIONAL PERFORMANCE PILLAR.xlsx"), "Corruption Perceptions Index ", "Corruption Perceptions Index", "Governance & Institutional Performance"),
    read_target_sheet(resolve_file_name("Vision_PEOPLE & SOCIETY PILLAR.xlsx"), "Global Innovation Index", "Global Innovation Index", "People & Society"),
    read_target_sheet(resolve_file_name("Vision_PEOPLE & SOCIETY PILLAR.xlsx"), "Global Talent Competitiveness", "Global Talent Competitiveness Index", "People & Society"),
    read_target_sheet(resolve_file_name("Vision_PEOPLE & SOCIETY PILLAR.xlsx"), "Legatum Prosperity – Social Cap", "Legatum Prosperity Index – Social Capital", "People & Society"),
    read_target_sheet(resolve_file_name("Vision_PEOPLE & SOCIETY PILLAR.xlsx"), "Legatum Prosperity Index – Heal", "Legatum Health Index", "People & Society"),
    read_target_sheet(resolve_file_name("Vision_PEOPLE & SOCIETY PILLAR.xlsx"), "Social Progress Index", "Social Progress Index", "People & Society"),
    read_target_sheet(resolve_file_name("Vision_SUSTAINABLE ENVIRONMENT PILLAR.xlsx"), "Environmental Performance Index", "Environmental Performance Index", "Sustainable Environment")
  )
}


load_economy <- function() {
  file <- resolve_file_name("Vision_ECONOMY & DEVELOPMENT PILLAR.xlsx")
  bind_rows(
    read_index_sheet(file, "Inflation Rate (CPI)", "Inflation Rate (CPI)", "Economy & Development"),
    read_index_sheet(file, "Economic Complexity Index", "Economic Complexity Index", "Economy & Development"),
    read_index_sheet(file, "Economic Freedom Index ", "Economic Freedom Index", "Economy & Development"),
    read_index_sheet(file, "Legatum Prosperity Index (Overa", "Legatum Prosperity Index (Overall)", "Economy & Development")
  )
}

load_governance <- function() {
  file <- resolve_file_name("Vision_GOVERNANCE & INSTITUTIONAL PERFORMANCE PILLAR.xlsx")
  bind_rows(
    read_index_sheet(file, "E-Government Development Index ", "E-Government Development Index", "Governance & Institutional Performance"),
    read_index_sheet(file, "Corruption Perceptions Index ", "Corruption Perceptions Index", "Governance & Institutional Performance")
  )
}

load_people <- function() {
  file <- resolve_file_name("Vision_PEOPLE & SOCIETY PILLAR.xlsx")
  bind_rows(
    read_index_sheet(file, "Global Innovation Index", "Global Innovation Index", "People & Society"),
    read_index_sheet(file, "Global Talent Competitiveness", "Global Talent Competitiveness Index", "People & Society"),
    read_index_sheet(file, "Legatum Prosperity – Social Cap", "Legatum Prosperity Index – Social Capital", "People & Society"),
    read_index_sheet(file, "Legatum Prosperity Index – Heal", "Legatum Health Index", "People & Society"),
    read_index_sheet(file, "Social Progress Index", "Social Progress Index", "People & Society")
  )
}

load_environment <- function() {
  file <- resolve_file_name("Vision_SUSTAINABLE ENVIRONMENT PILLAR.xlsx")
  read_index_sheet(file, "Environmental Performance Index", "Environmental Performance Index", "Sustainable Environment")
}

all_data <- bind_rows(
  load_economy(),
  load_governance(),
  load_people(),
  load_environment()
) %>%
  filter(!is.na(Country), Country %in% names(COUNTRY_COLORS)) %>%
  arrange(Pillar, Index, Country, Year)

all_targets <- load_targets()

COUNTRIES <- sort(unique(all_data$Country))
INDEX_CHOICES <- sort(unique(all_data$Index))

index_has_rank <- function(index_name) {
  any(all_data$Index == index_name & !is.na(all_data$Rank))
}

get_alphas <- function(highlight_country) {
  alphas <- setNames(rep(COUNTRY_ALPHA_FADED, length(COUNTRY_COLORS)), names(COUNTRY_COLORS))
  alphas["Oman"] <- COUNTRY_ALPHA_FULL
  if (!is.null(highlight_country) && highlight_country %in% names(alphas)) {
    alphas[highlight_country] <- COUNTRY_ALPHA_FULL
  }
  alphas
}


empty_plot <- function(message = "No data available for this chart") {
  plotly_empty(type = "scatter", mode = "markers") %>%
    layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(visible = FALSE), yaxis = list(visible = FALSE),
      annotations = list(list(
        text = message, x = 0.5, y = 0.5, xref = "paper", yref = "paper",
        showarrow = FALSE, font = list(size = 14, color = "#9baab8")
      ))
    )
}

# ============================================================
# UI
# ============================================================

header <- dashboardHeader(
  title = tags$span(
    logo_ui("38px"),
    tags$span("Oman Vision 2040", style = "font-weight:700; font-size:16px; color:#fff; vertical-align:middle;")
  ),
  titleWidth = 260
)

sidebar <- dashboardSidebar(
  width = 260,
  sidebarMenu(
    id = "sidebar",
    menuItem("Overview", tabName = "home", icon = icon("home")),
    menuItem("Economy & Development", tabName = "economy", icon = icon("chart-line")),
    menuItem("Governance", tabName = "governance", icon = icon("landmark")),
    menuItem("People & Society", tabName = "people", icon = icon("users")),
    menuItem("Sustainable Environment", tabName = "environment", icon = icon("leaf"))
  ),
  tags$div(
    style = "padding: 12px 14px 8px 14px; border-top:1px solid #1e2535;",
    tags$p("Country Highlight Mode", style = "color:#aaa; font-size:12px; margin-bottom:6px; text-transform:uppercase; letter-spacing:1px;"),
    radioButtons("highlight_mode", label = NULL,
                 choices = c("Only Oman" = "oman", "All countries" = "all"),
                 selected = "oman")
  ),
  tags$div(
    style = "position:absolute; bottom:16px; left:0; right:0; text-align:center;",
    tags$p("INFS 4475 · SQU · SP2026", style = "color:#555; font-size:11px;")
  )
)

custom_css <- tags$style(HTML("
  body, .content-wrapper, .main-footer { background: #0f1117 !important; }
  .skin-blue .main-header .logo { background:#161b27 !important; border-bottom:1px solid #1e2535; }
  .skin-blue .main-header .navbar { background:#161b27 !important; border-bottom:1px solid #1e2535; }
  .skin-blue .main-sidebar { background:#0d1120 !important; }
  .skin-blue .sidebar-menu > li > a { color:#c0c8d8 !important; font-size:13px; }
  .skin-blue .sidebar-menu > li.active > a, .skin-blue .sidebar-menu > li > a:hover { background:#1e2a42 !important; color:#fff !important; }
  .skin-blue .sidebar-menu > li.active > a { border-left:3px solid #E63946 !important; }
  .box { background:#161b27 !important; border:1px solid #1e2535 !important; border-radius:10px !important; }
  .box-header { background:#161b27 !important; border-bottom:1px solid #1e2535 !important; border-radius:10px 10px 0 0 !important; }
  .box-title { color:#e0e7f0 !important; font-weight:600 !important; font-size:14px !important; }
  .info-box { border-radius:10px !important; background:#161b27 !important; border:1px solid #1e2535; }
  .info-box-icon { border-radius:10px 0 0 10px !important; }
  .info-box-content { color:#e0e7f0 !important; }
  .info-box-number { color:#fff !important; font-size:22px !important; }
  .info-box-text { color:#9baab8 !important; font-size:12px !important; }
  h3.tab-title { color:#e0e7f0; font-size:22px; font-weight:700; margin-bottom:4px; }
  p.tab-sub { color:#7a8a9e; font-size:13px; margin-bottom:18px; }
  .selectize-input { background:#1e2535 !important; border:1px solid #2e3a50 !important; color:#e0e7f0 !important; border-radius:6px !important; }
  .selectize-dropdown { background:#1e2535 !important; border:1px solid #2e3a50 !important; color:#e0e7f0 !important; }
  .country-legend-dot { display:inline-block; width:12px; height:12px; border-radius:50%; margin-right:6px; }
  .legend-row { display:inline-flex; align-items:center; margin-right:16px; font-size:12px; color:#c0c8d8; }
  .chart-note { font-size:11px; color:#5a6a7e; margin-top:4px; font-style:italic; }
  .overview-card { background:#161b27; border:1px solid #1e2535; border-radius:12px; padding:20px; margin-bottom:16px; }
  .overview-card h4 { color:#e0e7f0; margin-bottom:8px; font-weight:600; }
  .overview-card p { color:#7a8a9e; font-size:13px; line-height:1.6; }
  .pillar-badge { display:inline-block; padding:3px 10px; border-radius:20px; font-size:11px; font-weight:600; margin-bottom:10px; }
  .fullscreen-plot-btn, .data-download-btn { float:right; margin:6px 6px 0 0; background:#1e2535 !important; color:#c0c8d8 !important; border:1px solid #2e3a50 !important; border-radius:6px !important; }
  .fullscreen-plot-btn:hover, .data-download-btn:hover { background:#2e3a50 !important; color:#ffffff !important; text-decoration:none !important; }
  .data-download-btn { padding:6px 12px !important; line-height:20px !important; }
  #plot-fullscreen-modal { display:none; position:fixed; z-index:99999; left:0; top:0; width:100vw; height:100vh; background:rgba(5,8,15,0.94); padding:22px; box-sizing:border-box; }
  #plot-fullscreen-panel { width:96vw; height:90vh; max-width:1500px; margin:0 auto; background:#161b27; border:1px solid #2e3a50; border-radius:12px; display:flex; flex-direction:column; box-shadow:0 0 30px rgba(0,0,0,0.45); }
  #plot-fullscreen-header { min-height:48px; padding:10px 14px; border-bottom:1px solid #2e3a50; display:flex; align-items:center; justify-content:space-between; color:#e0e7f0; font-weight:700; font-size:16px; box-sizing:border-box; }
  #plot-fullscreen-body { flex:1; min-height:0; padding:18px; overflow:hidden; box-sizing:border-box; }
  #plot-fullscreen-clone { width:100% !important; height:100% !important; min-height:0 !important; }
  #plot-fullscreen-close { background:#E63946 !important; color:#fff !important; border:0 !important; border-radius:8px !important; padding:6px 12px !important; font-size:15px !important; }
  .control-label { color:#9baab8 !important; }
  .radio label { color:#c0c8d8 !important; font-weight:400 !important; }
"))

country_legend_ui <- function() {
  tags$div(style = "margin-bottom: 12px; padding: 10px 14px; background:#0d1120; border-radius:8px; display:flex; flex-wrap:wrap; gap:4px;",
    lapply(names(COUNTRY_COLORS), function(cn) {
      tags$span(class = "legend-row",
        tags$span(style = paste0("color:", COUNTRY_COLORS[cn], "; font-size:18px; font-weight:800; margin-right:7px; line-height:1;"),
                  COUNTRY_SHAPE_LABELS[[cn]]),
        cn
      )
    })
  )
}

plot_action_button <- function(output_id) {
  tagList(
    tags$button("Full Screen", class = "fullscreen-plot-btn", `data-plot-id` = output_id, type = "button"),
    downloadButton(paste0("download_", output_id), "Download Data", class = "data-download-btn")
  )
}

fullscreen_plot_script <- tags$script(HTML("
  var fullscreenPlotState = null;

  function ensureFullscreenModal() {
    if ($('#plot-fullscreen-modal').length) return;
    $('body').append(`
      <div id='plot-fullscreen-modal'>
        <div id='plot-fullscreen-panel'>
          <div id='plot-fullscreen-header'>
            <span id='plot-fullscreen-title'>Full screen visualization</span>
            <button id='plot-fullscreen-close' type='button'>Close</button>
          </div>
          <div id='plot-fullscreen-body'><div id='plot-fullscreen-clone'></div></div>
        </div>
      </div>
    `);
  }

  function getPlotSize() {
    var body = $('#plot-fullscreen-body');
    return {
      width: Math.max(420, body.innerWidth() - 6),
      height: Math.max(360, body.innerHeight() - 6)
    };
  }

  function stripSizing(layout) {
    var clean = $.extend(true, {}, layout || {});
    delete clean.width;
    delete clean.height;
    clean.autosize = true;

    // Give the fullscreen copy enough internal space so axis titles and tick labels
    // do not get pushed outside the visible modal.
    clean.margin = $.extend(true, {}, clean.margin || {});
    clean.margin.l = Math.max(clean.margin.l || 0, 85);
    clean.margin.r = Math.max(clean.margin.r || 0, 90);
    clean.margin.t = Math.max(clean.margin.t || 0, 55);
    clean.margin.b = Math.max(clean.margin.b || 0, 120);

    if (clean.legend) {
      clean.legend = $.extend(true, {}, clean.legend);
      clean.legend.x = 1.02;
      clean.legend.y = 1;
      clean.legend.xanchor = 'left';
      clean.legend.yanchor = 'top';
    }
    return clean;
  }

  function drawFullscreenClone(sourcePlot) {
    var clone = document.getElementById('plot-fullscreen-clone');
    if (!clone || !sourcePlot) return;

    var size = getPlotSize();
    $(clone).css({
      width: size.width + 'px',
      height: size.height + 'px',
      display: 'block'
    });

    var dataCopy = $.extend(true, [], sourcePlot.data || []);
    var layoutCopy = stripSizing(sourcePlot.layout || {});
    layoutCopy.width = size.width;
    layoutCopy.height = size.height;
    layoutCopy.autosize = true;

    var configCopy = $.extend(true, {}, sourcePlot._context || {});
    configCopy.responsive = true;
    configCopy.displayModeBar = true;

    Plotly.react(clone, dataCopy, layoutCopy, configCopy).then(function() {
      Plotly.Plots.resize(clone);
    });
  }

  function closeFullscreenPlot() {
    if (!fullscreenPlotState) return;
    var sourcePlot = fullscreenPlotState.sourcePlot;

    var clone = document.getElementById('plot-fullscreen-clone');
    if (clone) Plotly.purge(clone);
    $('#plot-fullscreen-modal').hide();

    setTimeout(function() {
      if (sourcePlot) {
        Plotly.relayout(sourcePlot, {autosize: true});
        Plotly.Plots.resize(sourcePlot);
      }
      $('.selectized').each(function(){ if (this.selectize) this.selectize.close(); });
    }, 80);

    fullscreenPlotState = null;
  }

  $(document).on('click', '.fullscreen-plot-btn', function() {
    ensureFullscreenModal();
    if (fullscreenPlotState) closeFullscreenPlot();

    var id = $(this).data('plot-id');
    var sourcePlot = document.getElementById(id);
    if (!sourcePlot || !sourcePlot.data) return;

    $('.selectized').each(function(){ if (this.selectize) this.selectize.close(); });

    var boxTitle = $(this).closest('.box').find('.box-title').first().text().trim() || 'Full screen visualization';
    $('#plot-fullscreen-title').text(boxTitle);

    fullscreenPlotState = { sourcePlot: sourcePlot };
    $('#plot-fullscreen-modal').show();

    setTimeout(function(){ drawFullscreenClone(sourcePlot); }, 50);
    setTimeout(function(){ drawFullscreenClone(sourcePlot); }, 250);
  });

  $(document).on('click', '#plot-fullscreen-close', function() {
    closeFullscreenPlot();
  });

  $(document).on('keyup', function(e) {
    if (e.key === 'Escape') closeFullscreenPlot();
  });

  $(window).on('resize', function() {
    if (fullscreenPlotState && fullscreenPlotState.sourcePlot) {
      drawFullscreenClone(fullscreenPlotState.sourcePlot);
    }
  });
"))

rank_condition <- function(tab_id) {
  paste0("input.", tab_id, "_index != 'Inflation Rate (CPI)'")
}

pillar_tab_body <- function(tab_id, title, subtitle, badge_color, index_names) {
  tabItem(tabName = tab_id,
    tags$h3(class = "tab-title", title),
    tags$p(class = "tab-sub", subtitle),
    tags$div(class = "pillar-badge",
             style = paste0("background:", badge_color, "22; color:", badge_color, "; border:1px solid ", badge_color, "44;"),
             title),
    country_legend_ui(),
    fluidRow(
      column(12,
        tags$div(style = "background:#0d1120; border-radius:8px; padding:12px 16px; margin-bottom:16px;",
          tags$label("Select Index:", style="color:#9baab8; font-size:12px; display:block; margin-bottom:4px;"),
          selectInput(paste0(tab_id, "_index"), NULL, choices = index_names, width = "100%")
        )
      )
    ),

    conditionalPanel(
      condition = rank_condition(tab_id),
      fluidRow(
        box(width = 6, title = "📊 Rank Over Time", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_rank_line"), height = "320px"),
            plot_action_button(paste0(tab_id, "_rank_line")),
            tags$p(class = "chart-note", "Lower rank = better performance. Target lines are shown only up to 2026 to keep the chart readable.")),
        box(width = 6, title = "🏆 Score Comparison (Latest Year)", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_score_bar"), height = "320px"),
            plot_action_button(paste0(tab_id, "_score_bar")),
            tags$p(class = "chart-note", "Horizontal bar chart shows latest available scores per country."))
      ),
      fluidRow(
        box(width = 6, title = "🎯 Score Over Time", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_score_area"), height = "320px"),
            plot_action_button(paste0(tab_id, "_score_area")),
            tags$p(class = "chart-note", "Shows score trajectory over available years.")),
        box(width = 6, title = "🗺️ Rank Heatmap", solidHeader = TRUE,
            tags$div(style = "padding: 6px 6px 0 6px;",
              selectInput(paste0(tab_id, "_heatmap_palette"), "Heatmap colour mode:",
                          choices = c("Normal" = "normal",
                                      "Deuteranopia-friendly" = "deuteranopia",
                                      "Protanopia-friendly" = "protanopia",
                                      "Tritanopia-friendly" = "tritanopia"),
                          selected = "normal", width = "100%")
            ),
            plotlyOutput(paste0(tab_id, "_rank_heatmap"), height = "270px"),
            plot_action_button(paste0(tab_id, "_rank_heatmap")),
            tags$p(class = "chart-note", "Heatmap: lower rank = better. Choose a colour mode that is easiest for you to read."))
      ),
      fluidRow(
        box(width = 12, title = "🔵 Score Bubble Chart", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_bubble"), height = "320px"),
            plot_action_button(paste0(tab_id, "_bubble")),
            tags$p(class = "chart-note", "Two-variable view: Score and Rank in the latest available year."))
      ),
      fluidRow(
        box(width = 12, title = "🔮 Score Forecast: Next Two Years", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_score_forecast"), height = "360px"),
            plot_action_button(paste0(tab_id, "_score_forecast")),
            tags$p(class = "chart-note", "Simple linear forecast based on available past scores. Forecast values are estimates, not official targets."))
      )
    ),

    conditionalPanel(
      condition = paste0("!(", rank_condition(tab_id), ")"),
      fluidRow(
        box(width = 6, title = "🏆 Score Comparison (Latest Year)", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_score_bar_scoreonly"), height = "320px"),
            plot_action_button(paste0(tab_id, "_score_bar_scoreonly")),
            tags$p(class = "chart-note", "This index has scores only, so rank visuals were removed.")),
        box(width = 6, title = "🎯 Score Over Time", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_score_area_scoreonly"), height = "320px"),
            plot_action_button(paste0(tab_id, "_score_area_scoreonly")),
            tags$p(class = "chart-note", "Score-only trend chart for this index."))
      ),
      fluidRow(
        box(width = 12, title = "🔮 Score Forecast: Next Two Years", solidHeader = TRUE,
            plotlyOutput(paste0(tab_id, "_score_forecast_scoreonly"), height = "360px"),
            plot_action_button(paste0(tab_id, "_score_forecast_scoreonly")),
            tags$p(class = "chart-note", "Simple linear forecast based on available past scores. Forecast values are estimates, not official targets."))
      )
    ),

    fluidRow(
      box(width = 6, title = "📉 Oman Gap vs GCC Average", solidHeader = TRUE,
          plotlyOutput(paste0(tab_id, "_oman_gap"), height = "320px"),
          plot_action_button(paste0(tab_id, "_oman_gap")),
          tags$p(class = "chart-note", "Analytical comparison: Oman score minus GCC average score over time.")),
      box(width = 6, title = "⚖️ Oman vs Highlight Country", solidHeader = TRUE,
          tags$div(style = "padding: 6px 6px 0 6px;",
                   selectInput(paste0(tab_id, "_compare_country"), "Compare Oman with:",
                               choices = setdiff(COUNTRIES, "Oman"),
                               selected = setdiff(COUNTRIES, "Oman")[1], width = "100%")
          ),
          plotlyOutput(paste0(tab_id, "_highlight_gap"), height = "270px"),
          plot_action_button(paste0(tab_id, "_highlight_gap")),
          tags$p(class = "chart-note", "Analytical comparison using two variables: score difference and year. The selected country is also highlighted with Oman in the other charts on this tab."))
    )
  )
}

body <- dashboardBody(
  custom_css,
  fullscreen_plot_script,
  tabItems(
    tabItem(tabName = "home",
      fluidRow(
        column(12,
          tags$div(style = "text-align:center; padding: 30px 0 10px 0;",
            logo_ui("80px"),
            tags$h2("Oman Vision 2040 · GCC Index Dashboard", style = "color:#e0e7f0; font-weight:700; margin-top:12px;"),
            tags$p("Tracking Oman's progress across GCC nations on key development indices", style = "color:#7a8a9e; font-size:15px;")
          )
        )
      ),
      fluidRow(
        infoBoxOutput("home_ib1", width = 3),
        infoBoxOutput("home_ib2", width = 3),
        infoBoxOutput("home_ib3", width = 3),
        infoBoxOutput("home_ib4", width = 3)
      ),
      fluidRow(column(12, country_legend_ui())),
      fluidRow(
        box(width = 6, title = "🕸️ Radar Chart: Countries vs GCC Average", solidHeader = TRUE,
            selectizeInput("home_radar_countries", "Countries for radar comparison:",
                           choices = COUNTRIES, selected = "Oman", multiple = TRUE,
                           options = list(plugins = list("remove_button")), width = "100%"),
            plotlyOutput("home_radar_chart", height = "420px"),
            plot_action_button("home_radar_chart"),
            tags$p(class = "chart-note", "Radar chart compares one or more selected countries with the GCC average using the latest values from the Excel sheets for key Vision 2040 indices.")),
        box(width = 6, title = "🗺️ GCC Map Chart", solidHeader = TRUE,
            selectInput("home_map_index", "Map index:", choices = INDEX_CHOICES, selected = "Economic Freedom Index", width = "100%"),
            plotlyOutput("home_map_chart", height = "420px"),
            plot_action_button("home_map_chart"),
            tags$p(class = "chart-note", "Map chart shows the latest available score by GCC country for the selected index."))
      ),
      fluidRow(
        box(width = 12, title = "📈 Oman Score Trajectory Across All Pillars",
            plotlyOutput("home_overview_plot", height = "400px"),
            plot_action_button("home_overview_plot"),
            tags$p(class = "chart-note", "Scores for selected indices: Oman versus GCC average."))
      ),
      fluidRow(
        box(width = 6, title = "🏅 Latest Rankings: Oman vs GCC",
            plotlyOutput("home_rank_bubble", height = "430px"),
            plot_action_button("home_rank_bubble"),
            tags$p(class = "chart-note", "Each point shows the latest global rank. Rank numbers are on the X-axis and indexes are on the Y-axis. Lower rank is better. Score-only CPI is excluded.")),
        box(width = 6, title = "📊 GCC Score Distribution by Pillar",
            plotlyOutput("home_boxplot", height = "380px"),
            plot_action_button("home_boxplot"),
            tags$p(class = "chart-note", "Score spread within each pillar across GCC countries."))
      ),
      fluidRow(
        box(width = 12, title = "🔍 Analytical Two-Index Comparison", solidHeader = TRUE,
          fluidRow(
            column(6, selectInput("compare_x", "X-axis index", choices = INDEX_CHOICES, selected = "Economic Freedom Index")),
            column(6, selectInput("compare_y", "Y-axis index", choices = INDEX_CHOICES, selected = "E-Government Development Index"))
          ),
          plotlyOutput("home_two_index_compare", height = "420px"),
          plot_action_button("home_two_index_compare"),
          tags$p(class = "chart-note", "Analytical visualization with two variables: latest raw score in Index X vs latest raw score in Index Y. Each point is a GCC country."))
      ),
      fluidRow(
        box(width = 12, title = "🫧 Bubble Chart: E-Government × Innovation × Economic Freedom", solidHeader = TRUE,
            plotlyOutput("home_gov_innovation_freedom_bubble", height = "460px"),
            plot_action_button("home_gov_innovation_freedom_bubble"),
            tags$p(class = "chart-note", "3-variable chart: X-axis = E-Government score, Y-axis = Global Innovation score, and bubble size = Economic Freedom score. It helps show whether digital governance and innovation move together, and whether economic freedom is associated with either."))
      ),
      fluidRow(
        box(width = 12, title = "📌 Latest Oman Score Gap by Index", solidHeader = TRUE,
            plotlyOutput("home_oman_gap_all", height = "420px"),
            plot_action_button("home_oman_gap_all"),
            tags$p(class = "chart-note", "Positive value means Oman is above the GCC average for that index; negative means below."))
      )
    ),

    pillar_tab_body("economy", "Economy & Development",
      "Tracking economic complexity, freedom, prosperity and inflation across GCC",
      "#FFB703",
      c("Economic Complexity Index", "Economic Freedom Index", "Legatum Prosperity Index (Overall)", "Inflation Rate (CPI)")),

    pillar_tab_body("governance", "Governance & Institutional Performance",
      "Measuring e-government effectiveness and anti-corruption performance",
      "#2DC653",
      c("E-Government Development Index", "Corruption Perceptions Index")),

    pillar_tab_body("people", "People & Society",
      "Human capital, innovation, social progress, social capital and health across the GCC",
      "#9B5DE5",
      c("Global Innovation Index", "Global Talent Competitiveness Index", "Legatum Prosperity Index – Social Capital", "Legatum Health Index", "Social Progress Index")),

    pillar_tab_body("environment", "Sustainable Environment",
      "Environmental performance and sustainability metrics across GCC nations",
      "#48CAE4",
      c("Environmental Performance Index"))
  )
)

ui <- dashboardPage(header, sidebar, body, skin = "blue", title = "Oman Vision 2040 Dashboard")

# ============================================================
# SERVER
# ============================================================

server <- function(input, output, session) {

  get_compare_country <- function(tab_id) {
    input[[paste0(tab_id, "_compare_country")]] %||% setdiff(COUNTRIES, "Oman")[1]
  }


  get_pillar_index_data <- function(pillar_name, index_name) {
    all_data %>% filter(Pillar == pillar_name, Index == index_name, !is.na(Year))
  }

  safe_file_name <- function(x) {
    x <- gsub("[^A-Za-z0-9_]+", "_", x)
    x <- gsub("_+", "_", x)
    gsub("^_|_$", "", x)
  }

  register_data_download <- function(output_id, data_fun) {
    output[[paste0("download_", output_id)]] <- downloadHandler(
      filename = function() {
        paste0(safe_file_name(output_id), "_data_", Sys.Date(), ".csv")
      },
      content = function(file) {
        d <- data_fun()
        if (is.null(d) || nrow(as.data.frame(d)) == 0) {
          d <- data.frame(Message = "No data available for this visualization")
        }
        write.csv(as.data.frame(d), file, row.names = FALSE, na = "")
      }
    )
  }

  score_forecast_dataset <- function(df) {
    df <- df %>% filter(!is.na(Score), !is.na(Year))
    if (nrow(df) == 0) return(tibble())

    hist <- df %>%
      select(Pillar, Index, Country, Year, Score) %>%
      mutate(DataType = "Actual")

    forecast <- df %>%
      select(Pillar, Index, Country, Year, Score) %>%
      arrange(Country, Year) %>%
      group_by(Pillar, Index, Country) %>%
      group_modify(function(d, key) {
        d <- d %>% arrange(Year)
        max_year <- max(d$Year, na.rm = TRUE)
        future_years <- (max_year + 1):(max_year + 2)
        if (nrow(d) >= 2 && length(unique(d$Year)) >= 2) {
          model <- lm(Score ~ Year, data = d)
          pred <- as.numeric(predict(model, newdata = data.frame(Year = future_years)))
        } else {
          pred <- rep(tail(d$Score, 1), 2)
        }
        tibble(Year = future_years, Score = pred)
      }) %>%
      ungroup() %>%
      mutate(DataType = "Forecast")

    bind_rows(hist, forecast) %>% arrange(Index, Country, Year, DataType)
  }

  oman_gap_dataset <- function(df) {
    df %>%
      filter(!is.na(Score), !is.na(Year)) %>%
      group_by(Pillar, Index, Year) %>%
      summarise(
        Oman = Score[Country == "Oman"][1],
        GCC_Avg = mean(Score, na.rm = TRUE),
        Gap = Oman - GCC_Avg,
        .groups = "drop"
      ) %>%
      filter(!is.na(Oman))
  }

  highlight_gap_dataset <- function(df, hc) {
    if (is.null(hc) || hc == "Oman") return(tibble())
    df %>%
      filter(!is.na(Score), !is.na(Year), Country %in% c("Oman", hc)) %>%
      select(Pillar, Index, Country, Year, Score) %>%
      pivot_wider(names_from = Country, values_from = Score) %>%
      mutate(ComparisonCountry = hc, Gap = .data[["Oman"]] - .data[[hc]]) %>%
      filter(!is.na(Gap))
  }

  build_colors <- function(hc) {
    if (!is.null(input$highlight_mode) && input$highlight_mode == "all") {
      return(list(colors = COUNTRY_COLORS, alphas = setNames(rep(1, length(COUNTRY_COLORS)), names(COUNTRY_COLORS))))
    }
    list(colors = COUNTRY_COLORS, alphas = get_alphas(hc))
  }

  latest_by_country_index <- function(df) {
    df %>%
      filter(!is.na(Score)) %>%
      group_by(Country, Index) %>%
      arrange(desc(Year)) %>%
      slice(1) %>%
      ungroup()
  }

  make_rank_line <- function(df, hc) {
    ca <- build_colors(hc)
    df <- df %>% filter(!is.na(Rank), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot("This index does not contain rank data."))

    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      d <- df %>% filter(Country == cn) %>% arrange(Year)
      if (nrow(d) == 0) next
      alpha <- ca$alphas[cn]
      col <- ca$colors[cn]
      lw <- ifelse(cn == "Oman" || cn == hc, 3.5, 1.5)
      ms <- ifelse(cn == "Oman" || cn == hc, 9, 6)
      p <- add_trace(
        p,
        data = d,
        x = ~Year,
        y = ~Rank,
        type = "scatter",
        mode = "lines+markers",
        name = cn,
        line = list(color = col, width = lw),
        marker = list(color = col, size = ms, symbol = COUNTRY_SYMBOLS[[cn]]),
        opacity = alpha,
        hovertemplate = paste0("<b>", cn, "</b><br>Year: %{x}<br>Rank: %{y}<extra></extra>")
      )
    }

    p <- add_rank_targets(p, df)

    p <- p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "Year", gridcolor = "#1e2535", color = "#7a8a9e", tickformat = "d",
                   range = c(min(df$Year, na.rm = TRUE) - 0.2, TARGET_DISPLAY_MAX_YEAR + 0.2), dtick = 1),
      yaxis = list(title = "Global Rank", gridcolor = "#1e2535", color = "#7a8a9e", autorange = "reversed"),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 11)),
      margin = list(l = 50, r = 20, t = 20, b = 50),
      hovermode = "closest",
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
    default_mode <- input$highlight_mode %||% "oman"
    onRender(p, sprintf(
      "function(el, x) {
        var gd = document.getElementById(el.id);
        var mode = '%s';
        function applyDefault(){
          var op = [], lw = [], ms = [];
          gd.data.forEach(function(tr){
            var keep = (mode === 'all' || tr.name === 'Oman' || tr.name === '%s');
            op.push(keep ? 1.0 : 0.18);
            lw.push((tr.name === 'Oman' || tr.name === '%s') ? 3.5 : 1.5);
            ms.push((tr.name === 'Oman' || tr.name === '%s') ? 9 : 6);
          });
          Plotly.restyle(gd, {'opacity': op, 'line.width': lw, 'marker.size': ms});
        }
        gd.on('plotly_hover', function(e){
          var hovered = e.points[0].data.name;
          var op = [], lw = [], ms = [];
          gd.data.forEach(function(tr){
            var keep = (tr.name === 'Oman' || tr.name === hovered);
            op.push(keep ? 1.0 : 0.10);
            lw.push(keep ? 4.2 : 1.2);
            ms.push(keep ? 10 : 5);
          });
          Plotly.restyle(gd, {'opacity': op, 'line.width': lw, 'marker.size': ms});
        });
        gd.on('plotly_unhover', function(){ applyDefault(); });
        applyDefault();
      }", default_mode, hc, hc, hc)
    )
  }

  make_score_bar <- function(df, hc) {
    ca <- build_colors(hc)
    df <- df %>% filter(!is.na(Score))
    if (nrow(df) == 0) return(empty_plot())

    latest_year <- max(df$Year, na.rm = TRUE)
    d <- df %>% filter(Year == latest_year) %>% arrange(Score) %>% mutate(Country = factor(Country, levels = Country))
    colors_vec <- sapply(as.character(d$Country), function(cn) scales::alpha(COUNTRY_COLORS[cn], ca$alphas[cn]))

    plot_ly(d, x = ~Score, y = ~Country, type = "bar", orientation = "h",
            marker = list(color = colors_vec, line = list(color = "#0d1120", width = 1)),
            hovertemplate = "<b>%{y}</b><br>Score: %{x:.2f}<extra></extra>") %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
        font = list(color = "#c0c8d8", family = "Arial"),
        xaxis = list(title = paste0("Score (", latest_year, ")"), gridcolor = "#1e2535", color = "#7a8a9e"),
        yaxis = list(title = "", color = "#7a8a9e"),
        margin = list(l = 140, r = 20, t = 20, b = 50),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  }

  make_score_area <- function(df, hc) {
    ca <- build_colors(hc)
    df <- df %>% filter(!is.na(Score), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot())

    score_range <- range(df$Score, na.rm = TRUE)
    score_pad <- max(diff(score_range) * 0.18, 1)
    y_min <- max(0, score_range[1] - score_pad)
    y_max <- score_range[2] + score_pad

    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      d <- df %>% filter(Country == cn) %>% arrange(Year)
      if (nrow(d) == 0) next
      alpha <- ca$alphas[cn]
      col <- ca$colors[cn]
      lw <- ifelse(cn == "Oman" || cn == hc, 3, 1.5)
      ms <- ifelse(cn == "Oman" || cn == hc, 8, 5)
      p <- add_trace(p, data = d, x = ~Year, y = ~Score, type = "scatter",
                     mode = "lines+markers",
                     name = cn,
                     line = list(color = col, width = lw),
                     marker = list(color = col, size = ms, symbol = COUNTRY_SYMBOLS[[cn]]),
                     opacity = alpha,
                     hovertemplate = paste0("<b>", cn, "</b><br>Year: %{x}<br>Score: %{y:.2f}<extra></extra>"))
    }
    p <- add_score_targets(p, df)
    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "Year", gridcolor = "#1e2535", color = "#7a8a9e", tickformat = "d"),
      yaxis = list(title = "Score", gridcolor = "#1e2535", color = "#7a8a9e", range = c(y_min, y_max)),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 11)),
      margin = list(l = 50, r = 20, t = 20, b = 50),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  }

  make_rank_heatmap <- function(df, palette_mode = "normal") {
    df <- df %>% filter(!is.na(Rank), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot("This index does not contain rank data."))

    wide <- df %>% select(Country, Year, Rank) %>% pivot_wider(names_from = Year, values_from = Rank) %>% arrange(Country)
    countries <- wide$Country
    years_cols <- sort(as.integer(names(wide)[-1]))
    mat <- as.matrix(wide[, as.character(years_cols), drop = FALSE])
    rownames(mat) <- countries

    heatmap_scale <- switch(
      palette_mode,
      "deuteranopia" = list(list(0.00, "#003F5C"), list(0.50, "#7A5195"), list(1.00, "#FFA600")),
      "protanopia"   = list(list(0.00, "#003F5C"), list(0.50, "#58508D"), list(1.00, "#FFBC42")),
      "tritanopia"   = list(list(0.00, "#1B1B3A"), list(0.50, "#6D597A"), list(1.00, "#F4A261")),
      list(list(0.00, "#2DC653"), list(0.50, "#FFB703"), list(1.00, "#E63946"))
    )

    plot_ly(x = as.character(years_cols), y = countries, z = mat,
            type = "heatmap",
            colorscale = heatmap_scale,
            reversescale = FALSE,
            hovertemplate = "<b>%{y}</b><br>Year: %{x}<br>Rank: %{z}<extra></extra>",
            colorbar = list(title = "Rank", titlefont = list(color = "#c0c8d8"), tickfont = list(color = "#c0c8d8"))) %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
        font = list(color = "#c0c8d8", family = "Arial"),
        xaxis = list(title = "Year", color = "#7a8a9e"),
        yaxis = list(title = "", color = "#7a8a9e"),
        margin = list(l = 130, r = 80, t = 20, b = 50),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  }

  make_bubble <- function(df, hc) {
    ca <- build_colors(hc)
    df <- df %>% filter(!is.na(Score), !is.na(Rank), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot("This index does not contain both score and rank data."))

    latest_year <- max(df$Year, na.rm = TRUE)
    d <- df %>% filter(Year == latest_year) %>%
      mutate(BubbleSize = scales::rescale(abs(Score), to = c(12, 40)))

    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      dd <- d %>% filter(Country == cn)
      if (nrow(dd) == 0) next
      alpha <- ca$alphas[cn]
      p <- add_trace(
        p, data = dd, x = ~Rank, y = ~Score,
        type = "scatter", mode = "markers",
        name = cn,
        marker = list(
          color = COUNTRY_COLORS[[cn]],
          symbol = COUNTRY_SYMBOLS[[cn]],
          size = dd$BubbleSize,
          sizemode = "diameter",
          opacity = alpha,
          line = list(color = "#0d1120", width = 1.5)
        ),
        hovertemplate = paste0("<b>", cn, "</b><br>Rank: %{x}<br>Score: %{y:.2f}<extra></extra>")
      )
    }

    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = paste0("Rank (", latest_year, ") — lower is better"), gridcolor = "#1e2535", color = "#7a8a9e", autorange = "reversed"),
      yaxis = list(title = "Score", gridcolor = "#1e2535", color = "#7a8a9e"),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 11)),
      margin = list(l = 60, r = 20, t = 20, b = 60),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  }

  make_oman_gap <- function(df) {
    df <- df %>% filter(!is.na(Score), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot())

    d <- df %>%
      group_by(Year) %>%
      summarise(
        Oman = Score[Country == "Oman"][1],
        GCC_Avg = mean(Score, na.rm = TRUE),
        .groups = "drop"
      ) %>%
      filter(!is.na(Oman)) %>%
      mutate(Gap = Oman - GCC_Avg)

    plot_ly(d, x = ~Year, y = ~Gap, type = "scatter", mode = "lines+markers",
            line = list(color = "#E63946", width = 3),
            marker = list(color = "#E63946", size = 8),
            hovertemplate = "<b>Oman vs GCC Avg</b><br>Year: %{x}<br>Gap: %{y:.2f}<extra></extra>") %>%
      add_trace(data = d, x = ~Year, y = rep(0, nrow(d)), type = "scatter", mode = "lines",
                line = list(color = "#7a8a9e", width = 1, dash = "dot"), name = "Zero gap", showlegend = FALSE,
                hoverinfo = "skip") %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
        font = list(color = "#c0c8d8", family = "Arial"),
        xaxis = list(title = "Year", gridcolor = "#1e2535", color = "#7a8a9e", tickformat = "d"),
        yaxis = list(title = "Score gap", gridcolor = "#1e2535", color = "#7a8a9e"),
        showlegend = FALSE,
        margin = list(l = 60, r = 20, t = 20, b = 50),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  }

  make_highlight_gap <- function(df, hc) {
    df <- df %>% filter(!is.na(Score), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot())
    if (is.null(hc) || hc == "Oman") return(empty_plot("Choose a comparison country from the dropdown above."))
    if (!(hc %in% unique(df$Country))) return(empty_plot("The selected country is not available for this index."))
    if (!("Oman" %in% unique(df$Country))) return(empty_plot("Oman is not available for this index."))

    d <- df %>%
      filter(Country %in% c("Oman", hc)) %>%
      select(Country, Year, Score) %>%
      pivot_wider(names_from = Country, values_from = Score) %>%
      mutate(Gap = .data[["Oman"]] - .data[[hc]]) %>%
      filter(!is.na(Gap))

    if (nrow(d) == 0) return(empty_plot("No overlapping years for Oman and selected country."))

    plot_ly(d, x = ~Year, y = ~Gap, type = "bar",
            marker = list(color = ifelse(d$Gap >= 0, scales::alpha("#2DC653", 0.8), scales::alpha("#E63946", 0.8))),
            hovertemplate = paste0("<b>Oman - ", hc, "</b><br>Year: %{x}<br>Score difference: %{y:.2f}<extra></extra>")) %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
        font = list(color = "#c0c8d8", family = "Arial"),
        xaxis = list(title = "Year", gridcolor = "#1e2535", color = "#7a8a9e", tickformat = "d"),
        yaxis = list(title = paste0("Oman score - ", hc, " score"), gridcolor = "#1e2535", color = "#7a8a9e"),
        margin = list(l = 70, r = 20, t = 20, b = 50),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  }


  add_rank_targets <- function(p, df) {
    idx <- unique(df$Index)[1]
    if (is.na(idx)) return(p)
    targets <- all_targets %>% filter(Index == idx, Country == "Oman", !is.na(TargetRank)) %>% arrange(TargetYear)
    if (nrow(targets) == 0) return(p)
    x_min <- min(df$Year, na.rm = TRUE)
    x_max <- min(TARGET_DISPLAY_MAX_YEAR, max(c(df$Year, TARGET_DISPLAY_MAX_YEAR), na.rm = TRUE))
    for (i in seq_len(nrow(targets))) {
      p <- add_trace(p,
                     x = c(x_min, x_max), y = c(targets$TargetRank[i], targets$TargetRank[i]),
                     type = "scatter", mode = "lines",
                     name = paste0(targets$TargetYear[i], " Target: ", targets$TargetText[i]),
                     line = list(color = "#FFFFFF", width = 2, dash = ifelse(targets$TargetYear[i] == 2030, "dash", "dot")),
                     opacity = 0.85,
                     hovertemplate = paste0("<b>Oman ", targets$TargetYear[i], " target</b><br>", targets$TargetText[i], "<extra></extra>"))
    }
    p
  }

  add_score_targets <- function(p, df) {
    idx <- unique(df$Index)[1]
    if (is.na(idx)) return(p)
    targets <- all_targets %>% filter(Index == idx, Country == "Oman", !is.na(TargetScoreLow)) %>% arrange(TargetYear)
    if (nrow(targets) == 0) return(p)
    x_min <- min(df$Year, na.rm = TRUE)
    x_max <- min(TARGET_DISPLAY_MAX_YEAR, max(c(df$Year, TARGET_DISPLAY_MAX_YEAR), na.rm = TRUE))
    for (i in seq_len(nrow(targets))) {
      low <- targets$TargetScoreLow[i]
      high <- targets$TargetScoreHigh[i]
      p <- add_trace(p,
                     x = c(x_min, x_max), y = c(low, low),
                     type = "scatter", mode = "lines",
                     name = paste0(targets$TargetYear[i], " Target Low: ", targets$TargetText[i]),
                     line = list(color = "#FFFFFF", width = 2, dash = "dash"),
                     opacity = 0.9,
                     hovertemplate = paste0("<b>Oman ", targets$TargetYear[i], " target</b><br>", targets$TargetText[i], "<extra></extra>"))
      if (!is.na(high) && high != low) {
        p <- add_trace(p,
                       x = c(x_min, x_max), y = c(high, high),
                       type = "scatter", mode = "lines",
                       name = paste0(targets$TargetYear[i], " Target High: ", targets$TargetText[i]),
                       line = list(color = "#FFFFFF", width = 2, dash = "dot"),
                       opacity = 0.9,
                       hovertemplate = paste0("<b>Oman ", targets$TargetYear[i], " target</b><br>", targets$TargetText[i], "<extra></extra>"))
      }
    }
    p
  }

  make_score_forecast <- function(df, hc = "Oman") {
    ca <- build_colors(hc)
    df <- df %>% filter(!is.na(Score), !is.na(Year))
    if (nrow(df) == 0) return(empty_plot("No score data available to forecast."))

    hist <- df %>% arrange(Country, Year)
    forecast <- hist %>%
      group_by(Country) %>%
      group_modify(function(d, key) {
        d <- d %>% arrange(Year)
        max_year <- max(d$Year, na.rm = TRUE)
        future_years <- (max_year + 1):(max_year + 2)
        if (nrow(d) >= 2 && length(unique(d$Year)) >= 2) {
          model <- lm(Score ~ Year, data = d)
          pred <- as.numeric(predict(model, newdata = data.frame(Year = future_years)))
        } else {
          pred <- rep(tail(d$Score, 1), 2)
        }
        tibble(Year = future_years, Score = pred)
      }) %>%
      ungroup()

    score_range <- range(c(hist$Score, forecast$Score), na.rm = TRUE)
    score_pad <- max(diff(score_range) * 0.18, 1)
    y_min <- score_range[1] - score_pad
    y_max <- score_range[2] + score_pad

    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      dh <- hist %>% filter(Country == cn) %>% arrange(Year)
      dfc <- forecast %>% filter(Country == cn) %>% arrange(Year)
      if (nrow(dh) == 0) next
      alpha <- ca$alphas[cn]
      col <- COUNTRY_COLORS[[cn]]
      lw <- ifelse(cn == "Oman" || cn == hc, 3, 1.5)
      ms <- ifelse(cn == "Oman" || cn == hc, 8, 5)
      p <- add_trace(p, data = dh, x = ~Year, y = ~Score, type = "scatter",
                     mode = "lines+markers", name = paste0(cn, " actual"),
                     legendgroup = cn,
                     line = list(color = col, width = lw),
                     marker = list(color = col, size = ms, symbol = COUNTRY_SYMBOLS[[cn]]),
                     opacity = alpha,
                     hovertemplate = paste0("<b>", cn, " actual</b><br>Year: %{x}<br>Score: %{y:.2f}<extra></extra>"))
      if (nrow(dfc) > 0) {
        bridge <- tibble(Year = c(tail(dh$Year, 1), dfc$Year), Score = c(tail(dh$Score, 1), dfc$Score))
        p <- add_trace(p, data = bridge, x = ~Year, y = ~Score, type = "scatter",
                       mode = "lines+markers", name = paste0(cn, " forecast"),
                       legendgroup = cn, showlegend = FALSE,
                       line = list(color = col, width = lw, dash = "dash"),
                       marker = list(color = col, size = ms, symbol = COUNTRY_SYMBOLS[[cn]]),
                       opacity = alpha,
                       hovertemplate = paste0("<b>", cn, " forecast</b><br>Year: %{x}<br>Forecast score: %{y:.2f}<extra></extra>"))
      }
    }
    p <- add_score_targets(p, df)
    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "Year", gridcolor = "#1e2535", color = "#7a8a9e", tickformat = "d"),
      yaxis = list(title = "Score", gridcolor = "#1e2535", color = "#7a8a9e", range = c(y_min, y_max)),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 10), orientation = "h", y = -0.22),
      margin = list(l = 60, r = 20, t = 20, b = 95),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  }

  register_pillar <- function(tab_id, pillar_name) {
    observeEvent(input[[paste0(tab_id, "_index")]], {
      df_current <- get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]])
      choices <- sort(setdiff(unique(df_current$Country[!is.na(df_current$Score)]), "Oman"))
      if (length(choices) == 0) {
        updateSelectInput(session, paste0(tab_id, "_compare_country"), choices = character(0), selected = character(0))
      } else {
        updateSelectInput(session, paste0(tab_id, "_compare_country"), choices = choices, selected = choices[1])
      }
    }, ignoreInit = FALSE)

    output[[paste0(tab_id, "_rank_line")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_rank_line(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_score_bar")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_score_bar(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_score_area")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_score_area(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_rank_heatmap")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      req(input[[paste0(tab_id, "_heatmap_palette")]])
      make_rank_heatmap(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), input[[paste0(tab_id, "_heatmap_palette")]])
    })
    output[[paste0(tab_id, "_bubble")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_bubble(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_score_bar_scoreonly")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_score_bar(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_score_area_scoreonly")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_score_area(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_score_forecast")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_score_forecast(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_score_forecast_scoreonly")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_score_forecast(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), "Oman")
    })
    output[[paste0(tab_id, "_oman_gap")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      make_oman_gap(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]))
    })
    output[[paste0(tab_id, "_highlight_gap")]] <- renderPlotly({
      req(input[[paste0(tab_id, "_index")]])
      req(input[[paste0(tab_id, "_compare_country")]])
      make_highlight_gap(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), input[[paste0(tab_id, "_compare_country")]])
    })

    # CSV downloads for every pillar visualization
    register_data_download(paste0(tab_id, "_rank_line"), function() {
      req(input[[paste0(tab_id, "_index")]])
      get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Rank))
    })
    register_data_download(paste0(tab_id, "_score_bar"), function() {
      req(input[[paste0(tab_id, "_index")]])
      df <- get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Score))
      latest_year <- max(df$Year, na.rm = TRUE)
      df %>% filter(Year == latest_year)
    })
    register_data_download(paste0(tab_id, "_score_area"), function() {
      req(input[[paste0(tab_id, "_index")]])
      get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Score))
    })
    register_data_download(paste0(tab_id, "_rank_heatmap"), function() {
      req(input[[paste0(tab_id, "_index")]])
      get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Rank))
    })
    register_data_download(paste0(tab_id, "_bubble"), function() {
      req(input[[paste0(tab_id, "_index")]])
      df <- get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Score), !is.na(Rank))
      latest_year <- max(df$Year, na.rm = TRUE)
      df %>% filter(Year == latest_year)
    })
    register_data_download(paste0(tab_id, "_score_forecast"), function() {
      req(input[[paste0(tab_id, "_index")]])
      score_forecast_dataset(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]))
    })
    register_data_download(paste0(tab_id, "_score_bar_scoreonly"), function() {
      req(input[[paste0(tab_id, "_index")]])
      df <- get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Score))
      latest_year <- max(df$Year, na.rm = TRUE)
      df %>% filter(Year == latest_year)
    })
    register_data_download(paste0(tab_id, "_score_area_scoreonly"), function() {
      req(input[[paste0(tab_id, "_index")]])
      get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]) %>% filter(!is.na(Score))
    })
    register_data_download(paste0(tab_id, "_score_forecast_scoreonly"), function() {
      req(input[[paste0(tab_id, "_index")]])
      score_forecast_dataset(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]))
    })
    register_data_download(paste0(tab_id, "_oman_gap"), function() {
      req(input[[paste0(tab_id, "_index")]])
      oman_gap_dataset(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]))
    })
    register_data_download(paste0(tab_id, "_highlight_gap"), function() {
      req(input[[paste0(tab_id, "_index")]])
      req(input[[paste0(tab_id, "_compare_country")]])
      highlight_gap_dataset(get_pillar_index_data(pillar_name, input[[paste0(tab_id, "_index")]]), input[[paste0(tab_id, "_compare_country")]])
    })
  }

  register_pillar("economy", "Economy & Development")
  register_pillar("governance", "Governance & Institutional Performance")
  register_pillar("people", "People & Society")
  register_pillar("environment", "Sustainable Environment")

  latest_oman_rank <- function(index_name) {
    d <- all_data %>% filter(Index == index_name, Country == "Oman", !is.na(Rank)) %>% arrange(desc(Year)) %>% slice(1)
    if (nrow(d) == 0) return(list(value = "N/A", subtitle = "No rank"))
    list(value = paste0("Rank #", d$Rank[1]), subtitle = paste0(d$Year[1]))
  }

  output$home_ib1 <- renderInfoBox({
    d <- latest_oman_rank("Economic Freedom Index")
    infoBox("Economic Freedom", d$value, subtitle = d$subtitle, icon = icon("chart-line"), color = "yellow", fill = TRUE)
  })
  output$home_ib2 <- renderInfoBox({
    d <- latest_oman_rank("E-Government Development Index")
    infoBox("E-Government", d$value, subtitle = d$subtitle, icon = icon("landmark"), color = "green", fill = TRUE)
  })
  output$home_ib3 <- renderInfoBox({
    d <- latest_oman_rank("Global Innovation Index")
    infoBox("Innovation", d$value, subtitle = d$subtitle, icon = icon("lightbulb"), color = "purple", fill = TRUE)
  })
  output$home_ib4 <- renderInfoBox({
    d <- latest_oman_rank("Environmental Performance Index")
    infoBox("Environment", d$value, subtitle = d$subtitle, icon = icon("leaf"), color = "teal", fill = TRUE)
  })

  output$home_radar_chart <- renderPlotly({
    req(input$home_radar_countries)
    key_indices <- c("Economic Freedom Index", "E-Government Development Index", "Global Innovation Index", "Environmental Performance Index")

    # Use raw scores — no scaling
    df <- all_data %>%
      filter(Index %in% key_indices, !is.na(Score), !is.na(Year)) %>%
      group_by(Country, Index) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup()

    selected_countries <- input$home_radar_countries
    if (length(selected_countries) == 0) selected_countries <- "Oman"

    labels <- c("Economic Freedom", "E-Government", "Global Innovation", "Environment")
    index_lookup <- setNames(labels, key_indices)

    # Determine max raw score per index for the radial axis ceiling
    max_score <- df %>%
      group_by(Index) %>% summarise(mx = max(Score, na.rm = TRUE), .groups = "drop") %>%
      summarise(overall_max = max(mx)) %>% pull(overall_max)
    # Round up to nearest 10 for a clean axis
    radar_max <- ceiling(max_score / 10) * 10

    d_gcc <- df %>%
      group_by(Index) %>% summarise(Score = mean(Score, na.rm = TRUE), .groups = "drop") %>%
      mutate(Label = index_lookup[Index]) %>%
      arrange(match(Index, key_indices))

    p <- plot_ly(type = "scatterpolar")

    for (cn in selected_countries) {
      d_country <- df %>%
        filter(Country == cn) %>%
        select(Index, Score, Year) %>%
        right_join(tibble(Index = key_indices), by = "Index") %>%
        mutate(Label = index_lookup[Index], Score = ifelse(is.na(Score), 0, Score)) %>%
        arrange(match(Index, key_indices))

      theta_country <- c(d_country$Label, d_country$Label[1])
      r_country <- c(d_country$Score, d_country$Score[1])

      p <- add_trace(
        p, r = r_country, theta = theta_country,
        mode = "lines+markers", fill = "toself",
        name = cn,
        line = list(color = COUNTRY_COLORS[[cn]], width = ifelse(cn == "Oman", 3.5, 2.4)),
        marker = list(color = COUNTRY_COLORS[[cn]], symbol = COUNTRY_SYMBOLS[[cn]], size = ifelse(cn == "Oman", 8, 7)),
        fillcolor = scales::alpha(COUNTRY_COLORS[[cn]], ifelse(cn == "Oman", 0.22, 0.12)),
        hovertemplate = paste0("<b>", cn, "</b><br>%{theta}: %{r:.2f}<extra></extra>")
      )
    }

    theta_gcc <- c(d_gcc$Label, d_gcc$Label[1])
    r_gcc <- c(d_gcc$Score, d_gcc$Score[1])

    p %>%
      add_trace(r = r_gcc, theta = theta_gcc, mode = "lines+markers", fill = "toself",
                name = "GCC Average",
                line = list(color = "#c0c8d8", width = 2.2, dash = "dot"),
                marker = list(color = "#c0c8d8", size = 7),
                fillcolor = scales::alpha("#c0c8d8", 0.10),
                hovertemplate = "<b>GCC Average</b><br>%{theta}: %{r:.2f}<extra></extra>") %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
        font = list(color = "#c0c8d8", family = "Arial"),
        polar = list(
          bgcolor = "#161b27",
          radialaxis = list(visible = TRUE, range = c(0, radar_max), gridcolor = "#2e3a50", color = "#7a8a9e",
                            tickfont = list(size = 10)),
          angularaxis = list(gridcolor = "#2e3a50", color = "#c0c8d8", tickfont = list(size = 11))
        ),
        legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1),
        margin = list(l = 80, r = 80, t = 40, b = 60),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  })

  output$home_map_chart <- renderPlotly({
    req(input$home_map_index)
    d <- all_data %>%
      filter(Index == input$home_map_index, !is.na(Score)) %>%
      group_by(Country) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      mutate(
        ISO3 = COUNTRY_ISO3[Country],
        Lon = COUNTRY_LON[Country],
        Lat = COUNTRY_LAT[Country],
        ScoreScaled = scales::rescale(Score, to = c(14, 36))
      ) %>%
      filter(!is.na(ISO3), !is.na(Lon), !is.na(Lat))

    if (nrow(d) == 0) return(empty_plot("No map data available for this index."))

    plot_geo(d, lon = ~Lon, lat = ~Lat) %>%
      add_markers(
        text = ~paste0("<b>", Country, "</b><br>", input$home_map_index, "<br>Year: ", Year, "<br>Score: ", round(Score, 2)),
        hoverinfo = "text",
        marker = list(
          size = d$ScoreScaled,
          color = d$Score,
          colorscale = list(list(0, "#2A9D8F"), list(0.5, "#E9C46A"), list(1, "#E76F51")),
          showscale = TRUE,
          colorbar = list(title = "Score", titlefont = list(color = "#c0c8d8"), tickfont = list(color = "#c0c8d8")),
          line = list(color = "#0d1120", width = 1.5),
          opacity = 0.9
        )
      ) %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
        font = list(color = "#c0c8d8", family = "Arial"),
        geo = list(
          scope = "asia",
          projection = list(type = "mercator"),
          center = list(lon = 50, lat = 24),
          lonaxis = list(range = c(38, 62)),
          lataxis = list(range = c(14, 32)),
          bgcolor = "#161b27",
          showland = TRUE, landcolor = "#1e2535",
          showcountries = TRUE, countrycolor = "#7a8a9e",
          showocean = TRUE, oceancolor = "#0d1120",
          showframe = FALSE
        ),
        margin = list(l = 10, r = 10, t = 10, b = 10),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  })

  output$home_overview_plot <- renderPlotly({
    key_indices <- c("Economic Freedom Index", "E-Government Development Index", "Global Innovation Index", "Environmental Performance Index")
    df <- all_data %>%
      filter(Index %in% key_indices, !is.na(Score), !is.na(Year))

    p <- plot_ly()
    pillar_colors <- c(
      "Economic Freedom Index" = "#FFB703",
      "E-Government Development Index" = "#2DC653",
      "Global Innovation Index" = "#9B5DE5",
      "Environmental Performance Index" = "#48CAE4"
    )
    for (idx in key_indices) {
      d_oman <- df %>% filter(Index == idx, Country == "Oman") %>% arrange(Year)
      d_gcc  <- df %>% filter(Index == idx) %>% group_by(Year) %>% summarise(Score = mean(Score, na.rm = TRUE), .groups = "drop") %>% arrange(Year)
      col <- pillar_colors[idx]
      short <- str_remove_all(idx, " Index| Development")
      p <- add_trace(p, data = d_oman, x = ~Year, y = ~Score, type = "scatter", mode = "lines+markers",
                     name = paste0("Oman – ", short), line = list(color = col, width = 3), marker = list(color = "#E63946", size = 8),
                     hovertemplate = paste0("<b>Oman – ", idx, "</b><br>Year: %{x}<br>Score: %{y:.2f}<extra></extra>"))
      p <- add_trace(p, data = d_gcc, x = ~Year, y = ~Score, type = "scatter", mode = "lines",
                     name = paste0("GCC Avg – ", short), line = list(color = col, width = 1.5, dash = "dot"), opacity = 0.5,
                     hovertemplate = paste0("<b>GCC Avg – ", idx, "</b><br>Year: %{x}<br>Score: %{y:.2f}<extra></extra>"))
    }
    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27", font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "Year", gridcolor = "#1e2535", color = "#7a8a9e", tickformat = "d"),
      yaxis = list(title = "Score", gridcolor = "#1e2535", color = "#7a8a9e"),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 10), orientation = "h", y = -0.25),
      margin = list(l = 60, r = 20, t = 20, b = 100),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  })

  output$home_rank_bubble <- renderPlotly({
    d <- all_data %>%
      filter(!is.na(Rank)) %>%
      group_by(Country, Index) %>%
      arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      mutate(IndexShort = case_when(
        Index == "Corruption Perceptions Index" ~ "Corruption",
        Index == "Economic Complexity Index" ~ "Complexity",
        Index == "Economic Freedom Index" ~ "Economic Freedom",
        Index == "E-Government Development Index" ~ "E-Government",
        Index == "Environmental Performance Index" ~ "Environment",
        Index == "Global Innovation Index" ~ "Innovation",
        Index == "Global Talent Competitiveness Index" ~ "Talent",
        Index == "Legatum Prosperity Index (Overall)" ~ "Prosperity",
        Index == "Legatum Prosperity Index – Social Capital" ~ "Social Capital",
        Index == "Legatum Health Index" ~ "Health",
        Index == "Social Progress Index" ~ "Social Progress",
        TRUE ~ Index
      ))

    if (nrow(d) == 0) return(empty_plot("No rank data available."))

    index_order <- d %>%
      group_by(IndexShort) %>%
      summarise(best_rank = min(Rank, na.rm = TRUE), .groups = "drop") %>%
      arrange(best_rank) %>%
      pull(IndexShort)

    d <- d %>%
      mutate(
        IndexShort = factor(IndexShort, levels = rev(index_order)),
        BubbleSize = scales::rescale(max(Rank, na.rm = TRUE) - Rank + 1, to = c(9, 22))
      )

    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      dd <- d %>% filter(Country == cn)
      if (nrow(dd) == 0) next
      p <- add_trace(
        p,
        data = dd,
        x = ~Rank,
        y = ~IndexShort,
        type = "scatter",
        mode = "markers",
        name = cn,
        marker = list(
          color = COUNTRY_COLORS[[cn]],
          symbol = COUNTRY_SYMBOLS[[cn]],
          size = ifelse(cn == "Oman", pmax(dd$BubbleSize, 18), dd$BubbleSize),
          opacity = ifelse(cn == "Oman", 1, 0.75),
          line = list(color = "#0d1120", width = 1.2)
        ),
        text = ~paste0("<b>", Country, "</b><br>Index: ", Index,
                       "<br>Latest year: ", Year,
                       "<br>Rank: #", Rank),
        hoverinfo = "text"
      )
    }

    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "Latest rank (lower = better)", gridcolor = "#1e2535", color = "#7a8a9e", rangemode = "tozero"),
      yaxis = list(title = "Index", gridcolor = "#1e2535", color = "#7a8a9e", categoryorder = "array", categoryarray = rev(index_order)),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 10), orientation = "h", y = -0.25),
      margin = list(l = 150, r = 25, t = 20, b = 115),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  })

  output$home_boxplot <- renderPlotly({
    df <- all_data %>% filter(!is.na(Score))
    pillar_colors_box <- c(
      "Economy & Development" = "#FFB703",
      "Governance & Institutional Performance" = "#2DC653",
      "People & Society" = "#9B5DE5",
      "Sustainable Environment" = "#48CAE4"
    )
    p <- plot_ly()
    for (pl in names(pillar_colors_box)) {
      d <- df %>% filter(Pillar == pl)
      short_pl <- case_when(
        pl == "Economy & Development" ~ "Economy",
        pl == "Governance & Institutional Performance" ~ "Governance",
        pl == "People & Society" ~ "People",
        pl == "Sustainable Environment" ~ "Environment",
        TRUE ~ pl
      )
      p <- add_trace(p, data = d, y = ~Score, type = "box", name = short_pl,
                     marker = list(color = pillar_colors_box[pl]), line = list(color = pillar_colors_box[pl]),
                     fillcolor = scales::alpha(pillar_colors_box[pl], 0.2),
                     hovertemplate = paste0("<b>", short_pl, "</b><br>Score: %{y:.2f}<extra></extra>"))
    }
    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27", font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "Pillar", color = "#7a8a9e", gridcolor = "#1e2535"),
      yaxis = list(title = "Score", color = "#7a8a9e", gridcolor = "#1e2535"),
      showlegend = FALSE, margin = list(l = 60, r = 20, t = 20, b = 60),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  })

  output$home_two_index_compare <- renderPlotly({
    req(input$compare_x, input$compare_y)
    if (input$compare_x == input$compare_y) return(empty_plot("Choose two different indices for comparison."))

    # Use raw scores — no scaling
    latest_scores <- all_data %>%
      filter(Index %in% c(input$compare_x, input$compare_y), !is.na(Score)) %>%
      group_by(Country, Index) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      select(Country, Index, Score, Year) %>%
      pivot_wider(names_from = Index, values_from = c(Score, Year), names_sep = "__")

    xcol <- paste0("Score__", input$compare_x)
    ycol <- paste0("Score__", input$compare_y)
    if (!(xcol %in% names(latest_scores)) || !(ycol %in% names(latest_scores))) return(empty_plot("No overlapping country data for these two indices."))

    d <- latest_scores %>%
      mutate(XScore = .data[[xcol]], YScore = .data[[ycol]]) %>%
      filter(!is.na(XScore), !is.na(YScore))
    if (nrow(d) == 0) return(empty_plot("No overlapping country data for these two indices."))

    # Build per-country traces to avoid the plotly %{text} rendering bug
    cx <- input$compare_x
    cy <- input$compare_y
    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      dd <- d %>% filter(Country == cn)
      if (nrow(dd) == 0) next
      p <- add_trace(
        p,
        data = dd,
        x = ~XScore, y = ~YScore,
        type = "scatter", mode = "markers+text",
        name = cn,
        text = cn, textposition = "top center",
        marker = list(
          color = COUNTRY_COLORS[[cn]],
          symbol = COUNTRY_SYMBOLS[[cn]],
          size = 16,
          line = list(color = "#0d1120", width = 1.5)
        ),
        hovertemplate = paste0(
          "<b>", cn, "</b><br>",
          cx, ": %{x:.2f}<br>",
          cy, ": %{y:.2f}<extra></extra>"
        )
      )
    }

    x_range <- range(d$XScore, na.rm = TRUE)
    y_range <- range(d$YScore, na.rm = TRUE)
    x_pad <- diff(x_range) * 0.15 + 1
    y_pad <- diff(y_range) * 0.15 + 1

    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27", font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = paste0(cx, " score"), gridcolor = "#1e2535", color = "#7a8a9e",
                   range = c(x_range[1] - x_pad, x_range[2] + x_pad)),
      yaxis = list(title = paste0(cy, " score"), gridcolor = "#1e2535", color = "#7a8a9e",
                   range = c(y_range[1] - y_pad, y_range[2] + y_pad)),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 11)),
      margin = list(l = 80, r = 20, t = 20, b = 80),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  })

  output$home_gov_innovation_freedom_bubble <- renderPlotly({
    needed_indices <- c(
      "E-Government Development Index",
      "Global Innovation Index",
      "Economic Freedom Index"
    )

    latest_scores <- all_data %>%
      filter(Index %in% needed_indices, !is.na(Score)) %>%
      group_by(Country, Index) %>%
      arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      select(Country, Index, Score, Year) %>%
      pivot_wider(names_from = Index, values_from = c(Score, Year), names_sep = "__")

    xcol <- "Score__E-Government Development Index"
    ycol <- "Score__Global Innovation Index"
    sizecol <- "Score__Economic Freedom Index"

    if (!(xcol %in% names(latest_scores)) || !(ycol %in% names(latest_scores)) || !(sizecol %in% names(latest_scores))) {
      return(empty_plot("Required data is missing for the three-variable bubble chart."))
    }

    d <- latest_scores %>%
      mutate(
        EGovernment = .data[[xcol]],
        Innovation = .data[[ycol]],
        EconomicFreedom = .data[[sizecol]],
        BubbleSize = scales::rescale(EconomicFreedom, to = c(18, 62))
      ) %>%
      filter(!is.na(EGovernment), !is.na(Innovation), !is.na(EconomicFreedom))

    if (nrow(d) == 0) return(empty_plot("No overlapping country data for E-Government, Innovation, and Economic Freedom."))

    p <- plot_ly()
    for (cn in names(COUNTRY_COLORS)) {
      dd <- d %>% filter(Country == cn)
      if (nrow(dd) == 0) next
      p <- add_trace(
        p,
        data = dd,
        x = ~EGovernment,
        y = ~Innovation,
        type = "scatter",
        mode = "markers+text",
        name = cn,
        text = cn,
        textposition = "top center",
        marker = list(
          color = COUNTRY_COLORS[[cn]],
          symbol = "circle",
          size = dd$BubbleSize,
          sizemode = "diameter",
          opacity = ifelse(cn == "Oman", 1, 0.82),
          line = list(color = ifelse(cn == "Oman", "#ffffff", "#0d1120"), width = ifelse(cn == "Oman", 2.2, 1.3))
        ),
        hovertemplate = paste0(
          "<b>", cn, "</b><br>",
          "E-Government score: %{x:.2f}<br>",
          "Global Innovation score: %{y:.2f}<br>",
          "Economic Freedom score: ", round(dd$EconomicFreedom[1], 2),
          "<extra></extra>"
        )
      )
    }

    x_range <- range(d$EGovernment, na.rm = TRUE)
    y_range <- range(d$Innovation, na.rm = TRUE)
    x_pad <- diff(x_range) * 0.18 + 0.5
    y_pad <- diff(y_range) * 0.18 + 0.5

    p %>% layout(
      paper_bgcolor = "#161b27", plot_bgcolor = "#161b27",
      font = list(color = "#c0c8d8", family = "Arial"),
      xaxis = list(title = "E-Government score", gridcolor = "#1e2535", color = "#7a8a9e",
                   range = c(x_range[1] - x_pad, x_range[2] + x_pad)),
      yaxis = list(title = "Global Innovation score", gridcolor = "#1e2535", color = "#7a8a9e",
                   range = c(y_range[1] - y_pad, y_range[2] + y_pad)),
      legend = list(bgcolor = "#0d1120", bordercolor = "#1e2535", borderwidth = 1, font = list(size = 10), orientation = "h", y = -0.25),
      margin = list(l = 80, r = 25, t = 20, b = 115),
      annotations = list(list(
        x = 0.01, y = 1.05, xref = "paper", yref = "paper", showarrow = FALSE,
        text = "Bubble size = Economic Freedom score",
        font = list(color = "#9baab8", size = 12), align = "left"
      )),
      hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
    )
  })

  output$home_oman_gap_all <- renderPlotly({
    d <- all_data %>%
      filter(!is.na(Score)) %>%
      group_by(Index, Country) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      group_by(Index) %>%
      summarise(
        Oman = Score[Country == "Oman"][1],
        GCC_Avg = mean(Score, na.rm = TRUE),
        Pillar = Pillar[Country == "Oman"][1],
        .groups = "drop"
      ) %>%
      filter(!is.na(Oman)) %>%
      mutate(Gap = Oman - GCC_Avg) %>%
      arrange(Gap) %>%
      mutate(Index = factor(Index, levels = Index))

    plot_ly(d, x = ~Gap, y = ~Index, type = "bar", orientation = "h",
            marker = list(color = ifelse(d$Gap >= 0, scales::alpha("#2DC653", 0.8), scales::alpha("#E63946", 0.8))),
            hovertemplate = "<b>%{y}</b><br>Oman - GCC average: %{x:.2f}<extra></extra>") %>%
      layout(
        paper_bgcolor = "#161b27", plot_bgcolor = "#161b27", font = list(color = "#c0c8d8", family = "Arial"),
        xaxis = list(title = "Latest score gap", gridcolor = "#1e2535", color = "#7a8a9e"),
        yaxis = list(title = "", color = "#7a8a9e"),
        margin = list(l = 260, r = 20, t = 20, b = 60),
        hoverlabel = list(bgcolor = "#1e2535", bordercolor = "#2e3a50", font = list(color = "#e0e7f0"))
      )
  })

  # CSV downloads for every Overview visualization
  register_data_download("home_radar_chart", function() {
    req(input$home_radar_countries)
    key_indices <- c("Economic Freedom Index", "E-Government Development Index", "Global Innovation Index", "Environmental Performance Index")
    selected_countries <- input$home_radar_countries
    if (length(selected_countries) == 0) selected_countries <- "Oman"

    country_scores <- all_data %>%
      filter(Index %in% key_indices, Country %in% selected_countries, !is.na(Score)) %>%
      group_by(Country, Index) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      mutate(Series = Country) %>%
      select(Series, Country, Pillar, Index, Year, Score)

    gcc_average <- all_data %>%
      filter(Index %in% key_indices, !is.na(Score)) %>%
      group_by(Index, Country) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      group_by(Index) %>%
      summarise(Series = "GCC Average", Country = "GCC Average", Pillar = Pillar[1], Year = max(Year, na.rm = TRUE), Score = mean(Score, na.rm = TRUE), .groups = "drop") %>%
      select(Series, Country, Pillar, Index, Year, Score)

    bind_rows(country_scores, gcc_average) %>% arrange(Index, Series)
  })

  register_data_download("home_map_chart", function() {
    req(input$home_map_index)
    all_data %>%
      filter(Index == input$home_map_index, !is.na(Score)) %>%
      group_by(Country) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      mutate(ISO3 = COUNTRY_ISO3[Country], Lon = COUNTRY_LON[Country], Lat = COUNTRY_LAT[Country]) %>%
      select(Pillar, Index, Country, ISO3, Lon, Lat, Year, Score, Rank)
  })

  register_data_download("home_overview_plot", function() {
    key_indices <- c("Economic Freedom Index", "E-Government Development Index", "Global Innovation Index", "Environmental Performance Index")
    oman <- all_data %>%
      filter(Index %in% key_indices, Country == "Oman", !is.na(Score)) %>%
      mutate(Series = "Oman") %>%
      select(Series, Pillar, Index, Country, Year, Score)

    gcc <- all_data %>%
      filter(Index %in% key_indices, !is.na(Score)) %>%
      group_by(Pillar, Index, Year) %>%
      summarise(Series = "GCC Average", Country = "GCC Average", Score = mean(Score, na.rm = TRUE), .groups = "drop") %>%
      select(Series, Pillar, Index, Country, Year, Score)

    bind_rows(oman, gcc) %>% arrange(Index, Series, Year)
  })

  register_data_download("home_rank_bubble", function() {
    all_data %>%
      filter(!is.na(Rank)) %>%
      group_by(Country, Index) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      mutate(IndexShort = case_when(
        Index == "Corruption Perceptions Index" ~ "Corruption",
        Index == "Economic Complexity Index" ~ "Complexity",
        Index == "Economic Freedom Index" ~ "Economic Freedom",
        Index == "E-Government Development Index" ~ "E-Government",
        Index == "Environmental Performance Index" ~ "Environment",
        Index == "Global Innovation Index" ~ "Innovation",
        Index == "Global Talent Competitiveness Index" ~ "Talent",
        Index == "Legatum Prosperity Index (Overall)" ~ "Prosperity",
        Index == "Legatum Prosperity Index – Social Capital" ~ "Social Capital",
        Index == "Legatum Health Index" ~ "Health",
        Index == "Social Progress Index" ~ "Social Progress",
        TRUE ~ Index
      )) %>%
      select(Pillar, Index, IndexShort, Country, Year, Rank, Score)
  })

  register_data_download("home_boxplot", function() {
    all_data %>% filter(!is.na(Score)) %>% select(Pillar, Index, Country, Year, Score, Rank)
  })

  register_data_download("home_two_index_compare", function() {
    req(input$compare_x, input$compare_y)
    if (input$compare_x == input$compare_y) return(tibble(Message = "Choose two different indices for comparison."))
    latest_scores <- all_data %>%
      filter(Index %in% c(input$compare_x, input$compare_y), !is.na(Score)) %>%
      group_by(Country, Index) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      select(Country, Index, Score, Year) %>%
      pivot_wider(names_from = Index, values_from = c(Score, Year), names_sep = "__")
    xcol <- paste0("Score__", input$compare_x)
    ycol <- paste0("Score__", input$compare_y)
    latest_scores %>%
      mutate(
        X_Index = input$compare_x,
        Y_Index = input$compare_y,
        XScore = if (xcol %in% names(.)) .data[[xcol]] else NA_real_,
        YScore = if (ycol %in% names(.)) .data[[ycol]] else NA_real_
      ) %>%
      select(Country, X_Index, Y_Index, XScore, YScore, everything())
  })

  register_data_download("home_gov_innovation_freedom_bubble", function() {
    needed_indices <- c("E-Government Development Index", "Global Innovation Index", "Economic Freedom Index")
    latest_scores <- all_data %>%
      filter(Index %in% needed_indices, !is.na(Score)) %>%
      group_by(Country, Index) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      select(Country, Index, Score, Year) %>%
      pivot_wider(names_from = Index, values_from = c(Score, Year), names_sep = "__")
    latest_scores %>%
      mutate(
        EGovernment = .data[["Score__E-Government Development Index"]],
        Innovation = .data[["Score__Global Innovation Index"]],
        EconomicFreedom = .data[["Score__Economic Freedom Index"]]
      ) %>%
      select(Country, EGovernment, Innovation, EconomicFreedom, everything())
  })

  register_data_download("home_oman_gap_all", function() {
    all_data %>%
      filter(!is.na(Score)) %>%
      group_by(Index, Country) %>% arrange(desc(Year)) %>% slice(1) %>% ungroup() %>%
      group_by(Index) %>%
      summarise(
        Oman = Score[Country == "Oman"][1],
        GCC_Avg = mean(Score, na.rm = TRUE),
        Pillar = Pillar[Country == "Oman"][1],
        Gap = Oman - GCC_Avg,
        .groups = "drop"
      ) %>%
      filter(!is.na(Oman)) %>%
      arrange(Gap)
  })

}

# ============================================================
# RUN
# ============================================================

shinyApp(ui, server)
