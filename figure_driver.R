# ==========================================================
# figure_driver.R -- Figure spec reader + dispatcher (functions only)
# ==========================================================
# PURPOSE:
#  - Read a per-figure Excel sheet (e.g., FIG_1) stored as key/value pairs
#  - Parse/normalise types (nums/bools/vectors)
#  - Dispatch to the appropriate figure builder by type_of_plot
#
# IMPORTANT:
#  - This file should contain FUNCTIONS ONLY.
#  - It relies on plot builders (e.g., simulate_km_plot()) defined in plot_functions.R
# ==========================================================

`%||%` <- function(a, b) if (!is.null(a) && !is.na(a) && a != "") a else b

# ---------- FIGURE SPEC (key/value) reader ----------
read_fig_kv_spec <- function(excel_path, sheet) {
  x <- readxl::read_xlsx(excel_path, sheet = sheet)
  names(x) <- tolower(trimws(names(x)))
  
  if (!all(c("key", "value") %in% names(x))) {
    stop("Figure sheet '", sheet, "' must have columns: key, value", call. = FALSE)
  }
  
  key <- tolower(trimws(as.character(x$key)))
  val <- trimws(as.character(x$value))
  
  ok  <- !is.na(key) & key != ""
  key <- key[ok]
  val <- val[ok]
  
  spec <- as.list(val)
  names(spec) <- key
  
  spec
}

# ---------- helpers to parse values ----------
parse_bool <- function(x) {
  if (is.null(x) || is.na(x) || x == "") return(NA)
  xx <- tolower(trimws(as.character(x)))
  if (xx %in% c("true", "t", "1", "yes", "y"))  return(TRUE)
  if (xx %in% c("false", "f", "0", "no", "n"))  return(FALSE)
  NA
}

parse_num <- function(x) {
  if (is.null(x) || is.na(x) || x == "") return(NA_real_)
  suppressWarnings(as.numeric(trimws(as.character(x))))
}

strip_quotes <- function(x) {
  if (is.null(x) || is.na(x)) return(x)
  xx <- trimws(as.character(x))
  # remove one pair of surrounding single or double quotes
  xx <- gsub('^"(.*)"$', "\\1", xx)
  xx <- gsub("^'(.*)'$", "\\1", xx)
  xx
}

parse_pipe_chr <- function(x) {
  if (is.null(x) || is.na(x) || x == "") return(character(0))
  trimws(strsplit(as.character(x), "\\|", fixed = FALSE)[[1]])
}

parse_pipe_num <- function(x) {
  out <- suppressWarnings(as.numeric(parse_pipe_chr(x)))
  out
}

# ---------- normalise a FIG spec into typed fields ----------
normalise_fig_spec <- function(spec) {
  
  # tolerate common key variants/typos
  # (you mentioned show_conflint in Excel)
  if (!is.null(spec$show_conflint) && is.null(spec$show_confint)) {
    spec$show_confint <- spec$show_conflint
  }
  
  # Build typed output list
  out <- list(
    type_of_plot    = toupper(trimws(spec$type_of_plot %||% "")),
    figure_title    = spec$figure_title %||% "",
    
    # General / common
    time_unit       = spec$time_unit %||% "Weeks",
    arms            = parse_pipe_chr(spec$arms),
    n_per_arm       = parse_pipe_num(spec$n_per_arm),
    
    # KM-specific (but harmless defaults for others)
    followup_max    = parse_num(spec$followup_max),
    median_time_arm = parse_pipe_num(spec$median_time_arm),
    censoring_rate  = parse_num(spec$censoring_rate),
    seed            = as.integer(parse_num(spec$seed)),
    
    show_censors    = parse_bool(spec$show_censors),
    add_risktable   = parse_bool(spec$add_risktable),
    
    # CI options
    show_confint    = parse_bool(spec$show_confint),
    conf_level      = parse_num(spec$conf_level),
    conf_style      = strip_quotes(spec$conf_style),
    conf_alpha      = parse_num(spec$conf_alpha),
    
    # Optional figure size (inches)
    fig_width       = parse_num(spec$fig_width),
    fig_height      = parse_num(spec$fig_height),
    
    # Optional footnote keys, e.g. "fn1|fn2|pn1"
    footnotes       = parse_pipe_chr(spec$footnotes),
    
    
    sc_visit_numbers = parse_pipe_num(spec$sc_visit_numbers),
    sc_labelx        = spec$sc_labelx %||% "Baseline value",
    sc_labely        = spec$sc_labely %||% "Value",
    sc_corr          = parse_num(spec$sc_corr),
    sc_y_mode        = tolower(strip_quotes(spec$sc_y_mode %||% "aval")),
    sc_sd_baseline   = parse_num(spec$sc_sd_baseline),
    sc_sd_error      = parse_num(spec$sc_sd_error),
    sc_slope_per_unit= parse_num(spec$sc_slope_per_unit),
    
    sc_se_style = tolower(strip_quotes(spec$sc_se_style %||% "bar")),
    
    
    sc_mean       = parse_pipe_num(spec$sc_mean),
    sc_trt_effect = parse_pipe_num(spec$sc_trt_effect),
    
    analysis_set = spec$analysis_set %||% spec[["analysis set"]] %||% ""
    
    
    
    
  )
  
  # ---- set safe defaults if missing ----
  if (is.na(out$seed)) out$seed <- 1L
  if (is.na(out$censoring_rate)) out$censoring_rate <- 0.20
  if (is.na(out$conf_level)) out$conf_level <- 0.95
  if (is.na(out$conf_alpha)) out$conf_alpha <- 0.20
  
  
  if (is.na(out$sc_corr)) out$sc_corr <- 0.6
  if (is.na(out$sc_sd_baseline)) out$sc_sd_baseline <- 10
  if (is.na(out$sc_sd_error)) out$sc_sd_error <- 6
  if (is.na(out$sc_slope_per_unit)) out$sc_slope_per_unit <- 0
  if (is.null(out$sc_y_mode) || out$sc_y_mode == "") out$sc_y_mode <- "aval"
  
  if (is.null(out$sc_se_style) || !(out$sc_se_style %in% c("bar","ribbon","none"))) out$sc_se_style <- "bar"
  
  if (!length(out$sc_mean)) out$sc_mean <- 0
  if (!length(out$sc_trt_effect)) out$sc_trt_effect <- 0
  
  
  # conf_style: allow empty -> default "ribbon"
  if (is.null(out$conf_style) || is.na(out$conf_style) || out$conf_style == "") out$conf_style <- "ribbon"
  
  # ---- basic validation for KM ----
  if (out$type_of_plot == "KM") {
    if (length(out$arms) < 1) {
      stop("FIG spec missing 'arms' for KM", call. = FALSE)
    }
    if (length(out$n_per_arm) != length(out$arms)) {
      stop("'n_per_arm' must match length of 'arms'", call. = FALSE)
    }
    if (is.na(out$followup_max)) {
      stop("FIG spec missing 'followup_max' for KM", call. = FALSE)
    }
    if (length(out$median_time_arm) > 0 && length(out$median_time_arm) != length(out$arms)) {
      stop("'median_time_arm' must match length of 'arms' (or be blank)", call. = FALSE)
    }
    # If median_time_arm is blank, allow builder to recycle/default internally
  }
  
  out
}

# ---------- build a figure from spec (dispatcher) ----------
build_figure <- function(fig_spec) {
  
  tp <- toupper(trimws(fig_spec$type_of_plot %||% ""))
  
  if (tp == "KM") {
    
    if (!exists("simulate_km_plot", mode = "function")) {
      stop("simulate_km_plot() not found. Ensure plot_functions.R is sourced before figure_driver.R.", call. = FALSE)
    }
    
    # Call your KM builder; it returns list(plot=ggsurvplot, fit=..., data=...)
    return(
      simulate_km_plot(
        figure_title    = fig_spec$figure_title %||% "Kaplan–Meier Plot (Simulated)",
        time_unit       = fig_spec$time_unit %||% "Weeks",
        arms            = fig_spec$arms,
        n_per_arm       = fig_spec$n_per_arm,
        followup_max    = fig_spec$followup_max,
        median_time_arm = if (length(fig_spec$median_time_arm)) fig_spec$median_time_arm else NULL,
        censoring_rate  = fig_spec$censoring_rate,
        seed            = fig_spec$seed,
        
        show_censors    = isTRUE(fig_spec$show_censors),
        add_risktable   = isTRUE(fig_spec$add_risktable),
        
        show_confint    = isTRUE(fig_spec$show_confint),
        conf_level      = fig_spec$conf_level,
        conf_style      = fig_spec$conf_style,
        conf_alpha      = fig_spec$conf_alpha,
        
        # NEW (hard override for all KM figures)
        font_family     = "Courier New",
        legend_position = "right",
        add_stats       = TRUE
      )
    )
  }
  
  
  if (tp == "SCATTER") {
    
    if (!exists("simulate_scatter_plot", mode = "function")) {
      stop("simulate_scatter_plot() not found. Ensure plot_functions_scatter.R is sourced.", call. = FALSE)
    }
    
    return(
      simulate_scatter_plot(
        figure_title     = fig_spec$figure_title %||% "Scatterplot (Simulated)",
        time_unit        = fig_spec$time_unit %||% "Weeks",
        arms             = fig_spec$arms,
        n_per_arm        = fig_spec$n_per_arm,
        sc_visit_numbers = fig_spec$sc_visit_numbers,
        sc_labelx        = fig_spec$sc_labelx,
        sc_labely        = fig_spec$sc_labely,
        sc_corr          = fig_spec$sc_corr,
        sc_y_mode        = if (fig_spec$sc_y_mode %in% c("aval","chg")) fig_spec$sc_y_mode else "aval",
        sc_sd_baseline   = fig_spec$sc_sd_baseline,
        sc_sd_error      = fig_spec$sc_sd_error,
        sc_slope_per_unit= fig_spec$sc_slope_per_unit,
        seed             = fig_spec$seed,
        font_family      = "Courier New"
      )
    )
  }
  
  
  if (tp == "SCATTER2") {
    
    if (!exists("simulate_scatter2_plot", mode = "function")) {
      stop("simulate_scatter2_plot() not found. Ensure plot_functions_scatter2.R is sourced.", call. = FALSE)
    }
    
    return(
      simulate_scatter2_plot(
        figure_title      = fig_spec$figure_title %||% "Value Over Time (Simulated)",
        time_unit         = fig_spec$time_unit %||% "Weeks",
        arms              = fig_spec$arms,
        n_per_arm         = fig_spec$n_per_arm,
        sc_visit_numbers  = fig_spec$sc_visit_numbers,
        sc_labelx         = fig_spec$sc_labelx,
        sc_labely         = fig_spec$sc_labely,
        sc_corr           = fig_spec$sc_corr,
        sc_sd_baseline    = fig_spec$sc_sd_baseline,
        sc_sd_error       = fig_spec$sc_sd_error,
        sc_slope_per_unit = fig_spec$sc_slope_per_unit,
        sc_se_style       = fig_spec$sc_se_style,
        seed              = fig_spec$seed,
        sc_mean       = fig_spec$sc_mean,
        sc_trt_effect = fig_spec$sc_trt_effect,
        font_family       = "Courier New"
      )
    )
  }
  
  
  # Future-proof placeholders (add as you implement them)
  # if (tp == "SPAGHETTI") return(simulate_spaghetti_plot(...))
  # if (tp == "SCATTER")   return(simulate_scatter_plot(...))
  # if (tp == "BAR")       return(simulate_bar_plot(...))
  
  stop("Unsupported type_of_plot: '", tp, "'. Add a case in build_figure() for this plot type.", call. = FALSE)
}