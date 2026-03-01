

# ==========================================================
# plot_functions_scatter.R -- Scatterplot shell (functions only)
# Baseline vs Value at each visit (default)
# ==========================================================

simulate_scatter_plot <- function(
    figure_title     = "Scatterplot (Simulated)",
    time_unit        = "Weeks",
    arms             = c("Placebo", "Active"),
    n_per_arm        = 100,
    
    sc_visit_numbers = c(4, 8, 12),   # numeric vector; visits after baseline
    sc_labelx        = "Baseline value",
    sc_labely        = "Value",
    sc_corr          = 0.6,           # target corr(BASE, AVAL at visit)
    sc_y_mode        = c("aval", "chg"),  # default is aval (baseline vs value)
    sc_sd_baseline   = 10,
    sc_sd_error      = 6,
    sc_slope_per_unit= 0.0,           # optional drift per visit unit
    seed             = 1,
    font_family      = "Courier New",
    debug            = FALSE
) {
  
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required. Install with install.packages('ggplot2').", call. = FALSE)
  }
  
  sc_y_mode <- match.arg(sc_y_mode)
  
  # ---- helpers ----
  recycle_to_length <- function(x, n) {
    if (length(x) == n) return(x)
    if (length(x) == 1) return(rep(x, n))
    stop("Length mismatch: expected length 1 or ", n, ", got ", length(x), ".", call. = FALSE)
  }
  
  # ---- validate ----
  if (!is.numeric(sc_corr) || length(sc_corr) != 1 || is.na(sc_corr) || sc_corr <= -1 || sc_corr >= 1) {
    stop("sc_corr must be a single number in (-1, 1).", call. = FALSE)
  }
  if (!is.numeric(sc_visit_numbers) || any(is.na(sc_visit_numbers))) {
    stop("sc_visit_numbers must be numeric (e.g., 4|8|12).", call. = FALSE)
  }
  sc_visit_numbers <- sort(unique(as.numeric(sc_visit_numbers)))
  if (length(sc_visit_numbers) < 1) stop("sc_visit_numbers must contain at least one visit.", call. = FALSE)
  
  n_arms <- length(arms)
  n_per_arm <- recycle_to_length(n_per_arm, n_arms)
  
  set.seed(seed)
  
  # ---- simulate subject-level longitudinal data ----
  # Baseline ~ N(0, sd=sc_sd_baseline)
  # For visit v:
  #   AVAL_v correlated with BASE via a Gaussian construction:
  #     aval_z = corr * base_z + sqrt(1-corr^2) * eps
  #     AVAL   = aval_z * sd_baseline + mean(base) + drift(v) + measurement_error
  dat_list <- vector("list", n_arms)
  subj_counter <- 0
  
  for (k in seq_len(n_arms)) {
    arm_k <- arms[k]
    n_k   <- n_per_arm[k]
    
    usubjid <- sprintf("S%04d", seq(from = subj_counter + 1, length.out = n_k))
    subj_counter <- subj_counter + n_k
    
    base <- rnorm(n_k, mean = 0, sd = sc_sd_baseline)
    base_z <- (base - mean(base)) / stats::sd(base)
    
    visits <- c(0, sc_visit_numbers)
    
    long_k <- lapply(visits, function(v) {
      if (v == 0) {
        aval <- base
      } else {
        eps <- rnorm(n_k, mean = 0, sd = 1)
        
        aval_z <- sc_corr * base_z + sqrt(1 - sc_corr^2) * eps
        aval   <- aval_z * sc_sd_baseline + mean(base) + sc_slope_per_unit * v
        aval   <- aval + rnorm(n_k, mean = 0, sd = sc_sd_error)
      }
      
      data.frame(
        USUBJID    = usubjid,
        ARM        = factor(rep(arm_k, n_k), levels = arms),
        VISITN     = rep(v, n_k),
        TIME_UNIT  = rep(time_unit, n_k),
        BASE       = base,
        AVAL       = aval,
        stringsAsFactors = FALSE
      )
    })
    
    dat_list[[k]] <- do.call(rbind, long_k)
  }
  
  dat <- do.call(rbind, dat_list)
  dat$CHG <- dat$AVAL - dat$BASE
  
  # post-baseline only for scatter
  plot_dat <- dat[dat$VISITN %in% sc_visit_numbers, , drop = FALSE]
  
  # choose y
  yvar <- if (sc_y_mode == "chg") "CHG" else "AVAL"
  ylab <- if (nzchar(sc_labely)) sc_labely else if (sc_y_mode == "chg") "Change from baseline" else "Value"
  
  if (isTRUE(debug)) {
    byv <- split(plot_dat, plot_dat$VISITN)
    for (v in names(byv)) {
      cc <- stats::cor(byv[[v]]$BASE, byv[[v]]$AVAL)
      message(sprintf("[DEBUG] VISITN=%s achieved cor(BASE,AVAL)=%.3f", v, cc))
    }
  }
  
  # ---- plot ----
  p <- ggplot2::ggplot(plot_dat, ggplot2::aes(x = BASE, y = .data[[yvar]], colour = ARM)) +
    ggplot2::geom_point(alpha = 0.7, size = 1.4) +
    ggplot2::geom_smooth(method = "lm", se = FALSE, linewidth = 0.5) +
    ggplot2::facet_wrap(
      ~ VISITN,
      scales = "free_y",
      labeller = ggplot2::labeller(VISITN = function(v) paste0(time_unit, " ", v))
    ) +
    ggplot2::labs(
      title  = figure_title,
      x      = sc_labelx,
      y      = ylab,
      colour = "Arm"
    ) +
    ggplot2::theme_minimal(base_family = font_family) +
    ggplot2::theme(
      plot.title = ggplot2::element_text(hjust = 0.5),
      legend.position = "bottom"
    )
  
  invisible(list(
    plot = p,
    data = dat,
    meta = list(
      arms = arms,
      n_per_arm = n_per_arm,
      sc_visit_numbers = sc_visit_numbers,
      sc_corr = sc_corr,
      sc_y_mode = sc_y_mode,
      sc_sd_baseline = sc_sd_baseline,
      sc_sd_error = sc_sd_error,
      sc_slope_per_unit = sc_slope_per_unit,
      seed = seed
    )
  ))
}