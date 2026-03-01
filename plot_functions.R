

simulate_km_plot <- function(
    figure_title     = "Kaplan–Meier Plot (Simulated)",
    time_unit        = "Days",
    arms             = c("Placebo", "Active"),
    n_per_arm        = 100,
    followup_max     = 180,
    median_time_arm  = c(120, 160),
    censoring_rate   = 0.20,
    seed             = 1,
    x_label          = NULL,
    y_label          = "Survival probability",
    show_censors     = TRUE,
    add_risktable    = TRUE,
    
    # ---- CI options ----
    show_confint     = FALSE,
    conf_level       = 0.95,
    conf_style       = c("ribbon", "step"),
    conf_alpha       = 0.20,
    
    # ---- NEW display controls ----
    font_family      = "Courier New",
    legend_position  = "right",
    add_stats        = TRUE,
    
    debug            = FALSE
) {
  
  if (!requireNamespace("survival", quietly = TRUE)) {
    stop("Package 'survival' is required.", call. = FALSE)
  }
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required.", call. = FALSE)
  }
  if (!requireNamespace("survminer", quietly = TRUE)) {
    stop("Package 'survminer' is required.", call. = FALSE)
  }
  
  # ---- helpers ----
  recycle_to_length <- function(x, n) {
    if (length(x) == n) return(x)
    if (length(x) == 1) return(rep(x, n))
    stop("Length mismatch: expected length 1 or ", n, ", got ", length(x), ".", call. = FALSE)
  }
  
  cens_prop_exp_admin <- function(mu, lambda, F) {
    denom <- lambda + mu
    mu/denom * (1 - exp(-denom * F)) + exp(-denom * F)
  }
  
  solve_mu_for_censoring <- function(target_cens, lambda, F, arm_name = NULL) {
    target_cens <- min(max(target_cens, 0), 0.999999)
    min_cens <- exp(-lambda * F)
    
    if (target_cens < min_cens - 1e-12) {
      warning(
        sprintf("Target censoring %.3f infeasible for %s. Using mu=0.", target_cens,
                if (!is.null(arm_name)) paste0("arm '", arm_name, "'") else "this arm"),
        call. = FALSE
      )
      return(0)
    }
    
    if (abs(target_cens - min_cens) < 1e-12) return(0)
    
    f <- function(mu) cens_prop_exp_admin(mu, lambda, F) - target_cens
    lo <- 0
    hi <- max(1e-8, lambda)
    
    flo <- f(lo)
    fhi <- f(hi)
    iter <- 0
    while (flo * fhi > 0 && iter < 60) {
      hi  <- hi * 2
      fhi <- f(hi)
      iter <- iter + 1
    }
    if (flo * fhi > 0) {
      warning("Could not bracket censoring solution; using mu=0.", call. = FALSE)
      return(0)
    }
    
    stats::uniroot(f, lower = lo, upper = hi, tol = 1e-12)$root
  }
  
  fmt_p <- function(p) {
    if (is.na(p)) return("NA")
    if (p < 1e-4) return("<0.0001")
    formatC(p, format = "f", digits = 4)
  }
  
  # ---- normalize inputs ----
  n_arms <- length(arms)
  n_per_arm       <- recycle_to_length(n_per_arm, n_arms)
  median_time_arm <- recycle_to_length(median_time_arm, n_arms)
  censoring_rate  <- recycle_to_length(censoring_rate, n_arms)
  
  if (is.null(x_label)) x_label <- paste0("Time (", time_unit, ")")
  conf_style <- match.arg(conf_style)
  
  # ---- simulate ----
  set.seed(seed)
  eps_time <- followup_max / 1e6
  
  dat_list <- vector("list", n_arms)
  for (k in seq_len(n_arms)) {
    arm_k  <- arms[k]
    n_k    <- n_per_arm[k]
    med_k  <- median_time_arm[k]
    cens_k <- censoring_rate[k]
    
    lambda <- log(2) / med_k
    mu     <- solve_mu_for_censoring(cens_k, lambda, followup_max, arm_name = arm_k)
    
    T <- stats::rexp(n_k, rate = lambda)
    C <- if (mu > 0) stats::rexp(n_k, rate = mu) else rep(Inf, n_k)
    F <- rep(followup_max, n_k)
    
    time   <- pmin(T, C, F)
    status <- as.integer(T <= C & T <= F)
    time   <- pmax(time, eps_time)
    
    dat_list[[k]] <- data.frame(
      arm    = factor(rep(arm_k, n_k), levels = arms),
      time   = time,
      status = status
    )
  }
  dat <- do.call(rbind, dat_list)
  
  # ---- fit ----
  fit <- survival::survfit(survival::Surv(time, status) ~ arm, data = dat, conf.int = conf_level)
  
  breaks <- pretty(c(0, followup_max), n = 6)
  breaks <- breaks[breaks >= 0 & breaks <= followup_max]
  break_time_by <- if (length(breaks) >= 2) diff(breaks)[1] else followup_max / 5
  
  # ---- Force Courier across BOTH plot + risk table ----
  base_theme <- ggplot2::theme_minimal(base_family = font_family) +
    ggplot2::theme(
      text            = ggplot2::element_text(family = font_family),
      plot.title      = ggplot2::element_text(family = font_family),
      axis.title      = ggplot2::element_text(family = font_family),
      axis.text       = ggplot2::element_text(family = font_family),
      legend.text     = ggplot2::element_text(family = font_family),
      legend.title    = ggplot2::element_text(family = font_family),
      legend.position = legend_position
    )
  
  table_theme <- ggplot2::theme_minimal(base_family = font_family) +
    ggplot2::theme(
      text       = ggplot2::element_text(family = font_family),
      axis.text  = ggplot2::element_text(family = font_family),
      axis.title = ggplot2::element_text(family = font_family),
      plot.title = ggplot2::element_text(family = font_family),
      legend.position = "none"
    )
  
  g <- survminer::ggsurvplot(
    fit,
    data              = dat,
    title             = figure_title,
    xlab              = x_label,
    ylab              = y_label,
    
    censor            = isTRUE(show_censors),
    risk.table        = isTRUE(add_risktable),
    risk.table.title  = if (isTRUE(add_risktable)) "Number at risk" else NULL,
    risk.table.height = if (isTRUE(add_risktable)) 0.25 else NULL,
    break.time.by     = break_time_by,
    
    conf.int          = isTRUE(show_confint),
    conf.int.style    = conf_style,
    conf.int.alpha    = conf_alpha,
    
    # ggtheme           = theme_shell(...), #base_theme,
    # tables.theme      = table_theme
    
    
    ggtheme      = theme_shell(font_family = font_family, base_size = 9),
    tables.theme = theme_shell_risktable(font_family = font_family, base_size = 9)
    
  )
  
  # Belt & braces font enforcement
  g$plot  <- force_ggplot_text_family(g$plot,  font_family)
  if (isTRUE(add_risktable) && !is.null(g$table)) {
    g$table <- force_ggplot_text_family(g$table, font_family)
  }
  
  # Make margins identical so risk table columns line up with plot ticks
  m <- ggplot2::margin(t = 5.5, r = 25, b = 5.5, l = 5.5, unit = "pt")
  
  g$plot <- g$plot + ggplot2::theme(plot.margin = m)
  
  if (isTRUE(add_risktable) && !is.null(g$table)) {
    g$table <- g$table + ggplot2::theme(plot.margin = m)
  }
  
  g$table <- g$table + ggplot2::scale_y_discrete(labels = function(x) sub("^arm=", "", x))
  
  # Belt-and-braces: reapply in case survminer overwrites parts of theme
  if (!is.null(g$plot))  g$plot  <- g$plot  + base_theme
  if (isTRUE(add_risktable) && !is.null(g$table)) g$table <- g$table + table_theme
  
  # ---- stats block (top-right): log-rank p + HR + label on next row ----
  if (isTRUE(add_stats) && n_arms == 2) {
    
    sdif <- survival::survdiff(survival::Surv(time, status) ~ arm, data = dat)
    p_lr <- 1 - stats::pchisq(sdif$chisq, df = 1)
    
    cox <- survival::coxph(survival::Surv(time, status) ~ arm, data = dat)
    hr  <- unname(exp(stats::coef(cox))[1])
    ci  <- exp(stats::confint(cox))[1, ]
    
    # IMPORTANT: with factor levels arms[1], arms[2],
    # Cox coefficient corresponds to arms[2] vs arms[1]
    stat_label <- paste(
      paste0("Log-rank p=", fmt_p(p_lr)),
      sprintf("HR=%.3f (%.3f, %.3f)", hr, ci[1], ci[2]),
      sprintf("(%s vs %s)", arms[2], arms[1]),
      sep = "\n"
    )
    
    g$plot <- g$plot +
      ggplot2::annotate(
        "text",
        x = Inf, y = Inf,
        label  = stat_label,
        hjust  = 1.05, vjust = 1.15,
        family = font_family,
        size   = 3.0
      )
  }
  
  invisible(list(plot = g, fit = fit, data = dat))
}



force_ggplot_text_family <- function(p, family) {
  if (is.null(p) || !inherits(p, "gg")) return(p)
  
  # Theme-level (works for axes/title etc.)
  p <- p + ggplot2::theme(text = ggplot2::element_text(family = family))
  
  # Layer-level (needed for geom_text in survminer tables)
  if (!is.null(p$layers) && length(p$layers) > 0) {
    for (i in seq_along(p$layers)) {
      lyr <- p$layers[[i]]
      if (inherits(lyr$geom, "GeomText") || inherits(lyr$geom, "GeomLabel")) {
        if (is.null(lyr$aes_params)) lyr$aes_params <- list()
        lyr$aes_params$family <- family
        p$layers[[i]] <- lyr
      }
    }
  }
  p
}
















