# plot_theme.R

theme_shell <- function(font_family = "Courier New", base_size = 9) {
  ggplot2::theme_minimal(base_family = font_family, base_size = base_size) +
    ggplot2::theme(
      plot.title      = ggplot2::element_text(hjust = 0.5),
      legend.position = "right",
      panel.grid.minor = ggplot2::element_blank()
    )
}

# Risk table theme companion for survminer::ggsurvplot()
theme_shell_risktable <- function(font_family = "Courier New", base_size = 9) {
  ggplot2::theme_minimal(base_family = font_family, base_size = base_size) +
    ggplot2::theme(
      text            = ggplot2::element_text(family = font_family),
      legend.position = "none",
      panel.grid.minor = ggplot2::element_blank()
    )
}

