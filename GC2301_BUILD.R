# ==========================================================
# build.R  -- Multilevel Table Shell Builder + Figure Shell Builder
# ==========================================================
rm(list = ls())

# ---- Libraries ----
suppressPackageStartupMessages({
  library(officer)
  library(flextable)
  library(readxl)
  library(dplyr)
  library(tibble)
  library(ragg)
  library(gridExtra)
  library(grid)
  library(ggplot2)
})  

# ---- Source external functions (MUST be functions-only files) ----
source("plot_theme.R")              # longitudinal theme helpers
source("plot_functions.R")          # e.g., simulate_km_plot(), simulate_scatter_plot(), simulate_scatter2_plot()
source("figure_driver.R")           # read_fig_kv_spec(), normalise_fig_spec(), build_figure(), maybe %||%
source("plot_functions_scatter.R")  # SCATTER
source("plot_functions_scatter2.R") # longitudinal

# Safety fallback in case figure_driver.R does not define it
if (!exists("%||%", mode = "function")) {
  `%||%` <- function(a, b) if (!is.null(a) && !is.na(a) && a != "") a else b
}

# ==========================================================
# CONSTANTS (centralised here)
# ==========================================================
CONST <- list(
  PATHS = list(
    excel_path  = "GC2301_BUILDER.xlsx",
    out_dir     = file.path(getwd(), "outputs"),
    out_prefix  = "GC2301_SHELLS_"
  ),
  
  SHEETS = list(
    dict_sheet = "DICTIONARY"
  ),
  
  # REGEX = list(
  #   table_sheets = "^TABLE\\s*\\d+(?:\\.\\d+)*$"
  # ),
  
  
  # REGEX = list(
  #   table_sheets   = "^TABLE\\s*\\d+(?:\\.\\d+)*$",
  #   listing_sheets = "^LISTING\\s*\\d+(?:\\.\\d+)*$"
  # ),
  
  REGEX = list(
    # Accept: "T 1", "T1", "TABLE 1", "TABLE1", "T 14.1.1", "TABLE 14.1.1"
    table_sheets   = "^T(?:ABLE)?\\s*\\d+(?:\\.\\d+)*$",
    
    # Accept: "L 1", "L1", "LISTING 1", "LISTING1", "L 12.3", "LISTING 12.3"
    listing_sheets = "^L(?:ISTING)?\\s*\\d+(?:\\.\\d+)*$",
    
    # Accept: "F_1", "FIG_1", "F 1", "FIG 1", "F1", "FIG1"
    # (keeps underscore support for your current FIG_# convention)
    figure_sheets  = "^F(?:IG(?:URE)?)?[_\\s]*\\d+$"
  ),
  
  PAGE = list(
    page_size = list(orient = "landscape", width = 11.69, height = 8.27),
    margins   = list(top = 0.75, bottom = 0.75, left = 0.75, right = 0.75, header = 0.40, footer = 0.40)
  ),
  
  DISPLAY = list(
    font_family   = "Courier New",
    font_size     = 8.0,
    line_spacing  = 0.86,
    col_header_bg = "#ECEEEF"
  ),
  
  BORDERS = list(
    thick = list(color = "#8E959E", width = 1.0)
  ),
  
  HEADER = list(
    left_line1 = "GigaGen, Inc. A Grifols Company",
    left_line2 = "GC2301"
  ),
  
  FOOTER = list(
    program_placeholder  = "PROGRAM: path/xxx.SAS",
    datetime_placeholder = "(XXXxxx2026 xx:xx)"
  ),
  
  DEFAULTS = list(
    placeholder_npct = "xx (xx.x%)",
    default_col_keys = c("Description", "Statistic", "Xembify", "Gamunex-C", "Total"),
    default_header_labels = c(
      Description = "Description",
      Statistic   = "Statistic",
      Xembify     = "Xembify",
      `Gamunex-C` = "Gamunex-C",
      Total       = "Total"
    )
  ),
  
  TYPES = list(
    section   = "section",
    blankstat = "blankstat",
    pagebreak = "pagebreak",
    hline     = "hline",
    toprule   = "toprule",
    footnote  = "footnote",
    empty     = "empty"
  ),
  
  WORD = list(
    heading1_style_id = "Heading1"
  )
)

# ==========================================================
# USER SETTINGS (derived from CONST)
# ==========================================================
excel_path <- CONST$PATHS$excel_path
out_dir    <- CONST$PATHS$out_dir
out_prefix <- CONST$PATHS$out_prefix
DICT_SHEET <- CONST$SHEETS$dict_sheet
PAGE       <- CONST$PAGE

CFG <- list(
  DISPLAY = list(
    font_family   = CONST$DISPLAY$font_family,
    font_size     = CONST$DISPLAY$font_size,
    line_spacing  = CONST$DISPLAY$line_spacing,
    col_header_bg = CONST$DISPLAY$col_header_bg
  ),
  BORDERS = list(
    thick = fp_border(color = CONST$BORDERS$thick$color, width = CONST$BORDERS$thick$width)
  )
)

# IMPORTANT: We control page breaks in R → assume Heading 1 does NOT have "Page break before"
CFG$WORD$heading1_has_pagebreak_before <- FALSE

HDR <- list(
  left_line1  = CONST$HEADER$left_line1,
  left_line2  = CONST$HEADER$left_line2,
  font_family = CFG$DISPLAY$font_family,
  font_size   = CFG$DISPLAY$font_size
)

FTR <- list(
  program_text  = CONST$FOOTER$program_placeholder,
  datetime_text = CONST$FOOTER$datetime_placeholder,
  font_family   = CFG$DISPLAY$font_family,
  font_size     = CFG$DISPLAY$font_size
)

# ==========================================================
# Utility helpers
# ==========================================================


add_section_cover_page <- function(doc, heading_text, style = "heading 1") {
  # Ensure heading starts on a fresh page (but don’t double-break)
  doc <- add_page_break_safe(doc)
  mark_content_added()
  
  # Add heading
  doc <- officer::body_add_par(doc, heading_text, style = style)
  mark_content_added()
  
  # Force next content (first table/figure/listing) onto a new page
  doc <- force_page_break(doc)
  mark_content_added()
  
  doc
}



ensure_toc_styles <- function(doc,
                              levels = 3,
                              font_family = CFG$DISPLAY$font_family,
                              font_size   = CFG$DISPLAY$font_size,
                              base_on     = "Normal") {
  
  fp_t <- officer::fp_text(
    font.family = font_family,
    font.size   = font_size,
    bold        = FALSE,
    color       = "black"
  )
  
  # Keep TOC tight; adjust if you want more/less density
  fp_p <- officer::fp_par(padding = 0, line_spacing = CFG$DISPLAY$line_spacing)
  
  for (lvl in seq_len(levels)) {
    
    # Word style *name* must be "TOC 1", "TOC 2", ...
    style_name <- paste0("TOC ", lvl)
    
    # style_id must be ASCII and no spaces; pick a safe deterministic ID
    style_id <- paste0("toc", lvl)
    
    doc <- officer::docx_set_paragraph_style(
      x          = doc,
      style_id   = style_id,
      style_name = style_name,
      base_on    = base_on,
      fp_p       = fp_p,
      fp_t       = fp_t
    )
  }
  
  doc
}


bold_table_id_line <- function(ft, line1_text, ncols) {
  if (is.null(ft$header$dataset) || !nrow(ft$header$dataset)) return(ft)
  if (!is.character(line1_text) || !nzchar(trimws(line1_text))) return(ft)
  
  hd <- ft$header$dataset
  
  # Find any header row where ANY cell equals the target "Table X..." line
  # (after merges, it might live in first col or elsewhere depending on flextable internals)
  hit_rows <- which(apply(hd, 1, function(r) any(trimws(as.character(r)) == trimws(line1_text))))
  
  if (!length(hit_rows)) return(ft)
  
  # Bold only that row (across all columns)
  ft <- flextable::bold(ft, part = "header", i = hit_rows[1], j = seq_len(ncols), bold = TRUE)
  ft
}

STATE <- new.env(parent = emptyenv())
STATE$last_was_page_break <- FALSE

force_page_break <- function(doc) {
  STATE$last_was_page_break <- FALSE
  doc <- officer::body_add_break(doc)
  STATE$last_was_page_break <- TRUE
  doc
}

add_page_break_safe <- function(doc) {
  if (isTRUE(STATE$last_was_page_break)) return(doc)
  STATE$last_was_page_break <- TRUE
  officer::body_add_break(doc)
}

mark_content_added <- function() {
  STATE$last_was_page_break <- FALSE
}

normalize_empty <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  trimws(x)
}

norm_key <- function(x) {
  x <- normalize_empty(x)
  x <- tolower(x)
  x <- gsub("\\s+", " ", x)
  x
}

resolve_fig_footnotes <- function(keys, placeholder_dict = NULL) {
  keys <- trimws(as.character(keys))
  keys <- keys[!is.na(keys) & keys != ""]
  if (!length(keys)) return(character(0))
  
  if (is.null(placeholder_dict) || !length(placeholder_dict)) return(keys)
  
  kk <- tolower(keys)
  mapped <- placeholder_dict[kk]
  mapped[is.na(mapped) | mapped == ""] <- keys[is.na(mapped) | mapped == ""]
  trimws(unname(mapped))
}

# split_table_or_figure_title <- function(title) {
#   tt <- trimws(as.character(title))
#   if (!nzchar(tt)) return(list(line1 = "", line2 = ""))
#   
#   tt <- sub("^.*?\\b(Table|Figure)\\b", "\\1", tt, ignore.case = TRUE)
#   tt <- trimws(tt)
#   
#   m <- regexec(
#     "^(Table|Figure)\\s+([A-Za-z0-9]+(?:[./][A-Za-z0-9]+)*)\\s*([\\.:\\-])\\s*(.*)$",
#     tt, ignore.case = TRUE
#   )
#   mm <- regmatches(tt, m)[[1]]
#   
#   if (length(mm) == 5) {
#     kind <- mm[2]; id <- mm[3]; sep <- mm[4]; rest <- trimws(mm[5])
#     return(list(line1 = paste0(kind, " ", id, sep), line2 = rest))
#   }
#   
#   m2 <- regexec("^(Table|Figure)\\s+([A-Za-z0-9]+(?:[./][A-Za-z0-9]+)*)(?:\\s+(.+))?$",
#                 tt, ignore.case = TRUE)
#   mm2 <- regmatches(tt, m2)[[1]]
#   if (length(mm2) >= 3) {
#     kind <- mm2[2]; id <- mm2[3]
#     rest <- if (length(mm2) >= 4) trimws(mm2[4]) else ""
#     return(list(line1 = paste(kind, id), line2 = rest))
#   }
#   
#   list(line1 = tt, line2 = "")
# }

split_table_or_figure_title <- function(title) {
  tt <- trimws(as.character(title))
  if (!nzchar(tt)) return(list(line1 = "", line2 = ""))
  
  # Keep only from the first occurrence of Table/Figure/Listing
  tt <- sub("^.*?\\b(Table|Figure|Listing)\\b", "\\1", tt, ignore.case = TRUE)
  tt <- trimws(tt)
  
  m <- regexec(
    "^(Table|Figure|Listing)\\s+([A-Za-z0-9]+(?:[./][A-Za-z0-9]+)*)\\s*([\\.:\\-])\\s*(.*)$",
    tt, ignore.case = TRUE
  )
  mm <- regmatches(tt, m)[[1]]
  
  if (length(mm) == 5) {
    kind <- mm[2]; id <- mm[3]; sep <- mm[4]; rest <- trimws(mm[5])
    return(list(line1 = paste0(kind, " ", id, sep), line2 = rest))
  }
  
  m2 <- regexec(
    "^(Table|Figure|Listing)\\s+([A-Za-z0-9]+(?:[./][A-Za-z0-9]+)*)(?:\\s+(.+))?$",
    tt, ignore.case = TRUE
  )
  mm2 <- regmatches(tt, m2)[[1]]
  if (length(mm2) >= 3) {
    kind <- mm2[2]; id <- mm2[3]
    rest <- if (length(mm2) >= 4) trimws(mm2[4]) else ""
    return(list(line1 = paste(kind, id), line2 = rest))
  }
  
  list(line1 = tt, line2 = "")
}



usable_width_in <- function() {
  PAGE$page_size$width - PAGE$margins$left - PAGE$margins$right
}

usable_height_in <- function() {
  PAGE$page_size$height - PAGE$margins$top - PAGE$margins$bottom
}

.make_tabspec_right <- function() {
  officer::fp_tabs(officer::fp_tab(pos = usable_width_in(), style = "right"))
}

make_page_header_block <- function(left_line1, left_line2,
                                   font_family = CFG$DISPLAY$font_family,
                                   font_size   = CFG$DISPLAY$font_size) {
  
  tabspec <- .make_tabspec_right()
  fp_txt  <- officer::fp_text(font.family = font_family, font.size = font_size)
  
  p1 <- officer::fpar(
    officer::ftext(left_line1, prop = fp_txt),
    officer::run_tab(),
    officer::run_word_field("PAGE"),
    officer::ftext(" of ", prop = fp_txt),
    officer::run_word_field("NUMPAGES"),
    fp_p = officer::fp_par(tabs = tabspec, padding = 0, line_spacing = 1),
    fp_t = fp_txt
  )
  
  p2 <- officer::fpar(
    officer::ftext(left_line2, prop = fp_txt),
    fp_p = officer::fp_par(padding = 0, line_spacing = 1),
    fp_t = fp_txt
  )
  
  officer::block_list(p1, p2)
}

make_page_footer_block <- function(program_left,
                                   datetime_right,
                                   font_family = CFG$DISPLAY$font_family,
                                   font_size   = CFG$DISPLAY$font_size) {
  
  tabspec <- .make_tabspec_right()
  fp_txt  <- officer::fp_text(font.family = font_family, font.size = font_size)
  
  p <- officer::fpar(
    officer::ftext(program_left, prop = fp_txt),
    officer::run_tab(),
    officer::ftext(datetime_right, prop = fp_txt),
    fp_p = officer::fp_par(tabs = tabspec, padding = 0, line_spacing = 1),
    fp_t = fp_txt
  )
  
  officer::block_list(p)
}

hdr_block <- make_page_header_block(HDR$left_line1, HDR$left_line2, HDR$font_family, HDR$font_size)

ftr_block <- make_page_footer_block(
  program_left   = FTR$program_text,
  datetime_right = FTR$datetime_text,
  font_family    = FTR$font_family,
  font_size      = FTR$font_size
)

blank_ftr_block <- officer::block_list(
  officer::fpar(officer::ftext(""))
)

sec_landscape <- prop_section(
  page_size    = page_size(orient = PAGE$page_size$orient,
                           width  = PAGE$page_size$width,
                           height = PAGE$page_size$height),
  page_margins = page_mar(top    = PAGE$margins$top,
                          bottom = PAGE$margins$bottom,
                          left   = PAGE$margins$left,
                          right   = PAGE$margins$right,
                          header = PAGE$margins$header,
                          footer = PAGE$margins$footer),
  type = "continuous",
  
  header_default = hdr_block,
  header_first   = hdr_block,
  header_even    = hdr_block,
  
  footer_default = ftr_block,
  footer_first   = ftr_block,
  footer_even    = ftr_block
)

sec_landscape_toc <- prop_section(
  page_size    = page_size(orient = PAGE$page_size$orient,
                           width  = PAGE$page_size$width,
                           height = PAGE$page_size$height),
  page_margins = page_mar(top    = PAGE$margins$top,
                          bottom = PAGE$margins$bottom,
                          left   = PAGE$margins$left,
                          right   = PAGE$margins$right,
                          header = PAGE$margins$header,
                          footer = PAGE$margins$footer),
  type = "continuous",
  
  header_default = hdr_block,
  header_first   = hdr_block,
  header_even    = hdr_block,
  
  footer_default = blank_ftr_block,
  footer_first   = blank_ftr_block,
  footer_even    = blank_ftr_block
)

make_table_section_props <- function(title, analysis_set,
                                     base_header_block,
                                     base_footer_block) {
  
  prop_section(
    page_size    = page_size(orient = PAGE$page_size$orient,
                             width  = PAGE$page_size$width,
                             height = PAGE$page_size$height),
    page_margins = page_mar(top    = PAGE$margins$top,
                            bottom = PAGE$margins$bottom,
                            left   = PAGE$margins$left,
                            right   = PAGE$margins$right,
                            header = PAGE$margins$header,
                            footer = PAGE$margins$footer),
    type = "continuous",
    
    header_first   = base_header_block,
    header_default = base_header_block,
    header_even    = base_header_block,
    
    footer_first   = base_footer_block,
    footer_default = base_footer_block,
    footer_even    = base_footer_block
  )
}

if (is.null(CFG$FIG)) {
  CFG$FIG <- list(
    width_frac       = 0.88,
    max_width        = 9.2,
    default_aspect   = 0.58,
    max_height_frac  = 0.58,
    dpi              = 300
  )
}

fig_target_width_in <- function() {
  w <- usable_width_in() * CFG$FIG$width_frac
  if (is.finite(CFG$FIG$max_width)) w <- min(w, CFG$FIG$max_width)
  w
}

fig_target_height_in <- function(width_in = fig_target_width_in()) {
  width_in * CFG$FIG$default_aspect
}

fig_max_height_in <- function() {
  usable_height_in() * CFG$FIG$max_height_frac
}

cap_fig_wh_to_max_height <- function(w, h) {
  h_max <- fig_max_height_in()
  if (is.finite(h_max) && h > h_max) {
    scale <- h_max / h
    return(list(w = w * scale, h = h_max))
  }
  list(w = w, h = h)
}

safe_write_docx <- function(doc, out_file) {
  tmp <- tempfile(pattern = "docx_tmp_", fileext = ".docx")
  print(doc, target = tmp)
  Sys.sleep(0.3)
  
  if (!file.exists(tmp)) stop("DOCX write failed: temp file not created: ", tmp, call. = FALSE)
  
  dir.create(dirname(out_file), recursive = TRUE, showWarnings = FALSE)
  ok <- file.copy(tmp, out_file, overwrite = TRUE)
  if (!ok || !file.exists(out_file)) stop("DOCX write failed: temp created but copy failed to: ", out_file, call. = FALSE)
  
  unlink(tmp)
  invisible(out_file)
}

strip_plot_titles <- function(p) {
  if (inherits(p, "gg")) {
    return(
      p +
        ggplot2::labs(title = NULL, subtitle = NULL, caption = NULL) +
        ggplot2::theme(
          plot.title    = ggplot2::element_blank(),
          plot.subtitle = ggplot2::element_blank(),
          plot.caption  = ggplot2::element_blank()
        )
    )
  }
  
  if (inherits(p, "ggsurvplot")) {
    if (!is.null(p$plot) && inherits(p$plot, "gg")) {
      p$plot <- strip_plot_titles(p$plot)
    }
    if (!is.null(p$table) && inherits(p$table, "gg")) {
      p$table <- strip_plot_titles(p$table)
    }
    return(p)
  }
  
  p
}

render_plot_to_png <- function(plot_obj, png_file, width = 6.8, height = 4.6, dpi = 300) {
  
  p <- plot_obj
  if (is.list(plot_obj) && !inherits(plot_obj, c("gg", "ggsurvplot"))) {
    if (!is.null(plot_obj$plot)) p <- plot_obj$plot
  }
  
  p <- strip_plot_titles(p)
  
  dir.create(dirname(png_file), recursive = TRUE, showWarnings = FALSE)
  
  if (inherits(p, "ggsurvplot")) {
    if (requireNamespace("ragg", quietly = TRUE)) {
      ragg::agg_png(filename = png_file, width = width, height = height, units = "in", res = dpi)
    } else {
      grDevices::png(filename = png_file, width = width, height = height, units = "in", res = dpi)
    }
    on.exit(grDevices::dev.off(), add = TRUE)
    
    gplot  <- p$plot
    gtable <- p$table
    
    if (is.null(gtable)) {
      print(p, newpage = FALSE)
      return(invisible(png_file))
    }
    
    gp <- ggplot2::ggplotGrob(gplot)
    gt <- ggplot2::ggplotGrob(gtable)
    
    maxw <- grid::unit.pmax(gp$widths, gt$widths)
    gp$widths <- maxw
    gt$widths <- maxw
    
    combined <- gridExtra::arrangeGrob(gp, gt, ncol = 1, heights = c(3, 1))
    grid::grid.newpage()
    grid::grid.draw(combined)
    
    return(invisible(png_file))
  }
  
  if (inherits(p, "gg")) {
    ggplot2::ggsave(filename = png_file, plot = p, width = width, height = height, units = "in", dpi = dpi)
    return(invisible(png_file))
  }
  
  stop("Unsupported plot object class: ", paste(class(p), collapse = ", "), call. = FALSE)
}

add_figure_png_to_doc <- function(doc,
                                  fig_title,
                                  analysis_set = "",
                                  png_file,
                                  footnotes = character(0),
                                  img_width = 6.8,
                                  img_height = 4.6,
                                  style_title = "heading 1",
                                  style_text  = "Normal",
                                  font_family = "Courier New",
                                  font_size   = 8,
                                  show_analysis_set_line = TRUE,
                                  repeat_centered_title = TRUE) {
  
  analysis_set_clean <- trimws(as.character(analysis_set %||% ""))
  
  title_for_toc <- if (nzchar(analysis_set_clean)) {
    paste0(fig_title, " (", analysis_set_clean, ")")
  } else {
    fig_title
  }
  
  #doc <- officer::body_add_par(doc, title_for_toc, style = style_title)
  doc <- add_figure_toc_heading_hidden(
    doc = doc,
    fig_title = fig_title,
    analysis_set = analysis_set_clean,
    font_family = font_family,
    style = style_title
  )
  
  if (isTRUE(repeat_centered_title)) {
    spl <- split_table_or_figure_title(fig_title)
   # fp_txt <- officer::fp_text(font.family = font_family, font.size = font_size)
    fp_txt <- officer::fp_text(font.family = font_family, font.size = font_size, bold = FALSE)
    fp_ctr <- officer::fp_par(text.align = "center")
    
    if (nzchar(spl$line1)) {
      doc <- officer::body_add_fpar(doc, officer::fpar(officer::ftext(spl$line1, prop = fp_txt), fp_p = fp_ctr))
    }
    if (nzchar(spl$line2)) {
      doc <- officer::body_add_fpar(doc, officer::fpar(officer::ftext(spl$line2, prop = fp_txt), fp_p = fp_ctr))
    }
  }
  
  if (isTRUE(show_analysis_set_line) && nzchar(analysis_set_clean)) {
    doc <- officer::body_add_fpar(
      doc,
      officer::fpar(
        officer::ftext(analysis_set_clean, officer::fp_text(font.family = font_family, font.size = font_size)),
        fp_p = officer::fp_par(text.align = "center")
      )
    )
  }
  
  doc <- officer::body_add_par(doc, "", style = style_text)
  doc <- officer::body_add_par(doc, "", style = style_text)
  
  img <- officer::external_img(src = png_file, width = img_width, height = img_height)
  
  doc <- officer::body_add_fpar(doc, officer::fpar(img, fp_p = officer::fp_par(text.align = "center")))
  
  # <-- REMOVE the trailing blank line here
  
  if (length(footnotes)) {
    for (fn in footnotes) {
      if (!is.na(fn) && trimws(fn) != "") {
        doc <- officer::body_add_fpar(
          doc,
          officer::fpar(
            officer::ftext(fn, officer::fp_text(font.family = font_family, font.size = font_size))
          )
        )
      }
    }
    # <-- REMOVE the trailing blank line here too
  }
  
  doc
}

read_placeholder_dict <- function(excel_path, dict_sheet = CONST$SHEETS$dict_sheet) {
  sheets <- readxl::excel_sheets(excel_path)
  if (!(dict_sheet %in% sheets)) return(NULL)
  
  dd <- readxl::read_xlsx(excel_path, sheet = dict_sheet)
  names(dd) <- tolower(trimws(names(dd)))
  
  req  <- c("type", "placeholder")
  miss <- setdiff(req, names(dd))
  if (length(miss)) stop("Dictionary sheet '", dict_sheet, "' missing: ", paste(miss, collapse = ", "), call. = FALSE)
  
  dd$type        <- normalize_empty(dd$type)
  dd$placeholder <- normalize_empty(dd$placeholder)
  dd <- dd[dd$type != "", , drop = FALSE]
  if (!nrow(dd)) return(NULL)
  
  dict <- dd$placeholder
  names(dict) <- tolower(dd$type)
  dict
}

# read_column_spec <- function(excel_path, row_sheet_name) {
#   sheets <- readxl::excel_sheets(excel_path)
#   
#   norm <- function(x) {
#     x <- tolower(trimws(x))
#     x <- gsub("\\s+", " ", x)
#     x
#   }
#   nospace <- function(x) gsub("\\s+", "", norm(x))
#   
#   row_nospace <- nospace(row_sheet_name)
#   
#   candidates <- unique(c(
#     paste0(row_sheet_name, "__COLUMNS"),
#     paste0(row_sheet_name, " __COLUMNS"),
#     paste0(row_sheet_name, "_COLUMNS"),
#     paste0(row_sheet_name, " _COLUMNS"),
#     paste0(gsub("\\s+", "", row_sheet_name), "__COLUMNS"),
#     paste0(gsub("\\s+", "", row_sheet_name), "_COLUMNS"),
#     paste0(row_nospace, "__COLUMNS"),
#     paste0(row_nospace, "_COLUMNS")
#   ))
#   
#   spec_sheet <- candidates[candidates %in% sheets][1]
#   
#   if (is.na(spec_sheet) || !nzchar(spec_sheet)) {
#     sheet_map <- data.frame(raw = sheets, n1 = norm(sheets), n2 = nospace(sheets), stringsAsFactors = FALSE)
#     cand_map  <- data.frame(raw = candidates, n1 = norm(candidates), n2 = nospace(candidates), stringsAsFactors = FALSE)
#     
#     hit <- sheet_map$raw[sheet_map$n1 %in% cand_map$n1 | sheet_map$n2 %in% cand_map$n2][1]
#     spec_sheet <- hit
#   }
#   
#   if (is.na(spec_sheet) || !nzchar(spec_sheet)) return(NULL)
#   
#   cs <- readxl::read_xlsx(excel_path, sheet = spec_sheet)
#   names(cs) <- tolower(trimws(names(cs)))
#   
#   req  <- c("col_key", "level2")
#   miss <- setdiff(req, names(cs))
#   if (length(miss)) stop("ColumnSpec sheet '", spec_sheet, "' missing: ", paste(miss, collapse = ", "), call. = FALSE)
#   
#   if (!"level1"  %in% names(cs)) cs$level1  <- ""
#   if (!"level3"  %in% names(cs)) cs$level3  <- ""
#   if (!"width"   %in% names(cs)) cs$width   <- NA_real_
#   if (!"align"   %in% names(cs)) cs$align   <- ""
#   if (!"is_stub" %in% names(cs)) cs$is_stub <- FALSE
#   
#   cs$col_key <- normalize_empty(cs$col_key)
#   cs$level1  <- normalize_empty(cs$level1)
#   cs$level2  <- normalize_empty(cs$level2)
#   cs$level3  <- normalize_empty(cs$level3)
#   cs$align   <- normalize_empty(cs$align)
#   cs$is_stub <- as.logical(cs$is_stub)
#   
#   cs <- cs[cs$col_key != "", , drop = FALSE]
#   if (!nrow(cs)) return(NULL)
#   
#   if (!any(cs$is_stub)) cs$is_stub <- cs$col_key %in% c("Description", "Description2", "Statistic")
#   cs
#}


read_column_spec <- function(excel_path, row_sheet_name) {
  sheets <- readxl::excel_sheets(excel_path)
  
  norm <- function(x) {
    x <- tolower(trimws(x))
    x <- gsub("\\s+", " ", x)
    x
  }
  nospace <- function(x) gsub("\\s+", "", norm(x))
  
  row_nospace <- nospace(row_sheet_name)
  
  # NEW: allow short suffixes too
  suffixes <- c("C", "COL", "COLUMNS")
  joiners  <- c("__", " __", "_", " _")  # keep your existing join styles
  
  # Build candidate names like:
  # "T 3__C", "T 3__COL", "T 3__COLUMNS", "T 3 _COLUMNS", etc.
  base_names <- unique(c(
    row_sheet_name,
    gsub("\\s+", "", row_sheet_name),
    row_nospace
  ))
  
  candidates <- unique(unlist(lapply(base_names, function(bn) {
    unlist(lapply(joiners, function(jn) {
      paste0(bn, jn, suffixes)
    }))
  })))
  
  spec_sheet <- candidates[candidates %in% sheets][1]
  
  if (is.na(spec_sheet) || !nzchar(spec_sheet)) {
    sheet_map <- data.frame(raw = sheets, n1 = norm(sheets), n2 = nospace(sheets), stringsAsFactors = FALSE)
    
    cand_map  <- data.frame(
      raw = candidates,
      n1  = norm(candidates),
      n2  = nospace(candidates),
      stringsAsFactors = FALSE
    )
    
    hit <- sheet_map$raw[sheet_map$n1 %in% cand_map$n1 | sheet_map$n2 %in% cand_map$n2][1]
    spec_sheet <- hit
  }
  
  if (is.na(spec_sheet) || !nzchar(spec_sheet)) return(NULL)
  
  cs <- readxl::read_xlsx(excel_path, sheet = spec_sheet)
  names(cs) <- tolower(trimws(names(cs)))
  
  req  <- c("col_key", "level2")
  miss <- setdiff(req, names(cs))
  if (length(miss)) stop("ColumnSpec sheet '", spec_sheet, "' missing: ", paste(miss, collapse = ", "), call. = FALSE)
  
  if (!"level1"  %in% names(cs)) cs$level1  <- ""
  if (!"level3"  %in% names(cs)) cs$level3  <- ""
  if (!"width"   %in% names(cs)) cs$width   <- NA_real_
  if (!"align"   %in% names(cs)) cs$align   <- ""
  if (!"is_stub" %in% names(cs)) cs$is_stub <- FALSE
  
  cs$col_key <- normalize_empty(cs$col_key)
  cs$level1  <- normalize_empty(cs$level1)
  cs$level2  <- normalize_empty(cs$level2)
  cs$level3  <- normalize_empty(cs$level3)
  cs$align   <- normalize_empty(cs$align)
  cs$is_stub <- as.logical(cs$is_stub)
  
  cs <- cs[cs$col_key != "", , drop = FALSE]
  if (!nrow(cs)) return(NULL)
  
  if (!any(cs$is_stub)) cs$is_stub <- cs$col_key %in% c("Description", "Description2", "Statistic")
  cs
}























colspec_to_header_df <- function(col_spec) {
  hdr <- data.frame(
    col_keys = col_spec$col_key,
    level1   = normalize_empty(col_spec$level1),
    level2   = normalize_empty(col_spec$level2),
    level3   = normalize_empty(col_spec$level3),
    stringsAsFactors = FALSE
  )
  
  idx1 <- hdr$level1 == ""
  hdr$level1[idx1] <- hdr$level2[idx1]
  
  idx2 <- hdr$level2 == ""
  hdr$level2[idx2] <- hdr$level3[idx2]
  
  hdr
}

set_header_labels_if_needed <- function(ft, labels_named) {
  if (is.null(labels_named) || !length(labels_named)) return(ft)
  if (is.null(names(labels_named))) return(ft)
  
  current <- names(ft$body$dataset)
  common  <- intersect(current, names(labels_named))
  if (!length(common)) return(ft)
  
  needs <- any(unname(labels_named[common]) != common)
  if (!needs) return(ft)
  
  do.call(flextable::set_header_labels, c(list(x = ft), as.list(labels_named)))
}

apply_standard_theme <- function(ft) {
  d <- CFG$DISPLAY
  b <- CFG$BORDERS
  
  ft <- ft |>
    flextable::theme_booktabs() |>
    flextable::font(part = "all", fontname = d$font_family) |>
    flextable::fontsize(part = "all", size = d$font_size) |>
    flextable::align(part = "header", align = "center") |>
    flextable::bold(part = "header", bold = TRUE) |>
    flextable::line_spacing(space = d$line_spacing, part = "all") |>
    flextable::padding(part = "all",
                       padding.top = 1.5, padding.bottom = 1.5,
                       padding.left = 2, padding.right = 2) |>
    flextable::bg(part = "header", bg = d$col_header_bg) |>
    flextable::hline_bottom(part = "header", border = b$thick)
  
  ft
}

apply_multilevel_headers <- function(ft, col_spec) {
  hdr_df <- colspec_to_header_df(col_spec)
  ft <- flextable::set_header_df(ft, mapping = hdr_df, key = "col_keys")
  
  header_cols   <- setdiff(names(hdr_df), "col_keys")
  level_has_any <- vapply(header_cols, function(cc) any(normalize_empty(hdr_df[[cc]]) != ""), logical(1))
  used_levels   <- header_cols[level_has_any]
  n_header_rows <- max(1, length(used_levels))
  
  skip_last_hmerge <- length(used_levels) > 0 && tail(used_levels, 1) == "level3"
  
  if (n_header_rows >= 1) {
    for (i in seq_len(n_header_rows)) {
      if (skip_last_hmerge && i == n_header_rows) next
      level_name <- used_levels[i]
      vals <- normalize_empty(hdr_df[[level_name]])
      uniq_nonempty <- unique(vals[vals != ""])
      if (length(uniq_nonempty) == 1 && length(vals) > 1) next
      ft <- flextable::merge_h(ft, part = "header", i = i)
    }
  }
  
  ft <- flextable::font(ft, part = "header", fontname = CFG$DISPLAY$font_family)
  ft <- flextable::fontsize(ft, part = "header", size = CFG$DISPLAY$font_size)
  ft <- flextable::hline_bottom(ft, part = "header", border = CFG$BORDERS$thick)
  
  if ("level2" %in% used_levels) {
    level2_row <- match("level2", used_levels)
    if (is.finite(level2_row) && !is.na(level2_row)) {
      ft <- flextable::valign(ft, part = "header", i = level2_row, valign = "top")
    }
  }
  
  for (j in seq_len(nrow(col_spec))) {
    w <- col_spec$width[j]
    if (!is.na(w) && is.finite(w)) ft <- flextable::width(ft, j = j, width = w)
  }
  
  for (j in seq_len(nrow(col_spec))) {
    a <- tolower(col_spec$align[j])
    if (a %in% c("left", "center", "right")) {
      ft <- flextable::align(ft, j = j, part = "body", align = a)
    }
  }
  
  ft
}

apply_hline_rows <- function(ft, rows_spec, thick_border) {
  idx <- which(norm_key(rows_spec$type) == CONST$TYPES$hline)
  if (!length(idx)) return(ft)
  for (i in idx) ft <- flextable::hline(ft, i = i, border = thick_border, part = "body")
  ft
}

apply_toprule_if_requested <- function(ft, rows_spec, thick_border) {
  has_toprule <- any(norm_key(rows_spec$type) == CONST$TYPES$toprule)
  if (!isTRUE(has_toprule)) return(ft)
  flextable::hline_top(ft, part = "header", border = thick_border)
}

extract_footnotes <- function(rows_spec, placeholder_dict = NULL) {
  if (is.null(rows_spec) || !nrow(rows_spec)) return(character(0))
  if (!("type" %in% names(rows_spec)) || !("desc" %in% names(rows_spec))) return(character(0))
  
  idx <- which(norm_key(rows_spec$type) == CONST$TYPES$footnote)
  if (!length(idx)) return(character(0))
  
  keys <- normalize_empty(rows_spec$desc[idx])
  keys <- keys[keys != ""]
  if (!length(keys)) return(character(0))
  
  out <- character(0)
  for (k in keys) {
    kk <- tolower(k)
    if (!is.null(placeholder_dict) && length(placeholder_dict) > 0 &&
        kk %in% names(placeholder_dict) && normalize_empty(placeholder_dict[[kk]]) != "") {
      out <- c(out, normalize_empty(placeholder_dict[[kk]]))
    } else {
      out <- c(out, k)
    }
  }
  
  out <- normalize_empty(out)
  out[out != ""]
}

remove_footer_borders <- function(ft) {
  if ("part" %in% names(formals(flextable::border_remove))) {
    return(flextable::border_remove(ft, part = "footer"))
  }
  zero <- fp_border(color = "#FFFFFF", width = 0)
  flextable::border(ft, part = "footer", border = zero)
}

add_footnotes_footer <- function(ft, footnotes) {
  if (is.null(footnotes) || !length(footnotes)) return(ft)
  
  for (txt in footnotes) ft <- flextable::add_footer_lines(ft, values = txt)
  ft <- flextable::merge_h(ft, part = "footer")
  ft <- flextable::merge_v(ft, part = "footer")
  ft <- flextable::align(ft, part = "footer", align = "left")
  ft <- flextable::font(ft, part = "footer", fontname = CFG$DISPLAY$font_family)
  ft <- flextable::fontsize(ft, part = "footer", size = CFG$DISPLAY$font_size)
  ft <- flextable::bold(ft, part = "footer", bold = FALSE)
  ft <- remove_footer_borders(ft)
  ft <- flextable::padding(ft, part = "footer",
                           padding.top = 2, padding.bottom = 0,
                           padding.left = 2, padding.right = 2)
  ft
}

scale_col_widths_to_fit <- function(col_spec, max_width_in) {
  if (is.null(col_spec)) return(col_spec)
  if (!("width" %in% names(col_spec))) return(col_spec)
  if (is.null(col_spec$width)) return(col_spec)
  
  w <- suppressWarnings(as.numeric(col_spec$width))
  if (all(is.na(w))) return(col_spec)
  
  idx <- which(!is.na(w) & is.finite(w) & w > 0)
  if (!length(idx)) return(col_spec)
  
  total <- sum(w[idx])
  if (!is.finite(total) || total <= 0) return(col_spec)
  
  if (is.finite(max_width_in) && total > max_width_in) {
    factor <- max_width_in / total
    w[idx] <- w[idx] * factor
    col_spec$width <- w
  }
  
  col_spec
}

parse_blank_cols <- function(spec, col_keys, data_cols) {
  spec <- trimws(as.character(spec))
  if (!nzchar(spec)) return(character(0))
  spec_l <- tolower(spec)
  
  if (spec_l %in% c("data", "datacols", "all", "*")) {
    return(intersect(data_cols, col_keys))
  }
  
  parts <- unlist(strsplit(spec, ","))
  parts <- trimws(parts)
  parts <- parts[nzchar(parts)]
  if (!length(parts)) return(character(0))
  
  out <- character(0)
  for (p in parts) {
    
    if (grepl("^\\d+\\s*[-:]\\s*\\d+$", p)) {
      ab <- unlist(strsplit(gsub("\\s+", "", p), "[-:]"))
      a <- as.integer(ab[1]); b <- as.integer(ab[2])
      if (is.na(a) || is.na(b)) next
      idx <- seq(min(a, b), max(a, b))
      idx <- idx[idx >= 1 & idx <= length(col_keys)]
      out <- c(out, col_keys[idx])
      next
    }
    
    if (grepl("^\\d+$", p)) {
      k <- as.integer(p)
      if (!is.na(k) && k >= 1 && k <= length(col_keys)) out <- c(out, col_keys[k])
      next
    }
    
    if (grepl("^[^:]+:[^:]+$", p)) {
      lr <- unlist(strsplit(p, ":", fixed = TRUE))
      left <- trimws(lr[1]); right <- trimws(lr[2])
      i1 <- match(left, col_keys)
      i2 <- match(right, col_keys)
      if (!is.na(i1) && !is.na(i2)) {
        idx <- seq(min(i1, i2), max(i1, i2))
        out <- c(out, col_keys[idx])
      }
      next
    }
    
    if (p %in% col_keys) out <- c(out, p)
  }
  
  unique(out)
}

table_spec_to_body <- function(rows_spec, col_spec, placeholder_dict = NULL) {
  
  col_keys <- if (is.null(col_spec)) CONST$DEFAULTS$default_col_keys else col_spec$col_key
  n <- nrow(rows_spec)
  
  if (n == 0) {
    return(tibble::as_tibble(setNames(replicate(length(col_keys), character(0), simplify = FALSE), col_keys)))
  }
  
  types <- norm_key(rows_spec$type)
  desc  <- normalize_empty(rows_spec$desc)
  
  out <- as.data.frame(
    setNames(replicate(length(col_keys), rep("", n), simplify = FALSE), col_keys),
    stringsAsFactors = FALSE
  )
  
  is_hline <- types == CONST$TYPES$hline
  
  if ("Description" %in% col_keys)
    out$Description[!is_hline] <- desc[!is_hline]
  
  desc2 <- if ("desc2" %in% names(rows_spec)) normalize_empty(rows_spec$desc2) else rep("", n)
  if ("Description2" %in% col_keys)
    out$Description2[!is_hline] <- desc2[!is_hline]
  
  if ("Statistic" %in% col_keys)
    out$Statistic[] <- ""
  
  special <- c(CONST$TYPES$section, CONST$TYPES$blankstat, CONST$TYPES$pagebreak,
               CONST$TYPES$hline, CONST$TYPES$toprule, CONST$TYPES$footnote)
  is_data_row <- !(types %in% special)
  
  data_cols <- if (!is.null(col_spec)) {
    col_spec$col_key[!col_spec$is_stub]
  } else {
    setdiff(col_keys, c("Description", "Description2", "Statistic"))
  }
  data_cols <- intersect(data_cols, col_keys)
  
  if (!length(data_cols) || !any(is_data_row)) {
    return(tibble::as_tibble(out))
  }
  
  ph <- rep(CONST$DEFAULTS$placeholder_npct, n)
  
  if (!is.null(placeholder_dict) && length(placeholder_dict) > 0) {
    idx <- which(is_data_row & types != "" & types %in% names(placeholder_dict))
    if (length(idx)) {
      mapped <- as.character(placeholder_dict[types[idx]])
      is_empty_type <- (types[idx] == CONST$TYPES$empty)
      
      mapped[!is_empty_type & (is.na(mapped) | mapped == "")] <- CONST$DEFAULTS$placeholder_npct
      mapped[is_empty_type] <- ""
      ph[idx] <- mapped
    }
  }
  
  for (cc in data_cols) out[[cc]][is_data_row] <- ph[is_data_row]
  
  if ("blank_cols" %in% names(rows_spec)) {
    blank_spec <- normalize_empty(rows_spec$blank_cols)
    if (any(blank_spec != "")) {
      for (i in which(blank_spec != "")) {
        cols_to_blank <- parse_blank_cols(blank_spec[i], col_keys = col_keys, data_cols = data_cols)
        cols_to_blank <- intersect(cols_to_blank, names(out))
        if (length(cols_to_blank)) {
          for (cc in cols_to_blank) out[[cc]][i] <- ""
        }
      }
    }
  }
  
  tibble::as_tibble(out)
}

add_title_block <- function(doc, title, analysis_set,
                            show_analysis_set_line = TRUE,
                            toc_sep = " (", toc_end = ")") {
  
  analysis_set_clean <- trimws(as.character(analysis_set %||% ""))
  
  title_for_toc <- if (nzchar(analysis_set_clean)) {
    paste0(title, toc_sep, analysis_set_clean, toc_end)
  } else {
    title
  }
  
  doc <- officer::body_add_par(doc, value = title_for_toc, style = "heading 1")
  
  if (isTRUE(show_analysis_set_line) && nzchar(analysis_set_clean)) {
    doc <- officer::body_add_fpar(
      doc,
      officer::fpar(
        officer::ftext(
          analysis_set_clean,
          officer::fp_text(
            font.size   = CFG$DISPLAY$font_size,
            font.family = CFG$DISPLAY$font_family
          )
        ),
        fp_p = officer::fp_par(text.align = "center")
      )
    )
  }
  
  doc <- officer::body_add_par(doc, "", style = "Normal")
  doc
}


add_repeating_table_title_header <- function(ft, title, analysis_set = "",
                                             ncols,
                                             font_family = CFG$DISPLAY$font_family,
                                             font_size   = CFG$DISPLAY$font_size,
                                             make_top_line_bold = FALSE,
                                             add_rule_under_block = TRUE) {
  
  spl <- split_table_or_figure_title(title)
  as_clean <- trimws(as.character(analysis_set %||% ""))
  
  lines <- c(spl$line1, spl$line2, as_clean)
  lines <- lines[nzchar(trimws(lines))]
  if (!length(lines)) return(ft)
  
  ncols <- as.integer(ncols)
  if (!is.finite(ncols) || ncols <= 0) return(ft)
  
  # Add lines so final order at top is: line1, line2, analysis_set
  for (k in rev(lines)) {
    ft <- flextable::add_header_row(
      ft,
      values    = rep(k, ncols),
      colwidths = rep(1, ncols)
    )
  }
  
  n_added <- length(lines)
  
  # Format those newly-added title rows (they are now header rows 1:n_added)
  for (i in seq_len(n_added)) {
    ft <- flextable::merge_h(ft, part = "header", i = i)
    ft <- flextable::align(ft, part = "header", i = i, align = "center")
    ft <- flextable::font(ft, part = "header", i = i, fontname = font_family)
    ft <- flextable::fontsize(ft, part = "header", i = i, size = font_size)
    
    # Force title lines NOT bold and NOT shaded
    ft <- flextable::bold(ft, part = "header", i = i, bold = FALSE)
    ft <- flextable::bg(ft,   part = "header", i = i, bg = "white")
  }
  
  # (Optional) if you ever want top line bold again, keep the hook:
  if (isTRUE(make_top_line_bold)) {
    ft <- flextable::bold(ft, part = "header", i = 1, bold = TRUE)
  }
  
  if (isTRUE(add_rule_under_block)) {
    ft <- flextable::hline(ft, part = "header", i = n_added, border = CFG$BORDERS$thick)
  }
  
  ft
}





apply_desc_indent <- function(ft, rows_spec, indent_col = "desc_indent",
                              j = "Description",
                              base_padding_left = 2, indent_step = 10,
                              valid_cols = NULL) {
  
  if (is.null(valid_cols)) {
    valid_cols <- tryCatch(names(ft$body$dataset), error = function(e) character(0))
  }
  if (!(indent_col %in% names(rows_spec))) return(ft)
  if (!length(valid_cols) || !(j %in% valid_cols)) return(ft)
  
  ind <- suppressWarnings(as.numeric(rows_spec[[indent_col]]))
  ind[is.na(ind)] <- 0
  ind <- pmax(0, ind)
  
  for (i in seq_along(ind)) {
    ft <- flextable::padding(
      ft, i = i, j = j, part = "body",
      padding.left = base_padding_left + indent_step * ind[i]
    )
  }
  ft
}




add_figure_toc_heading_hidden <- function(doc, fig_title, analysis_set = "",
                                          font_family = CFG$DISPLAY$font_family,
                                          toc_font_size = 1,
                                          style = "heading 1") {
  
  analysis_set_clean <- trimws(as.character(analysis_set %||% ""))
  
  spl  <- split_table_or_figure_title(fig_title)
  base <- trimws(paste(spl$line1, spl$line2))
  
  toc_text <- if (nzchar(analysis_set_clean)) {
    paste0(base, " (", analysis_set_clean, ")")
  } else {
    base
  }
  
  fp_txt <- officer::fp_text(
    font.family = font_family,
    font.size   = toc_font_size,
    bold        = FALSE,
    color       = "white"
  )
  
  fp_par <- officer::fp_par(padding = 0, line_spacing = 0.6)
  
  doc <- officer::body_add_fpar(
    doc,
    officer::fpar(officer::ftext(toc_text, prop = fp_txt), fp_p = fp_par),
    style = style
  )
  
  doc
}

add_table_toc_heading_hidden <- function(doc, title, analysis_set = "",
                                         font_family = CFG$DISPLAY$font_family,
                                         toc_font_size = 1,
                                         style = "heading 1") {
  analysis_set_clean <- trimws(as.character(analysis_set %||% ""))
  
  # Build a single-line TOC string from your 3-level block
  spl <- split_table_or_figure_title(title)
  base <- trimws(paste(spl$line1, spl$line2))
  toc_text <- if (nzchar(analysis_set_clean)) {
    paste0(base, " (", analysis_set_clean, ")")
  } else {
    base
  }
  
  # Make it "invisible": tiny + white + not bold
  fp_txt <- officer::fp_text(
    font.family = font_family,
    font.size   = toc_font_size,
    bold        = FALSE,
    color       = "white"
  )
  
  # Keep paragraph tight (minimises vertical footprint)
  fp_par <- officer::fp_par(padding = 0, line_spacing = 0.6)
  
  doc <- officer::body_add_fpar(
    doc,
    officer::fpar(officer::ftext(toc_text, prop = fp_txt), fp_p = fp_par),
    style = style
  )
  
  doc
}

make_table <- function(doc, title, analysis_set, rows_spec, col_spec = NULL,
                       placeholder_dict = NULL) {
  
  has_toprule <- any(norm_key(rows_spec$type) == CONST$TYPES$toprule)
  footnotes   <- extract_footnotes(rows_spec, placeholder_dict = placeholder_dict)
  
  rows_spec2 <- rows_spec |>
    dplyr::filter(!(norm_key(type) %in% c(CONST$TYPES$pagebreak, CONST$TYPES$toprule, CONST$TYPES$footnote)))
  
  if (!is.null(col_spec)) {
    col_spec <- scale_col_widths_to_fit(col_spec, max_width_in = usable_width_in())
  }
  
  tbl_body <- table_spec_to_body(rows_spec2, col_spec = col_spec, placeholder_dict = placeholder_dict)
  
  if (!is.null(col_spec)) {
    needed <- col_spec$col_key
    tbl_body <- tbl_body[, needed, drop = FALSE]
  }
  
  ft <- flextable::flextable(tbl_body)
  ft <- apply_standard_theme(ft)
  
  if (!is.null(col_spec)) {
    ft <- apply_multilevel_headers(ft, col_spec)
  } else {
    ft <- set_header_labels_if_needed(ft, CONST$DEFAULTS$default_header_labels)
  }
  
  ft <- add_repeating_table_title_header(
    ft,
    title        = title,
    analysis_set = analysis_set,
    ncols        = ncol(tbl_body),
    font_family  = CFG$DISPLAY$font_family,
    font_size    = CFG$DISPLAY$font_size
  )
  
#  spl <- split_table_or_figure_title(title)  # same helper you already use
#ft  <- bold_table_id_line(ft, line1_text = spl$line1, ncols = ncol(tbl_body))
  # DEBUG: inspect header dataset (remove after)
  #print(ft$header$dataset)
  
  # Re-enforce bold for the top repeating title line
 # ft <- flextable::bold(ft, part = "header", i = 1, bold = TRUE)  # new to have bold titles on tables!
  
  ft <- flextable::align(ft, part = "header", align = "center")
  
  if (isTRUE(has_toprule)) {
    ft <- apply_toprule_if_requested(ft, rows_spec, CFG$BORDERS$thick)
  }
  
  ft <- apply_hline_rows(ft, rows_spec2, CFG$BORDERS$thick)
  
  if (nrow(tbl_body) > 0) {
    ft <- flextable::hline(ft, i = nrow(tbl_body), border = CFG$BORDERS$thick, part = "body")
  }
  
  ft <- add_footnotes_footer(ft, footnotes)
  
  ft <- apply_desc_indent(ft, rows_spec2, indent_col = "desc_indent",
                          j = "Description", base_padding_left = 2, indent_step = 10,
                          valid_cols = names(tbl_body))
  
  if ("Description" %in% names(tbl_body))
    ft <- flextable::align(ft, j = "Description", part = "body", align = "left")
  if ("Description2" %in% names(tbl_body))
    ft <- flextable::align(ft, j = "Description2", part = "body", align = "left")
  if ("Statistic" %in% names(tbl_body))
    ft <- flextable::align(ft, j = "Statistic", part = "body", align = "center")
  
  other_cols <- setdiff(names(tbl_body), c("Description", "Description2", "Statistic"))
  if (length(other_cols) > 0) {
    ft <- flextable::align(ft, j = other_cols, part = "body", align = "center")
  }
  
  ft <- flextable::set_table_properties(
    ft,
    layout = "fixed",
    width  = 1,
    align  = "center",
    opts_word = list(repeat_headers = TRUE, split = TRUE)
  )
  


  # Add a hidden Heading 1 so the TOC picks up the table entry
  doc <- add_table_toc_heading_hidden(doc, title = title, analysis_set = analysis_set)
  mark_content_added()
  
  # (Optional) small spacer – you may not need it anymore
  # doc <- officer::body_add_par(doc, "", style = "Normal")
  # mark_content_added()
  
  doc <- flextable::body_add_flextable(doc, ft)
  mark_content_added()
  
  
    
  sec_tbl <- make_table_section_props(
    title = title,
    analysis_set = analysis_set,
    base_header_block = hdr_block,
    base_footer_block = ftr_block
  )
  
  doc <- officer::body_end_block_section(doc, value = officer::block_section(sec_tbl))
  mark_content_added()
  
  doc
}

read_and_create_table <- function(doc, excel_path, sheet_name, placeholder_dict = NULL) {
  rows_spec <- readxl::read_xlsx(excel_path, sheet = sheet_name)
  names(rows_spec) <- tolower(trimws(names(rows_spec)))
  
  if (!all(c("type", "desc") %in% names(rows_spec))) {
    warning("Skipping sheet '", sheet_name, "' (missing type/desc).")
    return(doc)
  }
  
  title <- if ("title" %in% names(rows_spec)) as.character(rows_spec$title[1]) else sheet_name
  if (is.na(title) || title == "") title <- sheet_name
  
  analysis_set <- if ("analysis set" %in% names(rows_spec)) as.character(rows_spec[["analysis set"]][1]) else ""
  if (is.na(analysis_set)) analysis_set <- ""
  
  col_spec <- read_column_spec(excel_path, sheet_name)
  
  doc <- make_table(
    doc              = doc,
    title            = title,
    analysis_set     = analysis_set,
    rows_spec        = rows_spec,
    col_spec         = col_spec,
    placeholder_dict = placeholder_dict
  )
  
  doc
}

# ==========================================================
# MAIN
# ==========================================================
dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

placeholder_dict <- read_placeholder_dict(excel_path, dict_sheet = DICT_SHEET)

doc <- officer::read_docx()

doc <- officer::body_set_default_section(doc, value = sec_landscape)

doc <- officer::body_add_par(doc, "Contents", style = "heading 1"); mark_content_added()


doc <- ensure_toc_styles(
  doc,
  levels      = 3,                     # or 1 if you truly only want TOC 1
  font_family = CFG$DISPLAY$font_family,
  font_size   = CFG$DISPLAY$font_size
)


doc <- officer::body_add_toc(doc, level = 1); mark_content_added()
doc <- officer::body_add_par(doc, "", style = "Normal"); mark_content_added()



doc <- officer::body_end_block_section(doc, value = officer::block_section(sec_landscape_toc))

# Force new page after TOC (controlled in R)
doc <- force_page_break(doc)


#----------------------
# After TOC section ends
doc <- officer::body_end_block_section(doc, value = officer::block_section(sec_landscape_toc))

# Force new page after TOC (controlled in R)
doc <- force_page_break(doc)
mark_content_added()

# ---- TABLES cover page ----
doc <- add_section_cover_page(doc, "Tables")
#----------------------


sheets <- readxl::excel_sheets(excel_path)

# ---- TABLES ----
row_sheets <- sheets[grepl(CONST$REGEX$table_sheets, sheets, ignore.case = TRUE)]
if (!length(row_sheets)) {
  stop("No table sheets matched table_sheets regex. Rename sheets like 'TABLE 1' / 'TABLE 14.1.1.1'.", call. = FALSE)
}

message("Processing row-spec sheets:")
message(paste(" -", row_sheets, collapse = "\n"))

first_table <- TRUE
for (sh in row_sheets) {
  
  # Add page break BEFORE each table EXCEPT the very first one
  if (!first_table) {
    doc <- force_page_break(doc)
  }
  first_table <- FALSE
  
  doc <- read_and_create_table(doc, excel_path, sh, placeholder_dict = placeholder_dict)
}

# ---- FIGURES ----
#fig_sheets <- sheets[grepl("^FIG_\\d+$", sheets, ignore.case = TRUE)]
fig_sheets <- sheets[grepl(CONST$REGEX$figure_sheets, sheets, ignore.case = TRUE)]

# if (length(fig_sheets)) {
#   
#   # Use safe break instead of always forcing one
#   doc <- add_page_break_safe(doc)
#   
#   doc <- officer::body_add_par(doc, "Figures", style = "heading 1"); mark_content_added()
#   doc <- officer::body_add_par(doc, "", style = "Normal"); mark_content_added()
# }

##---new

if (length(fig_sheets)) {
  doc <- add_section_cover_page(doc, "Figures")
}


##---



fig_out_dir <- file.path(out_dir, "figures")
dir.create(fig_out_dir, recursive = TRUE, showWarnings = FALSE)

first_figure <- TRUE
if (length(fig_sheets)) {
  for (fs in fig_sheets) {
    
    # Add page break BEFORE each figure EXCEPT the very first one
    if (!first_figure) {
      doc <- force_page_break(doc)
    }
    first_figure <- FALSE
    
    raw_spec <- read_fig_kv_spec(excel_path, sheet = fs)
    fig_spec <- normalise_fig_spec(raw_spec)
    
    fig_obj <- build_figure(fig_spec)
    
    fig_fns <- resolve_fig_footnotes(fig_spec$footnotes, placeholder_dict = placeholder_dict)
    
    png_file <- file.path(fig_out_dir, paste0(fs, ".png"))
    
    w <- if (!is.na(fig_spec$fig_width))  fig_spec$fig_width  else fig_target_width_in()
    h <- if (!is.na(fig_spec$fig_height)) fig_spec$fig_height else fig_target_height_in(w)
    
    wh <- cap_fig_wh_to_max_height(w, h)
    w  <- wh$w
    h  <- wh$h
    
    render_plot_to_png(fig_obj, png_file, width = w, height = h, dpi = CFG$FIG$dpi)
    
    doc <- add_figure_png_to_doc(
      doc          = doc,
      fig_title    = fig_spec$figure_title %||% fs,
      analysis_set = fig_spec$analysis_set %||% "",
      png_file     = png_file,
      footnotes    = fig_fns,
      img_width    = w,
      img_height   = h,
      font_family  = CFG$DISPLAY$font_family,
      font_size    = CFG$DISPLAY$font_size
    )
    
    mark_content_added()
  }
}

# ---- LISTINGS ----
# listing_sheets <- sheets[grepl(CONST$REGEX$listing_sheets, sheets, ignore.case = TRUE)]
# 
# if (length(listing_sheets)) {
#   
#   # Start Listings section on a new page (safe to avoid double blank pages)
#   doc <- add_page_break_safe(doc)
#   
#   # Section heading
#   doc <- officer::body_add_par(doc, "Listings", style = "heading 1"); mark_content_added()
#   doc <- officer::body_add_par(doc, "", style = "Normal"); mark_content_added()
#   
#   first_listing <- TRUE
#   for (sh in listing_sheets) {
#     
#     # Page break BEFORE each listing except first listing in the section
#     if (!first_listing) {
#       doc <- force_page_break(doc)
#     }
#     first_listing <- FALSE
#     
#     # Reuse the same table machinery
#     doc <- read_and_create_table(doc, excel_path, sh, placeholder_dict = placeholder_dict)
#   }
# }


# ---- LISTINGS ----
listing_sheets <- sheets[grepl(CONST$REGEX$listing_sheets, sheets, ignore.case = TRUE)]

if (length(listing_sheets)) {
  doc <- add_section_cover_page(doc, "Listings")
  
  first_listing <- TRUE
  for (sh in listing_sheets) {
    if (!first_listing) doc <- force_page_break(doc)
    first_listing <- FALSE
    
    doc <- read_and_create_table(doc, excel_path, sh, placeholder_dict = placeholder_dict)
  }
}


# ---- Write output ----
timestamp <- format(Sys.time(), "%Y-%m-%d_%H-%M-%S")
out_file  <- file.path(out_dir, paste0(out_prefix, timestamp, ".docx"))

message("Writing to: ", normalizePath(out_file, winslash = "/", mustWork = FALSE))
safe_write_docx(doc, out_file)

message("DONE: wrote ", normalizePath(out_file, winslash = "/", mustWork = TRUE))
message("Size (bytes): ", file.info(out_file)$size)

try(shell.exec(normalizePath(out_file, winslash = "\\", mustWork = TRUE)), silent = TRUE)