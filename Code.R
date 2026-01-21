
# 1 Libraries

# Keep packages minimal and common to support portability
pkgs <- c("tidyverse", "readODS", "janitor", "lubridate", "stringr")

# Stop early if any packages are missing, instead of failing later
missing <- pkgs[!sapply(pkgs, requireNamespace, quietly = TRUE)]
if (length(missing) > 0) {
  stop(
    paste0(
      "Missing packages: ", paste(missing, collapse = ", "),
      "\nInstall them once with install.packages(c(",
      paste0('"', missing, '"', collapse = ","),
      "))"
    ),
    call. = FALSE
  )
}
invisible(lapply(pkgs, library, character.only = TRUE))

# 2 File paths

# Assumption: the three .ods files are stored in the same folder as this script
file_3123 <- "table-3123-trains-planned-and-cancellations-by-operator-and-cause.ods"
file_3133 <- "table-3133-train-punctuality-at-recorded-station-stops-by-operator.ods"
file_1220 <- "table-1220-passenger-journeys.ods"

# 3 Output folders

# Create a clean outputs structure so results are easy to find and commit to GitHub
dir.create("outputs", showWarnings = FALSE)
dir.create("outputs/plots", showWarnings = FALSE)
dir.create("outputs/tables", showWarnings = FALSE)
dir.create("outputs/data", showWarnings = FALSE)

# 4 Helper functions

# Convert mixed text to numeric safely (e.g., "1,234", "12.3%", "n/a")
to_num <- function(x) suppressWarnings(readr::parse_number(as.character(x)))

# Remove bracketed notes commonly present in ORR tables (e.g., "[r]", "[p]")
drop_brackets <- function(x) {
  x %>%
    as.character() %>%
    stringr::str_replace_all("\\s*\\[[^\\]]+\\]", "") %>%
    stringr::str_squish()
}

# Pick the first column name that matches a regex pattern
# This helps the script survive small naming variations after janitor::clean_names()
pick1 <- function(df, pattern) {
  hit <- names(df)[stringr::str_detect(names(df), pattern)]
  if (length(hit) == 0) stop(paste0("Missing column matching: ", pattern), call. = FALSE)
  hit[1]
}

# Parse the ORR quarterly label to a quarter start date
# Example inputs often look like "Apr to Jun 2019"
parse_period_start_quarter <- function(tp) {
  tp <- drop_brackets(tp)
  m <- stringr::str_match(tp, "^([A-Za-z]{3})")[, 2]
  y <- stringr::str_match(tp, "([0-9]{4})$")[, 2]
  as.Date(sprintf("%s-%02d-01", y, match(m, month.abb)))
}

# Parse the ORR financial-year label to the FY start date
# Example inputs often look like "Apr 2019 to Mar 2020"
parse_period_start_finyear <- function(tp) {
  tp <- drop_brackets(tp)
  m <- stringr::str_match(tp, "^([A-Za-z]{3})")[, 2]
  y_first <- stringr::str_match(tp, "([0-9]{4})")[, 2]
  as.Date(sprintf("%s-%02d-01", y_first, match(m, month.abb)))
}

# Convert a date to UK financial year label (Apr–Mar)
fy_label <- function(d) {
  y <- lubridate::year(d)
  m <- lubridate::month(d)
  fy0 <- dplyr::if_else(m >= 4, y, y - 1)
  paste0(fy0, "/", substr(fy0 + 1, 3, 4))
}

# Consistent plot style across all figures
theme_report <- function() {
  ggplot2::theme_minimal(base_size = 12) +
    ggplot2::theme(
      plot.title.position = "plot",
      plot.caption.position = "plot",
      panel.grid.minor = ggplot2::element_blank()
    )
}

# HTML escaping for Word-openable HTML tables
escape_html <- function(x) {
  x <- as.character(x)
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  x
}

# Turn a data frame into a simple HTML table that Word can render
df_to_html_table <- function(df, digits = NULL) {
  d <- df
  
  # Optional rounding (only for specified numeric columns)
  if (!is.null(digits)) {
    for (nm in names(digits)) {
      if (nm %in% names(d) && is.numeric(d[[nm]])) d[[nm]] <- round(d[[nm]], digits[[nm]])
    }
  }
  
  # Escape for safe HTML rendering
  for (j in seq_along(d)) d[[j]] <- escape_html(d[[j]])
  
  head_cells <- paste0("<th>", names(d), "</th>", collapse = "")
  body_rows <- apply(d, 1, function(r) {
    paste0("<tr>", paste0("<td>", r, "</td>", collapse = ""), "</tr>")
  })
  
  paste0(
    "<table class='tbl'>",
    "<thead><tr>", head_cells, "</tr></thead>",
    "<tbody>", paste(body_rows, collapse = ""), "</tbody>",
    "</table>"
  )
}


# 5 Read and clean data

# 5.1 ORR 3123 Cancellations

# Sheet name is fixed to the ORR file layout
canc_raw <- readODS::read_ods(file_3123, sheet = "3123_Cancellations_by_operator", skip = 3) %>%
  tibble::as_tibble() %>%
  janitor::clean_names()

# Identify required columns (robust to minor name changes)
c_tp   <- pick1(canc_raw, "^time_period$")
c_op   <- pick1(canc_raw, "^national_or_operator$")
c_plan <- pick1(canc_raw, "^number_of_trains_planned$")
c_ctot <- pick1(canc_raw, "^cancellations$")
c_rate <- pick1(canc_raw, "^quarterly_cancellations.*percentage$")

# Responsibility splits (used later for clustering features)
c_inm  <- pick1(canc_raw, "^cancellations_by_responsibility_infrastructure.*network_management$")
c_ioe  <- pick1(canc_raw, "^cancellations_by_responsibility_infrastructure_owner.*external_event$")
c_tof  <- pick1(canc_raw, "^cancellations_by_responsibility_train_operator_fault$")
c_oev  <- pick1(canc_raw, "^cancellations_by_responsibility_operator_external_event$")

# Clean cancellations into a tidy quarterly table
canc <- canc_raw %>%
  dplyr::transmute(
    time_period   = drop_brackets(.data[[c_tp]]),
    operator      = drop_brackets(.data[[c_op]]),
    trains_planned = to_num(.data[[c_plan]]),
    canc_total     = to_num(.data[[c_ctot]]),
    canc_infra_nm  = to_num(.data[[c_inm]]),
    canc_infra_owner_ext = to_num(.data[[c_ioe]]),
    canc_operator_fault  = to_num(.data[[c_tof]]),
    canc_operator_ext    = to_num(.data[[c_oev]]),
    cancel_rate    = to_num(.data[[c_rate]]),
    quarter_start  = parse_period_start_quarter(time_period),
    fy             = fy_label(quarter_start)
  ) %>%
  # Shares are only meaningful when cancellations exist
  dplyr::mutate(
    canc_infra_nm_share = dplyr::if_else(canc_total > 0, canc_infra_nm / canc_total, NA_real_),
    canc_operator_fault_share = dplyr::if_else(canc_total > 0, canc_operator_fault / canc_total, NA_real_)
  )



# 5.2 ORR 3133 Punctuality

pun_raw <- readODS::read_ods(file_3133, sheet = "3133_Punctuality_(quarterly)", skip = 4) %>%
  tibble::as_tibble() %>%
  janitor::clean_names()

p_tp    <- pick1(pun_raw, "^time_period$")
p_op    <- pick1(pun_raw, "^national_or_operator$")
p_stops <- pick1(pun_raw, "^number_of_recorded.*stops$")
p_3     <- pick1(pun_raw, "^trains_arriving_within_3.*percentage$")

pun <- pun_raw %>%
  dplyr::transmute(
    time_period   = drop_brackets(.data[[p_tp]]),
    operator      = drop_brackets(.data[[p_op]]),
    stops         = to_num(.data[[p_stops]]),
    punct_3min    = to_num(.data[[p_3]]),
    quarter_start = parse_period_start_quarter(time_period),
    fy            = fy_label(quarter_start)
  )



# 5.3 ORR 1220 Passenger journeys

# Table 1220 is published at financial-year granularity
jour_raw <- readODS::read_ods(file_1220, sheet = "1220_Passenger_journeys", skip = 4) %>%
  tibble::as_tibble() %>%
  janitor::clean_names()

j_tp  <- pick1(jour_raw, "^time_period$")
j_val <- pick1(jour_raw, "^passenger_journeys")

jour <- jour_raw %>%
  dplyr::transmute(
    time_period = drop_brackets(.data[[j_tp]]),
    journeys_m  = to_num(.data[[j_val]])
  ) %>%
  # Keep only Apr–Mar financial-year rows
  dplyr::filter(stringr::str_detect(time_period, "^Apr\\b") & stringr::str_detect(time_period, "to\\s+Mar")) %>%
  dplyr::mutate(
    fy_start_date = parse_period_start_finyear(time_period),
    fy = fy_label(fy_start_date)
  )



# 6 Build analysis datasets


# Operator-quarter panel for modelling and operator benchmarking
panel <- canc %>%
  dplyr::inner_join(pun, by = c("operator", "time_period", "quarter_start", "fy")) %>%
  dplyr::filter(quarter_start >= as.Date("2019-04-01")) %>%
  dplyr::arrange(operator, quarter_start)

# Export cleaned panel for transparency and reuse
readr::write_csv(panel, "outputs/data/operator_quarter_panel.csv")



# 7 Tables and methods

# 7.1 Table 1 Operator benchmark (weighted)

# Weighting avoids over-emphasising small quarters
op_summary <- panel %>%
  dplyr::filter(!operator %in% c("England and Wales", "Scotland")) %>%
  dplyr::group_by(operator) %>%
  dplyr::summarise(
    quarters = dplyr::n(),
    trains_planned = sum(trains_planned, na.rm = TRUE),
    canc_total = sum(canc_total, na.rm = TRUE),
    cancel_rate_w = 100 * canc_total / trains_planned,
    stops = sum(stops, na.rm = TRUE),
    punct_3min_w = sum(stops * punct_3min, na.rm = TRUE) / stops,
    .groups = "drop"
  ) %>%
  dplyr::mutate(
    cancel_rate_w = round(cancel_rate_w, 2),
    punct_3min_w  = round(punct_3min_w, 2)
  ) %>%
  dplyr::arrange(desc(punct_3min_w))

readr::write_csv(op_summary, "outputs/tables/table1_operator_benchmark.csv")



# 7.2 Method 1 Fixed-effects regression

# Model interpretation is within-operator over time, controlling for quarter shocks
model_df <- panel %>%
  dplyr::filter(!operator %in% c("Great Britain", "England and Wales", "Scotland")) %>%
  dplyr::filter(!is.na(punct_3min), !is.na(cancel_rate))

m1 <- lm(punct_3min ~ cancel_rate + factor(operator) + factor(quarter_start), data = model_df)

# Extract coefficient and basic inference for the main predictor
se <- sqrt(diag(vcov(m1)))[["cancel_rate"]]
est <- coef(m1)[["cancel_rate"]]
tval <- est / se
pval <- 2 * pt(abs(tval), df = df.residual(m1), lower.tail = FALSE)

tab2 <- tibble::tibble(
  Predictor = "Cancellation rate (pp)",
  Estimate  = round(est, 3),
  `Std. Error` = round(se, 3),
  t = round(tval, 2),
  p = signif(pval, 3),
  N = stats::nobs(m1),
  R2 = round(summary(m1)$r.squared, 3),
  Adj_R2 = round(summary(m1)$adj.r.squared, 3)
)

readr::write_csv(tab2, "outputs/tables/table2_regression_cancel_vs_punct.csv")


# 7.3 Method 2 Clustering k-means (k=4)

# Features combine reliability outcomes with cancellation responsibility shares
cause_op <- panel %>%
  dplyr::filter(!operator %in% c("Great Britain", "England and Wales", "Scotland")) %>%
  dplyr::group_by(operator) %>%
  dplyr::summarise(
    denom_canc = sum(canc_total, na.rm = TRUE),
    infra_nm_share = dplyr::if_else(
      denom_canc > 0, sum(canc_infra_nm, na.rm = TRUE) / denom_canc, NA_real_
    ),
    operator_fault_share = dplyr::if_else(
      denom_canc > 0, sum(canc_operator_fault, na.rm = TRUE) / denom_canc, NA_real_
    ),
    .groups = "drop"
  ) %>%
  dplyr::select(-denom_canc)

cluster_features <- op_summary %>%
  dplyr::filter(operator != "Great Britain") %>%
  dplyr::left_join(cause_op, by = "operator") %>%
  dplyr::select(operator, cancel_rate_w, punct_3min_w, infra_nm_share, operator_fault_share) %>%
  dplyr::mutate(dplyr::across(where(is.numeric), ~ dplyr::if_else(is.finite(.x), .x, NA_real_))) %>%
  tidyr::drop_na()

# Prepare numeric matrix for k-means (scaled to equalise units)
Xdf <- cluster_features %>% dplyr::select(-operator)

# Remove any zero-variance columns (scale would create NaN)
sds <- sapply(Xdf, sd, na.rm = TRUE)
Xdf <- Xdf[, sds > 0, drop = FALSE]
X <- scale(Xdf)

set.seed(1) # reproducible clusters
km <- kmeans(X, centers = 4, nstart = 50)

cluster_out <- cluster_features %>%
  dplyr::mutate(
    cluster = paste0("Cluster ", km$cluster),
    infra_nm_share = round(infra_nm_share, 3),
    operator_fault_share = round(operator_fault_share, 3)
  ) %>%
  dplyr::arrange(cluster, desc(punct_3min_w))

readr::write_csv(cluster_out, "outputs/tables/table3_operator_clusters.csv")



# 8 Figures

# Figure 1 GB quarterly cancellation rate
p1 <- panel %>%
  dplyr::filter(operator == "Great Britain") %>%
  ggplot2::ggplot(ggplot2::aes(x = quarter_start, y = cancel_rate)) +
  ggplot2::geom_line() +
  ggplot2::geom_point(size = 1.6) +
  ggplot2::scale_x_date(date_breaks = "6 months", date_labels = "%Y-%m") +
  ggplot2::labs(
    title = "Figure 1. Great Britain: quarterly cancellation rate",
    subtitle = "ORR Table 3123 (Apr 2019–Sep 2025)",
    x = "Quarter", y = "Cancellations (%)"
  ) +
  theme_report() +
  ggplot2::theme(axis.text.x = ggplot2::element_text(angle = 45, hjust = 1))

ggplot2::ggsave("outputs/plots/figure1_gb_cancellation_rate.png", p1, width = 10, height = 5.5, dpi = 300)



# Figure 2 GB quarterly punctuality within 3 minutes

p2 <- panel %>%
  dplyr::filter(operator == "Great Britain") %>%
  ggplot2::ggplot(ggplot2::aes(x = quarter_start, y = punct_3min)) +
  ggplot2::geom_line() +
  ggplot2::geom_point(size = 1.6) +
  ggplot2::scale_x_date(date_breaks = "6 months", date_labels = "%Y-%m") +
  ggplot2::labs(
    title = "Figure 2. Great Britain: punctuality within 3 minutes",
    subtitle = "ORR Table 3133 (Apr 2019–Sep 2025)",
    x = "Quarter", y = "Punctuality within 3 minutes (%)"
  ) +
  theme_report() +
  ggplot2::theme(axis.text.x = ggplot2::element_text(angle = 45, hjust = 1))

ggplot2::ggsave("outputs/plots/figure2_gb_punctuality_3min.png", p2, width = 10, height = 5.5, dpi = 300)



# Figure 3 Operator benchmark scatter
# Label only a small number of extremes to keep the plot readable
op_scatter <- op_summary %>%
  dplyr::filter(operator != "Great Britain") %>%
  dplyr::mutate(
    label = dplyr::if_else(
      dplyr::min_rank(cancel_rate_w) <= 3 | dplyr::min_rank(dplyr::desc(punct_3min_w)) <= 3,
      operator, NA_character_
    )
  )

p3 <- ggplot2::ggplot(op_scatter, ggplot2::aes(x = cancel_rate_w, y = punct_3min_w)) +
  ggplot2::geom_point(alpha = 0.85) +
  ggplot2::geom_text(
    ggplot2::aes(label = label),
    na.rm = TRUE, check_overlap = TRUE, vjust = -0.8, size = 3
  ) +
  ggplot2::labs(
    title = "Figure 3. Operator benchmark: cancellations vs punctuality (2019–2025)",
    subtitle = "Labels show top/bottom performers",
    x = "Weighted cancellation rate (%)",
    y = "Weighted punctuality within 3 minutes (%)"
  ) +
  theme_report()

ggplot2::ggsave("outputs/plots/figure3_operator_scatter.png", p3, width = 10, height = 6, dpi = 300)



# Figure 4 Passenger journeys vs reliability (financial year)
# Aggregate reliability to FY to match passenger journeys table
gb_annual <- panel %>%
  dplyr::filter(operator == "Great Britain") %>%
  dplyr::group_by(fy) %>%
  dplyr::summarise(
    trains_planned = sum(trains_planned, na.rm = TRUE),
    canc_total = sum(canc_total, na.rm = TRUE),
    cancel_rate_w = 100 * canc_total / trains_planned,
    stops = sum(stops, na.rm = TRUE),
    punct_3min_w = sum(stops * punct_3min, na.rm = TRUE) / stops,
    .groups = "drop"
  )

jour_model <- gb_annual %>%
  dplyr::left_join(jour, by = "fy") %>%
  dplyr::filter(!is.na(journeys_m))

# Reshape for faceted plot (one panel per reliability metric)
jour_long <- jour_model %>%
  dplyr::select(fy, journeys_m, cancel_rate_w, punct_3min_w) %>%
  tidyr::pivot_longer(
    cols = c(cancel_rate_w, punct_3min_w),
    names_to = "metric",
    values_to = "value"
  ) %>%
  dplyr::mutate(
    metric = dplyr::recode(
      metric,
      cancel_rate_w = "Weighted cancellation rate (%)",
      punct_3min_w  = "Weighted punctuality within 3 minutes (%)"
    )
  )

p4 <- ggplot2::ggplot(jour_long, ggplot2::aes(x = value, y = journeys_m, label = fy)) +
  ggplot2::geom_point(size = 2) +
  ggplot2::geom_text(check_overlap = TRUE, vjust = -0.8, size = 3) +
  ggplot2::geom_smooth(method = "lm", se = FALSE) +
  ggplot2::facet_wrap(~metric, scales = "free_x") +
  ggplot2::labs(
    title = "Figure 4. Passenger journeys vs reliability (Great Britain, financial year)",
    subtitle = "Journeys: ORR Table 1220; reliability aggregated from Tables 3123 & 3133",
    x = "Reliability metric (annual, weighted)",
    y = "Passenger journeys (million)"
  ) +
  theme_report()

ggplot2::ggsave("outputs/plots/figure4_journeys_vs_reliability.png", p4, width = 10, height = 6, dpi = 300)



# 9 Word-openable results file (.doc as HTML)

# This produces a self-contained summary that opens in Word
# It references images saved in outputs/plots by relative path
html <- c(
  "<html><head><meta charset='utf-8'>",
  "<style>
    body{font-family:Calibri,Arial,sans-serif; font-size:11pt; line-height:1.35;}
    h1{font-size:18pt;} h2{font-size:14pt; margin-top:18px;}
    .cap{color:#444; margin:6px 0 14px 0;}
    img{max-width: 700px; height:auto; border:1px solid #ddd; padding:6px;}
    table.tbl{border-collapse:collapse; margin:10px 0 18px 0; width:100%;}
    table.tbl th, table.tbl td{border:1px solid #bbb; padding:6px; text-align:left; vertical-align:top;}
    table.tbl th{background:#f3f3f3;}
  </style></head><body>",
  "<h1>IJC437 — UK Rail Reliability Results</h1>",
  "<p class='cap'>Datasets: ORR Tables 3123 (cancellations), 3133 (punctuality), 1220 (passenger journeys). Period: Apr 2019–Sep 2025 (quarterly) and Apr 2019–Mar 2025 (financial year).</p>",
  "<h2>Methods included</h2>",
  "<p><b>Prediction</b> Fixed effects regression with operator and quarter fixed effects. <b>Clustering</b> k means segmentation using operator level reliability and responsibility shares.</p>",
  "<h1>Figures</h1>",
  "<h2>Figure 1. Great Britain: quarterly cancellation rate</h2>",
  "<img src='plots/figure1_gb_cancellation_rate.png'>",
  "<h2>Figure 2. Great Britain: punctuality within 3 minutes</h2>",
  "<img src='plots/figure2_gb_punctuality_3min.png'>",
  "<h2>Figure 3. Operator benchmark: cancellations vs punctuality</h2>",
  "<img src='plots/figure3_operator_scatter.png'>",
  "<h2>Figure 4. Passenger journeys vs reliability (financial year)</h2>",
  "<img src='plots/figure4_journeys_vs_reliability.png'>",
  "<h1>Tables</h1>",
  "<h2>Table 1. Operator reliability benchmark (Apr 2019–Sep 2025)</h2>",
  df_to_html_table(
    op_summary %>% dplyr::select(operator, quarters, trains_planned, canc_total, cancel_rate_w, punct_3min_w),
    digits = c(cancel_rate_w = 2, punct_3min_w = 2)
  ),
  "<h2>Table 2. Fixed effects regression: punctuality explained by cancellation rate</h2>",
  df_to_html_table(tab2, digits = c(Estimate = 3, `Std. Error` = 3, t = 2, R2 = 3, Adj_R2 = 3)),
  "<h2>Table 3. Clustering (k means, k = 4): operator segments</h2>",
  df_to_html_table(
    cluster_out %>% dplyr::select(operator, cluster, cancel_rate_w, punct_3min_w, infra_nm_share, operator_fault_share),
    digits = c(cancel_rate_w = 2, punct_3min_w = 2, infra_nm_share = 3, operator_fault_share = 3)
  ),
  "</body></html>"
)

writeLines(html, "outputs/IJC437_Rail_Results.doc")
writeLines(html, "outputs/IJC437_Rail_Results.html")

message("Done. Open: outputs/IJC437_Rail_Results.doc (Word) or outputs/IJC437_Rail_Results.html (browser)")
