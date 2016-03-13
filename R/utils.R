# Capture dots. Primarily used to pass calls to R6 methods ---------------------
capture_dots <- function(...) {
  eval(substitute(alist(...)))
}

# Get renamed columns from lazy_dots in dplyr::select and dplyr::rename.
# Return is equivalent to: setNames(old_name, new_name) ------------------------
renamed_vars <- function(dots) {
  expr <- dots[!is.na(names(dots)) & names(dots) != ""]
  if (length(expr)) {
    nms <- lapply(names(expr), function(nm) { x <- expr[[nm]]$expr; if (x != nm) x })
    nms <- setNames(as.character(unlist(nms)), names(expr))
  } else {
    nms <- NULL
  }
  nms
}

# Check which OS we are on -----------------------------------------------------
on_windows <- function() {
  Sys.info()["sysname"] == "Windows"
}

on_osx <- function() {
  Sys.info()["sysname"] == "Darwin"
}

# match_all returns ALL matching indices for x in table, ordered by x ----------
match_all <- function(x, table) {
  unlist(lapply(x, function(x) which(table == x)))
}

# normalizes paths and removes trailing /'s ------------------------------------
clean_path <- function(path) {
  if (!is.string(path)) stop("Path must be a string.")

  # Normalize if path is not absolute.
  if (!stri_detect(path, regex = "^(/|[A-Za-z]:|\\\\|~)")) {
    path <- normalizePath(path, winslash = "/", mustWork = FALSE)
  }

  # Remove trailing slashes
  if (stri_detect(path, regex = "/$")) {
    path <- stri_replace(path, "", regex = "/$")
  }

  path

}

# Retrieve the filename sans extension -----------------------------------------
filename_no_ext <- function(file)  {
  stri_replace(basename(file), "$1", regex = stri_c("(.*)\\.", tools::file_ext(file), "$"))
}

# Check if x is a string (length 1 character vector.) --------------------------
is.string <- function(x) {
  is.character(x) && length(x) == 1
}

# See if a list or data.frame contains any labelled vectors. -------------------
any_labelled <- function(x) {
  any(vapply(x, inherits, what = "labelled", logical(1)))
}

# Like is.list, except it does not return true for data.frame ------------------
is.list2 <- function(x) {
  inherits(x, "list")
}

# Hadley's %||% ----------------------------------------------------------------
`%||%` <- function(a, b) if (!is.null(a)) a else b

