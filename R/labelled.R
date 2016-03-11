#' Convert from labelled data
#'
#' When reading SPSS files with \code{\link{read_data}}, the output contains vectors
#' of class \code{labelled}. This function extracts the labels and attaches them
#' as an attribute (by the same name) to the data, and converts any \code{labelled}
#' variables into factors.
#'
#' @param df A data.frame as returned from \code{read_data} or \code{haven::read_sav}.
#' \code{data.frame} is returned from the function.
#' @param copy Return a copy if input is a \code{data.table}.
#' @author Kristian D. Olsen
#' @export
#' @examples
#' df <- haven::read_sav(system.file("extdata", "sample.sav", package = "reporttoolDT"))
#' df <- from_labelled(df)
#' attr(df, "labels")

from_labelled <- function(df, copy = TRUE) UseMethod("from_labelled")

#' @export
from_labelled.data.table <- function(df, copy = TRUE) {
  if (copy) {
    df <- data.table::copy(df)
  }
  from_labelled_impl(df)
}

#' @export
from_labelled.data.frame <- function(df, ...) {
  df <- data.table::as.data.table(df)
  as.data.frame(from_labelled_impl(df))
}

#' @export
from_labelled.tbl_df <- function(df, ...) {
  df <- data.table::as.data.table(df)
  structure(as.data.frame(from_labelled_impl(df)), class = c("tbl_df", "tbl", "data.frame"))
}

#' @export
from_labelled.tbl_dt <- function(df, ...) {
  df <- data.table::as.data.table(df)
  structure(from_labelled_impl(df), class = c("tbl_dt", "tbl", "data.table", "data.frame"))
}

from_labelled_impl <- function(dt) {
  # Preserve labels
  label <- lapply(dt, attr, which = "label")
  label <- unlist(lapply(label, function(x) { if(is.null(x)) NA else x }))

  # Differentiate between scale and factor variables
  labelled <- vapply(dt, inherits, what = "labelled", logical(1))
  labels <- lapply(dt, function(x) if (!is.null(attr(x, "labels"))) names(attr(x, "labels")) else attr(x, "levels"))

  scale_pattern <- get_default("pattern")$detect_scale
  scales <- vapply(labels, function(x) { sum(stri_detect(x, regex = scale_pattern)) == 10L }, logical(1))

  # Check labelled scales for consistency and convert to factor
  cols <- names(dt)
  dt[, cols[scales] := Map(fix_labelled, .SD, names(.SD)), .SDcols = cols[scales], with = FALSE]
  dt[, cols[labelled] := lapply(.SD, haven::as_factor, drop_na = FALSE, ordered = FALSE), .SDcols = cols[labelled], with = FALSE]

  # Return
  data.table::setattr(dt, "labels", label)
  dt
}

# #' Convert to labelled
# #'
# #' Reverses the process from \code{\link{from_labelled}}, by converting a
# #' \code{\link{survey}} back to a format appropriate for writing as a .sav file
# #' using \code{haven}. I.e., it converts factors to \code{labelled} and includes
# #' the label for each variable.
# #'
# #' @param df A data.frame, or \code{Survey}.
# #' @author Kristian D. Olsen
# #' @note Because of a limitation in \code{ReadStat} (it can't write strings longer
# #' than 256 characters), \code{\link{write_data}} will write the long strings as
# #' a separate .Rdata file. If you use \code{\link{read_data}}, you will get them back.
# #' @export
# #' @examples
# #' df <- read_data("test.sav")
# #' fl <- from_labelled(df)
# #' tl <- to_labelled(fl)
#
# # TODO - Update this function.
# to_labelled <- function(survey) {
#
#   # Convert to factors/scales
#   vars <- survey$mm$manifest[survey$mm$type %in% c("scale", "factor")]
#
#   # Make sure all factor/scale variables are factors
#   survey <- factor_data(survey, vars)
#
#   # Convert variables
#   survey$df[] <- lapply(names(survey$df), function(nm, df, mm) {
#
#     x <- df[[nm]]
#
#     # All factors should be 'labelled'
#     if (is.factor(x)) {
#       v <- levels(x)
#       x <- as.numeric(x); x <- haven::labelled(x, setNames(as.numeric(1:length(v)), v), is_na = NULL)
#     } else if (is.character(x)) {
#       # Make sure encoding is native
#       x <- collect_warnings(stri_enc_tonative(x))
#       if (!is.null(x$warnings)) {
#         warnings <- unlist(lapply(x$warnings, "[[", "message"))
#         warning("Warnings when encoding ", nm, " to native:\n",
#                 stri_c(unique(warnings), collapse = "\n"), call. = FALSE)
#       }
#       x <- x$value
#     }
#
#     # Set attributes/class and return
#     attr(x, "label") <- mm$question[mm$manifest %in% nm]
#     x
#
#   }, survey$df, survey$mm)
#
#   # Return
#   survey
#
# }

# Fixes 10-point scales in from_labelled.
# TODO: This is an internal function, and should only be applied if specified.
fix_labelled <- function(x, nm) {

  labels <- attr(x, "labels")
  differ <- setdiff(unique(x[!is.na(x)]), labels)
  nm <- stri_c("In column ", nm, ": ")

  # If it has a 'do not know', fix
  if (length(differ)) {
    if (length(labels) > 10) {
      if (!differ %in% c(11, 98)) {
        warning(nm, "Assigned ", differ, " to label:\n", labels[length(labels)], call. = FALSE)
      }
      # Set last value to 'do not know'
      labels[length(labels)] <- differ

      # Assign the fixed labels
      attr(x, "labels") <- labels

    } else {
      warning(nm, differ, " has been set to NA.", call. = FALSE)
      x[x %in% differ] <- NA
    }
  }

  # Return
  x

}
