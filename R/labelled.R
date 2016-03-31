#' Convert from labelled data
#'
#' When reading SPSS files with \code{\link{read_data}}, the output contains vectors
#' of class \code{labelled}. This function extracts the labels and attaches them
#' as an attribute (by the same name) to the data, and converts any \code{labelled}
#' variables into factors.
#'
#' @param df A data.frame as returned from \code{read_data} or \code{haven::read_sav}.
#' @author Kristian D. Olsen
#' @note \code{data.table} input will return a \code{data.frame}.
#' @export
#' @examples
#' # TODO

from_labelled <- function(df) UseMethod("from_labelled")

#' @rdname from_labelled
#' @export
from_labelled.data.frame <- function(df, ...) {
  # Store variable label
  label <- lapply(df, attr, which = "label")
  label <- unlist(lapply(label, function(x) { if(is.null(x)) NA else x }))

  # Convert all labelled variables to factor
  is_labelled <- vapply(df, inherits, what = "labelled", logical(1))
  if (any(is_labelled)) {
    df[is_labelled] <- lapply(df[is_labelled], haven::as_factor, drop_na = FALSE, ordered = FALSE)
  }

  # Strip labels from variables, and set them as an attribute of data.
  df[] <- lapply(df, strip_label)
  attr(df, "labels") <- label
  df
}

#' @rdname from_labelled
#' @export
from_labelled.data.table <- function(df) {
  from_labelled(as.data.frame(df))
}

#' Convert to labelled
#'
#' Reverses the process from \code{\link{from_labelled}}, by attempting to create
#' labelled variables in place of \code{factor}, and adding labels to each variable.
#'
#' @param df A data.frame, or \code{Survey}.
#' @author Kristian D. Olsen
#' @note Because of a limitation in \code{ReadStat} (it can't write strings longer
#' than 256 characters), \code{\link{write_data}} will write the long strings as
#' a separate .Rdata file. If you use \code{\link{read_data}}, you will get them back.
#' @export
#' @examples
#' # TODO

to_labelled <- function(x) UseMethod("to_labelled")

#' @rdname to_labelled
#' @export
to_labelled.data.frame <- function(x) {
  labels <- attr(x, "labels")
  if (is.null(labels)) {
    # Assign label to the appropriate variable (against a baseline of NA)
    labels <- c(setNames(rep(NA, length(names(x))), names(x)), labels)
    labels <- labels[!duplicated(labels, fromLast = FALSE)]
  } else {
    # If labels are not an attribute of the data, check variables.
    # (e.g., if you have just read in data and not converted from_labelled yet.)
    labels <- lapply(df, attr, which = "label")
    labels <- unlist(lapply(labels, function(x) { if(is.null(x)) NA else x }))
    labels <- setNames(labels, names(x))
  }

  # Convert all factors to labelled.
  is_factor <- vapply(x, is.factor, logical(1))
  if (any(is_factor)) {
    x[is_factor] <- lapply(x[is_factor], as_labelled)
  }

  # Strip labels from data, and set them as an attribute of the variables.
  x[] <- Map(function(v, l) {attr(v, "label") <- l; v}, x, labels[names(x)])
  attr(x, "labels") <- NULL
  x

}

#' @rdname to_labelled
#' @export
to_labelled.data.table <- function(x) {
  to_labelled(as.data.frame(x))
}

# Convert factors to a (integer based) labelled variable -----------------------
as_labelled <- function(x) {
  stopifnot(is.factor(x))
  levels <- levels(x)
  labels <- setNames(as.integer(1:length(levels)), levels)
  haven::labelled(as.integer(x), labels = labels, is_na = NULL)
}

# Strip the label attribute from a variable (after reading data with haven) ----
strip_label <- function(x) {
  stopifnot(is.atomic(x))
  attr(x, "label") <- NULL
  x
}
