#' Convert from labelled data
#'
#' When reading SPSS files with \code{\link{read_data}}, the output contains vectors
#' of class \code{labelled}. This function extracts the labels and attaches them
#' as an attribute (by the same name) to the data, and converts any \code{labelled}
#' variables into factors.
#'
#' @param x A \code{\link[haven]{labelled}} vector or \code{data.frame} containing it.
#' @author Kristian D. Olsen
#' @note \code{\link[data.table]{data.table}} input returns a copy of the \code{data.table}.
#' @export
#' @examples
#' vals <- c("Agree", "Neutral", "Disagree", "Don't know")
#' var <- haven::labelled(c(1, 5), labels = setNames(1:4, vals))
#'
#' # Works with just the labelled vector
#' from_labelled(var)
#'
#' # And a data.frame containing labelled vectors
#' from_labelled(as.data.frame(var))$var


from_labelled <- function(x) UseMethod("from_labelled")

#' @export
from_labelled.default <- function(x) {
  attr(x, "label") <- NULL
  attr(x, "format.stata") <- NULL
  attr(x, "format.spss") <- NULL
  attr(x, "format.sas") <- NULL
  x
}

#' @rdname from_labelled
#' @export
from_labelled.labelled <- function(x) {
  x <- haven::as_factor(x, levels = "default", ordered = FALSE)
  attr(x, "label") <- NULL
  x
}

#' @rdname from_labelled
#' @export
from_labelled.data.frame <- function(x) {
  label <- lapply(x, attr, which = "label", exact = TRUE)
  label <- lapply(label, function(lab) if (is.null(lab)) NA else lab)
  label <- setNames(unlist(label), names(x))

  # Convert, set label attribute and return.
  x[] <- lapply(x, from_labelled)
  structure(x, labels = label)
}

#' @export
from_labelled.data.table <- function(x) {
  x <- from_labelled(as.data.frame(x))
  if (!requireNamespace("data.table", quietly = TRUE)) {
    warning("data.table not installed, returning data.frame.")
    x
  } else {
    data.table::as.data.table(x)
  }
}

#' Convert to labelled
#'
#' Reverses the process from \code{\link{from_labelled}}, by attempting to create
#' labelled variables in place of \code{\link[base]{factor}}, and adding labels
#' to each variable.
#'
#' @param x A \code{factor} or \code{data.frame}.
#' @param label Optional: Set a label when converting a vector to
#' @param ... Ignored.
#' \code{\link[haven]{labelled}}.
#' @author Kristian D. Olsen
#' @note Because of a limitation in \pkg{ReadStat} (it can't write strings longer
#' than 256 characters), \code{\link{write_data}} will write the long strings as
#' a separate .Rdata file. If you use \code{\link{read_data}}, you will get them back.
#' @export
#' @examples
#' levs <- c("Agree", "Neutral", "Disagree", "Don't know")
#' var <- factor(levs[c(1, 3, 4)], levels = levs)
#'
#' # Works with just the factor
#' to_labelled(var)
#'
#' # And a data.frame containing labelled vectors
#' to_labelled(as.data.frame(var))$var

to_labelled <- function(x, label = NULL) UseMethod("to_labelled")

#' @export
to_labelled.default <- function(x, label = NULL) {
  structure(x, label = label %||% attr(x, "label", exact = TRUE))
}

#' @rdname to_labelled
#' @export
to_labelled.factor <- function(x, label = NULL) {
  levels <- levels(x)
  label <- label %||% attr(x, "label", exact = TRUE)
  labels <- setNames(as.integer(1:length(levels)), levels)
  structure(haven::labelled(as.integer(x), labels = labels), label = label)
}

#' @rdname to_labelled
#' @export
to_labelled.data.frame <- function(x, ...) {
  x[] <- lapply(x, to_labelled)

  # Return early if labels are not an attr of the data
  labels <- attr(x, "labels", exact = TRUE)
  if (is.null(labels)) return(x)

  # Assign label to the appropriate variable (against a baseline of NA)
  labels <- c(setNames(rep(NA, length(names(x))), names(x)), labels)
  labels <- labels[!duplicated(names(labels), fromLast = TRUE)]

  # Strip labels from data, and set them as an attribute of the variables.
  x[] <- Map(function(var, lab) {attr(var, "label") <- lab; var}, x, labels[names(x)])
  attr(x, "labels") <- NULL

  x
}

#' @export
to_labelled.data.table <- function(x, ...) {
  df <- to_labelled(as.data.frame(x))
  if (!requireNamespace("data.table", quietly = TRUE)) {
    warning("data.table not installed, returning data.frame.")
    df
  } else {
    data.table::as.data.table(df)
  }
}

