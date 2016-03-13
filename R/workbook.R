to_excel

# Workbook (R6 Class) ----------------------------------------------------------
#' @importFrom R6 R6Class
Workbook <- R6::R6Class("Workbook",
  public = list(
    objects = NULL,

    initialize = function() {
      self$tables = list()
      self$plots = list()
    },

    append = function(x) {
      if (!is.data.frame(x) || !inherits(x, "plot")) {
        stop("Input must either be a data.frame or a plot.")
      }
      self$objects <- c(self$objects, x)
      self
    },

    ppt_output = function() {

    },

    excel_output = function() {

    },

    print = function() {
      cat("Workbook\n")
      cat("Tables: ", length(self$tables))
      cat("Plots: ", length(self$plots))
    }

  )
)
