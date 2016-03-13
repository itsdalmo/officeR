Workbook <- R6::R6Class("Workbook",
  public = list(
    tables = NULL,
    plots = NULL,
    initialize = function() {
      self$tables = list()
      self$plots = list()
    },
    append = function(x) {
      if (inherits(x, "plot")) {
        self$plots <- c(self$plots, x)
      } else if (is.data.frame(x)) {
        self$tables <- c(self$tables, x)
      }

      invisible(self)
    },

  )
)