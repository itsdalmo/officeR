# ReporteRs flextable theme
format_flextable <- function(df, fontsize = 9) {

  # Specify the properties for the table
  tp <- ReporteRs::textProperties(font.size = fontsize)
  bp <- ReporteRs::borderProperties(style="none")

  # Create flextable with proper fontsizes
  ft <- ReporteRs::FlexTable(df, body.text.props = tp, header.text.props = tp)

  # Set columnwidths (assumes question text is found in the second column)
  ft <- ReporteRs::setFlexTableWidths(ft, widths=c(.5, 5, 1.3, 1.3, .5))

  # Add borderproperties
  ft <- ReporteRs::setFlexTableBorders(ft, inner.vertical=bp, inner.horizontal=bp,
                                       outer.vertical=bp, outer.horizontal=bp)

  # Center text in columns with scores
  ft[, 3:ncol(df)] <-  ReporteRs::parProperties(text.align = "center")
  ft[, 3:ncol(df), to = "header"] <-  ReporteRs::parProperties(text.align = "center")

  # Zebrastyle the rows
  ft <- ReporteRs::setZebraStyle(ft, odd = "#DDDDDD", even = "#FFFFFF")

  # Return the flextable
  return(ft)

}