from openpyxl import styles

# Grey background lines
fill_grey = styles.PatternFill('solid', start_color='FF999999',
                                             end_color='FF999999')
font_bold = styles.Font('Verdana', sz=10, b=True)

# Top row lay-out
font_titles = styles.Font('Verdana', sz=10, b=False)
alignment_titles = styles.Alignment(horizontal='center')
thick_bottom_border = styles.borders.Border(
                    left=styles.borders.Side(style='thin'),
                    right=styles.borders.Side(style='thin'),
                    top=styles.borders.Side(style='thin'),
                    bottom=styles.borders.Side(style='thick'))