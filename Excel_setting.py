import xlwt


def set_ChartStyle():  # Set the style of top and bottom of chart
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    font.name = '標楷體'
    font.bold = True
    font.height = 220
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    alignment.wrap = 1
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THICK
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THICK
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_ContentStyle():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    font.name = '標楷體'
    font.height = 220
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_ThickLineOnRight():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    font.name = '標楷體'
    font.height = 220
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_ThickLineOnTop():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    font.name = '標楷體'
    font.height = 220
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THIN
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_ThickLineOnTopRight():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    alignment = xlwt.Alignment()
    font.name = '標楷體'
    font.height = 220
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THIN
    style.font = font
    style.alignment = alignment
    style.borders = borders
    return style


def set_GreenBackground():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    alignment = xlwt.Alignment()
    font.name = '標楷體'
    font.bold = True
    font.height = 220
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    alignment.wrap = 1
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THICK
    borders.right = xlwt.Borders.THICK
    borders.top = xlwt.Borders.THICK
    borders.bottom = xlwt.Borders.THICK
    pattern.pattern_fore_colour = 42
    style.font = font
    style.alignment = alignment
    style.borders = borders
    style.pattern = pattern
    return style
