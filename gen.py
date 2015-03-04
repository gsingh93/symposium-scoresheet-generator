#!/usr/bin/env python

import xlsxwriter

def copy_fmt(wb, f, properties={}):
    new_fmt = wb.add_format()

    new_fmt.num_format = 0
    new_fmt.num_format_index = 0
    new_fmt.font_index = 0
    new_fmt.has_font = 0
    new_fmt.has_dxf_font = 0

    new_fmt.bold = f.bold
    new_fmt.underline = f.underline
    new_fmt.italic = f.italic
    new_fmt.font_name = f.font_name
    new_fmt.font_size = f.font_size
    new_fmt.font_color = 0x0
    new_fmt.font_strikeout = 0
    new_fmt.font_outline = 0
    new_fmt.font_shadow = 0
    new_fmt.font_script = 0
    new_fmt.font_family = 2
    new_fmt.font_charset = 0
    new_fmt.font_scheme = 'minor'
    new_fmt.font_condense = 0
    new_fmt.font_extend = 0
    new_fmt.theme = 0
    new_fmt.hyperlink = 0

    new_fmt.hidden = 0
    new_fmt.locked = 1

    new_fmt.text_h_align = f.text_h_align
    new_fmt.text_wrap = 0
    new_fmt.text_v_align = f.text_v_align
    new_fmt.text_justlast = 0
    new_fmt.rotation = 0
    new_fmt.center_across = 0

    new_fmt.fg_color = f.fg_color
    new_fmt.bg_color = f.bg_color
    new_fmt.pattern = 0
    new_fmt.has_fill = 0
    new_fmt.has_dxf_fill = 0
    new_fmt.fill_index = 0
    new_fmt.fill_count = 0

    new_fmt.border_index = f.border_index
    new_fmt.has_border = f.has_border
    new_fmt.has_dxf_border = f.has_dxf_border
    new_fmt.border_count = f.border_count

    new_fmt.bottom = f.bottom
    new_fmt.bottom_color = f.bottom_color
    new_fmt.diag_border = f.diag_border
    new_fmt.diag_color = f.diag_color
    new_fmt.diag_type = f.diag_type
    new_fmt.left = f.left
    new_fmt.left_color = f.left_color
    new_fmt.right = f.right
    new_fmt.right_color = f.right_color
    new_fmt.top = f.top
    new_fmt.top_color = f.top_color

    new_fmt.indent = 0
    new_fmt.shrink = 0
    new_fmt.merge_range = 0
    new_fmt.reading_order = 0
    new_fmt.just_distrib = 0
    new_fmt.color_indexed = 0
    new_fmt.font_only = 0

    for key, value in properties.items():
        getattr(new_fmt, 'set_' + key)(value)

    return new_fmt

title = '2014 Sikh Youth Symposium'
subtitle = 'By: Sikh Youth Alliance of North America'
region = 'Michigan - Windsor'
local = 'Detroit'

participants = [
    'Ravleen Kaur',
    'Gaurik Singh',
    'Gurnoor Kaur',
    'Harjot Singh',
    'Nishan Singh',
    'Tajvir Singh',
    'Tegbir Singh',
]

groups = [
    {
        'number': 2,
        'age_min': 9,
        'age_max': 10,
        'time_limit': 6,
        'book_name': 'Selected Episodes from Sikh History',
        'judges': ['Dilpreet Singh', 'Daljeet Singh', 'Gurmeet Singh'],
        'participants': participants
    },
]

# TODO: Make this not global
wb = xlsxwriter.Workbook('symposium2014.xlsx')
base_fmt = wb.add_format({'border': 1})
center_fmt = copy_fmt(wb, base_fmt, {'align': 'center'})
bold_fmt = copy_fmt(wb, base_fmt, {'bold': True})
cb_fmt = copy_fmt(wb, center_fmt, {'bold': True})
cb_10_fmt = copy_fmt(wb, cb_fmt, {'font_size': 10})

def create_group_worksheet(wb, group):
        ws = wb.add_worksheet()
        ws.set_column('B:C', 11)
        ws.set_column('G:G', 11)
        ws.set_column('K:K', 12)
        ws.set_column('L:L', 12)
        ws.set_row(5, 25)

        write_row(ws, [(19, title)], cb_fmt, 1)
        write_row(ws, [(19, subtitle)], cb_fmt, 2)
#        worksheet.merge_range('A1:S1', title, cb_fmt)
        #worksheet.merge_range('A2:S2', , cb_fmt)

        ws.merge_range('H3:I3', 'Region/Local:', cb_fmt)
        ws.merge_range('J3:L3', region + ' / ' + local, center_fmt)
        ws.write('M3', '', center_fmt)
        ws.write('N3', '', center_fmt)

        ws.write('A4', 'Book:', cb_fmt)
        ws.merge_range('B4:D4', group['book_name'], center_fmt)
        ws.write('E4', 'Group:', cb_fmt)
        ws.write('F4', group['number'], center_fmt)
        ws.write('G4', 'Age:', cb_fmt)
        age_str = str(group['age_min']) + ' to ' + str(group['age_max']) + ' yrs'
        ws.merge_range('H4:I4', age_str, center_fmt)
        ws.write('J4', 'Judge:', cb_fmt)
        ws.merge_range('K4:L4', group['judges'][0], center_fmt) # TODO
        ws.merge_range('M4:N4', 'Total Score', cb_fmt)

        ws.merge_range('A5:B5', 'Time Allowed:', cb_fmt)
        ws.write('C5', str(group['time_limit']) + ' minutes', center_fmt)
        ws.merge_range('D5:F5', 'Questions on Contents', center_fmt)
        ws.merge_range('H5:L5', 'Presentation', center_fmt)
        ws.write('M5', '', center_fmt)
        ws.write('N5', '', center_fmt)

        ws.write('A6', 'No.', cb_fmt)
        ws.merge_range('B6:C6', 'Participant Name', cb_fmt)

        point_categories = [('1', 20), ('2', 20), ('3', 20), ('Overtime', '-'), ('Style &\nDelivery', 8),
                            ('Eye\nContact', 8), ('Voice &\nDiction', 8),
                            ('Language', 8), ('Effectiveness', 8)]

        ws.merge_range('B7:C7', 'Maximum marks:', center_fmt)
        for i in range(len(point_categories)):
            col = chr(ord('D') + i)
            ws.write(col + '6', point_categories[i][0], cb_10_fmt)
            ws.write(col + '7', point_categories[i][1], cb_10_fmt)
        ws.write('M6', 100, cb_10_fmt)
        ws.write('N6', 'Rank', cb_10_fmt)
        ws.write('M7', '', cb_10_fmt)
        ws.write('N7', '', cb_10_fmt)

        cb_valign_fmt = copy_fmt(wb, cb_fmt, {'valign': 'vcenter'})
        for i in range(len(group['participants'])):
            participant = group['participants'][i]
            row_num = str(8 + i)
            ws.set_row(int(row_num) - 1, 30)
            ws.write('A' + row_num, i + 1, cb_fmt)
            ws.merge_range('B' + row_num + ':' + 'C' + row_num, participant, cb_valign_fmt)
            for i in range(len(point_categories)):
                col = chr(ord('D') + i)
                ws.write(col + row_num, '', center_fmt)
            col = chr(ord('D') + len(point_categories))
            ws.write(col + row_num, '=SUM(D{0}:L{0})-2*G{0}'.format(row_num), cb_fmt)
            col = chr(ord(col) + 1)
            ws.write(col + row_num, '=RANK(M{0},M{0}:M{0},0)'.format(row_num), cb_fmt)

def write_row(ws, data, fmt, row, start_col='A'):
    def inc_col(col, n):
        return chr(ord(col) + n)

    end_col = start_col
    for d in data:
        assert d[0] > 0
        end_col = inc_col(start_col, d[0] - 1)
        cells = '{1}{0}:{2}{0}'.format(row, start_col, end_col)
        format = d[2] if len(d) == 3 else fmt
        ws.merge_range(cells, d[1], format)

def main():
    for group in groups:
        create_group_worksheet(wb, group)

if __name__ == '__main__':
    main()
    wb.close()
