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
        'judges': ['Dilpreet Singh', 'Daljeet Singh'],
        'participants': participants
    },
]

# TODO: Make this not global
wb = xlsxwriter.Workbook('symposium2014.xlsx')
clear_fmt = wb.add_format()
base_fmt = wb.add_format({'border': 1})
center_fmt = copy_fmt(wb, base_fmt, {'align': 'center'})
bold_fmt = copy_fmt(wb, base_fmt, {'bold': True})
cb_fmt = copy_fmt(wb, center_fmt, {'bold': True})
cb_10_fmt = copy_fmt(wb, cb_fmt, {'font_size': 10})
cb_valign_fmt = copy_fmt(wb, cb_fmt, {'valign': 'vcenter'})
rank_fmt = copy_fmt(wb, cb_fmt, {'num_format': 1})

def create_group_worksheet(wb, group):
    for judge_num in range(len(group['judges'])):
        judge = group['judges'][judge_num]
        ws = wb.add_worksheet('Judge ' + str(judge_num + 1))
        column_sizes = [('B:C', 11),
                        ('G', 11),
                        ('K', 12),
                        ('L', 12),
                        ('P', 11),
                        ('S:T', 11),
                        ('O', 0.3),
                        ('R', 0.3)]
        set_column_sizes(ws, column_sizes)
        ws.set_row(5, 25)

        write_row(ws, [(14, title)], cb_fmt, 1)
        write_row(ws, [(14, subtitle)], cb_fmt, 2)

        row_data = [(3, 'Region/Local:', cb_fmt),
                    (3, region + ' / ' + local),
                    (1, ''), (1, '')]
        write_row(ws, row_data, center_fmt, 3, 'E')

        age_str = str(group['age_min']) + ' to ' + str(group['age_max']) + ' yrs'
        row_data = [(1, 'Book:'),
                    (3, group['book_name'], center_fmt),
                    (1, 'Group:'),
                    (1, group['number'], center_fmt),
                    (1, 'Age:'),
                    (2, age_str, center_fmt),
                    (1, 'Judge:'),
                    (2, judge, center_fmt),
                    (2, 'Total Score')]
        write_row(ws, row_data, cb_fmt, 4)

        row_data = [(2, 'Time Allowed:', cb_fmt),
                    (1, str(group['time_limit']) + ' minutes'),
                    (3, 'Questions on Contents'),
                    (1, ''),
                    (5, 'Presentation'),
                    (1, ''), (1, '')]
        write_row(ws, row_data, center_fmt, 5)

        point_categories = [('1', 20), ('2', 20), ('3', 20), ('Overtime', '-'),
                            ('Style &\nDelivery', 8), ('Eye\nContact', 8),
                            ('Voice &\nDiction', 8), ('Language', 8),
                            ('Effectiveness', 8)]
        row_data = [(1, 'No.', cb_fmt), (2, 'Participant Name', cb_fmt)] \
                   + [(1, p[0], cb_10_fmt) for p in point_categories] \
                   + [(1, '100'),
                      (1, 'Rank'),
                      (1, '', clear_fmt),
                      (1, 'Material Total'),
                      (1, 'Material\nRank'),
                      (1, '', clear_fmt),
                      (1, 'Presentation\nTotal'),
                      (1, 'Presentation\nRank')]
        write_row(ws, row_data, cb_10_fmt, 6)

        row_data = [(1, ''), (2, 'Maximum marks:', center_fmt)] \
                   + [(1, p[1], cb_10_fmt) for p in point_categories] \
                   + [(1, ''), (1, '')]
        write_row(ws, row_data, cb_10_fmt, 7)

        start_row = 8
        end_row = start_row + len(group['participants']) - 1
        for i in range(len(group['participants'])):
            participant = group['participants'][i]
            row_num = start_row + i
            ws.set_row(row_num - 1, 30)
            row_data = [(1, i + 1), (2, participant)] \
                       + [(1, '', center_fmt)] * len(point_categories) \
                       + [(1, '=SUM(D{0}:L{0})-2*G{0}'.format(row_num)),
                          (1, '=RANK(M{0},M{1}:M{2},0)'.format(
                              row_num, start_row, end_row)),
                          (1, '', clear_fmt),
                          (1, '=SUM(D{0}:F{0})-G{0}'.format(row_num)),
                          (1, '=RANK(P{0},P{1}:P{2},0)'.format(
                              row_num, start_row, end_row)),
                          (1, '', clear_fmt),
                          (1, '=SUM(H{0}:L{0})'.format(row_num)),
                          (1, '=RANK(S{0},S{1}:S{2},0)'.format(
                              row_num, start_row, end_row))]
            write_row(ws, row_data, cb_valign_fmt, row_num)

def create_final_scoresheet(wb, group):
    ws = wb.add_worksheet('Final Scores')
    column_sizes = [('B:C', 11),
                    ('D:F', 12),
                    ('G', 11),
                    ('H', 14),
                    ('K', 12),
                    ('L', 12),
                    ('P', 11),
                    ('S:T', 11),
                    ('O', 0.3),
                    ('R', 0.3)]
    set_column_sizes(ws, column_sizes)
    ws.set_row(5, 30)

    write_row(ws, [(14, title)], cb_fmt, 1)
    write_row(ws, [(14, subtitle)], cb_fmt, 2)

    row_data = [(14, 'Final Rank Sheet')]
    write_row(ws, row_data, center_fmt, 3)

    age_str = str(group['age_min']) + ' to ' + str(group['age_max']) + ' yrs'
    row_data = [(1, 'Book:'),
                (3, group['book_name'], center_fmt),
                (1, 'Group:'),
                (1, group['number'], center_fmt),
                (1, 'Age:'),
                (2, age_str, center_fmt),
                (2, 'Region/Local:', cb_fmt),
                (3, region + ' / ' + local, center_fmt)]
    write_row(ws, row_data, cb_fmt, 4)

    num_judges = len(group['judges'])
    row_data = [(1, 'No.'), (2, 'Participant Name'), (num_judges, 'Ranks given by judges')]
    write_row(ws, row_data, cb_fmt, 5)

    row_data = [(1, ''), (2, 'Judges:', cb_fmt)] \
               + [(1, "='Judge %d'!K4" % (i + 1)) for i in range(num_judges)] \
               + [(1, 'Time used'), (1, 'Punjabi/English'), (1, 'Final\nRank'),
                  (1, 'Final\nPosition'), (1, 'Material\nTie-breaker'),
                  (1, 'Rank')]
    write_row(ws, row_data, cb_10_fmt, 6)

    start_row = 7
    end_row = start_row + len(group['participants']) - 1
    for i in range(len(group['participants'])):
        name = group['participants'][i]
        row_num = start_row + i
        ws.set_row(row_num - 1, 30)
        end_col = inc_col('D', num_judges - 1)
        tie_breaker_formula = '='
        for j in range(num_judges):
            tie_breaker_formula += "'Judge %d'!P{0}" % (j + 1)
            if j != num_judges - 1:
                tie_breaker_formula += ' + '

        rank_col1 = inc_col('F', num_judges)
        rank_col2 = inc_col('G', num_judges)
        rank_col3 = inc_col('H', num_judges)
        row_data = [(1, i + 1), (2, name)] \
                   + [(1, "='Judge %d'!N%d" % (i + 1, row_num + 1))
                      for i in range(num_judges)] \
                   + [(1, ''), (1, ''),
                      (1, '=SUM(D{0}:{1}{0})'.format(row_num, end_col)),
                      (1, '=RANK({3}{0},{3}{1}:{3}{2},1) - 0.0001 * {4}{0}'.format(
                          row_num, start_row, end_row, rank_col1, rank_col3),
                          rank_fmt),
                      (1,  tie_breaker_formula.format(int(row_num) + 1)),
                      (1, '=RANK({3}{0},{3}{1}:{3}{2},1)'.format(
                          row_num, start_row, end_row, rank_col2))]
        write_row(ws, row_data, cb_fmt, row_num)


def set_column_sizes(ws, sizes):
    for s in sizes:
        if ':' in s[0]:
            ws.set_column(s[0], s[1])
        else:
            assert len(s[0]) == 1
            ws.set_column('{0}:{0}'.format(s[0]), s[1])

def inc_col(col, n):
    return chr(ord(col) + n)

def write_row(ws, data, fmt, row, start_col='A'):
    end_col = start_col
    for d in data:
        assert d[0] > 0
        end_col = inc_col(start_col, d[0] - 1)
        format = d[2] if len(d) == 3 else fmt

        if start_col == end_col:
            ws.write(start_col + str(row), d[1], format)
        else:
            cells = '{1}{0}:{2}{0}'.format(row, start_col, end_col)
            ws.merge_range(cells, d[1], format)
        start_col = inc_col(end_col, 1)

def main():
    # ws = wb.add_worksheet('test')
    # fmt = wb.add_format()
    # fmt.set_num_format('0.00')
    # ws.write(0, 0, .9999, fmt)
    # ws.write(1, 0, '=A1')
    for group in groups:
        create_group_worksheet(wb, group)
        create_final_scoresheet(wb, group)

if __name__ == '__main__':
    main()
    wb.close()
