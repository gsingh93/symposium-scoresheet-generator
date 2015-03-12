#!/usr/bin/env python

import datetime
import xlrd
import xlsxwriter

def copy_fmt(wb, f, properties={}):
    property_names = [p[4:] for p in dir(f) if p[0:4] == 'set_']
    dft_fmt = wb.add_format()
    new_fmt = wb.add_format({k : v for k, v in f.__dict__.iteritems()
                            if k in property_names and dft_fmt.__dict__[k] != v})

    for key, value in properties.items():
        getattr(new_fmt, 'set_' + key)(value)

    return new_fmt

year = datetime.datetime.now().year
title = str(year) + ' Sikh Youth Symposium'
subtitle = 'By: Sikh Youth Alliance of North America'
region = 'Michigan - Windsor'
local = 'Detroit'

def create_group_worksheet(wb, group, f):
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

        write_row(ws, [(14, title)], f['cb_fmt'], 1)
        write_row(ws, [(14, subtitle)], f['cb_fmt'], 2)

        row_data = [(3, 'Region/Local:', f['cb_fmt']),
                    (3, region + ' / ' + local),
                    (1, ''), (1, '')]
        write_row(ws, row_data, f['center_fmt'], 3, 'E')

        age_str = str(group['age_min']) + ' to ' + str(group['age_max']) + ' yrs'
        row_data = [(1, 'Book:'),
                    (3, group['book_name'], f['center_fmt']),
                    (1, 'Group:'),
                    (1, group['number'], f['center_fmt']),
                    (1, 'Age:'),
                    (2, age_str, f['center_fmt']),
                    (1, 'Judge:'),
                    (2, judge, f['center_fmt']),
                    (2, 'Total Score')]
        write_row(ws, row_data, f['cb_fmt'], 4)

        row_data = [(2, 'Time Allowed:', f['cb_fmt']),
                    (1, str(group['time_limit']) + ' minutes'),
                    (3, 'Questions on Contents'),
                    (1, ''),
                    (5, 'Presentation'),
                    (1, ''), (1, '')]
        write_row(ws, row_data, f['center_fmt'], 5)

        point_categories = [('1', 20), ('2', 20), ('3', 20), ('Overtime', '-'),
                            ('Style &\nDelivery', 8), ('Eye\nContact', 8),
                            ('Voice &\nDiction', 8), ('Language', 8),
                            ('Effectiveness', 8)]
        row_data = [(1, 'No.', f['cb_fmt']), (2, 'Participant Name', f['cb_fmt'])] \
                   + [(1, p[0], f['cb_10_fmt']) for p in point_categories] \
                   + [(1, '100'),
                      (1, 'Rank'),
                      (1, '', f['clear_fmt']),
                      (1, 'Material Total'),
                      (1, 'Material\nRank'),
                      (1, '', f['clear_fmt']),
                      (1, 'Presentation\nTotal'),
                      (1, 'Presentation\nRank')]
        write_row(ws, row_data, f['cb_10_fmt'], 6)

        row_data = [(1, ''), (2, 'Maximum marks:', f['center_fmt'])] \
                   + [(1, p[1], f['cb_10_fmt']) for p in point_categories] \
                   + [(1, ''), (1, '')]
        write_row(ws, row_data, f['cb_10_fmt'], 7)

        start_row = 8
        end_row = start_row + len(group['participants']) - 1
        for i in range(len(group['participants'])):
            participant = group['participants'][i]
            row_num = start_row + i
            ws.set_row(row_num - 1, 30)
            row_data = [(1, i + 1), (2, participant)] \
                       + [(1, '', f['center_fmt'])] * len(point_categories) \
                       + [(1, '=SUM(D{0}:L{0})-2*G{0}'.format(row_num)),
                          (1, '=RANK(M{0},M{1}:M{2},0)'.format(
                              row_num, start_row, end_row)),
                          (1, '', f['clear_fmt']),
                          (1, '=SUM(D{0}:F{0})-G{0}'.format(row_num)),
                          (1, '=RANK(P{0},P{1}:P{2},0)'.format(
                              row_num, start_row, end_row)),
                          (1, '', f['clear_fmt']),
                          (1, '=SUM(H{0}:L{0})'.format(row_num)),
                          (1, '=RANK(S{0},S{1}:S{2},0)'.format(
                              row_num, start_row, end_row))]
            write_row(ws, row_data, f['cb_valign_fmt'], row_num)

def create_final_scoresheet(wb, group, f):
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

    write_row(ws, [(14, title)], f['cb_fmt'], 1)
    write_row(ws, [(14, subtitle)], f['cb_fmt'], 2)

    row_data = [(14, 'Final Rank Sheet')]
    write_row(ws, row_data, f['center_fmt'], 3)

    age_str = str(group['age_min']) + ' to ' + str(group['age_max']) + ' yrs'
    row_data = [(1, 'Book:'),
                (3, group['book_name'], f['center_fmt']),
                (1, 'Group:'),
                (1, group['number'], f['center_fmt']),
                (1, 'Age:'),
                (2, age_str, f['center_fmt']),
                (2, 'Region/Local:', f['cb_fmt']),
                (3, region + ' / ' + local, f['center_fmt'])]
    write_row(ws, row_data, f['cb_fmt'], 4)

    num_judges = len(group['judges'])
    row_data = [(1, 'No.'), (2, 'Participant Name'), (num_judges, 'Ranks given by judges')]
    write_row(ws, row_data, f['cb_fmt'], 5)

    row_data = [(1, ''), (2, 'Judges:', f['cb_fmt'])] \
               + [(1, "='Judge %d'!K4" % (i + 1)) for i in range(num_judges)] \
               + [(1, 'Time used'), (1, 'Punjabi/English'), (1, 'Final\nRank'),
                  (1, 'Final\nPosition'), (1, 'Material\nTie-breaker'),
                  (1, 'Rank')]
    write_row(ws, row_data, f['cb_10_fmt'], 6)

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
                          f['rank_fmt']),
                      (1,  tie_breaker_formula.format(int(row_num) + 1)),
                      (1, '=RANK({3}{0},{3}{1}:{3}{2},1)'.format(
                          row_num, start_row, end_row, rank_col2))]
        write_row(ws, row_data, f['cb_fmt'], row_num)


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

def read_groups(filename):
    groups = []

    workbook = xlrd.open_workbook(filename)
    for worksheet in workbook.sheets():
        info_row = 4
        name_col = 0
        judge_col = 4
        g = {}

        g['number'] = int(worksheet.cell_value(info_row, 0))
        g['age_min'] = int(worksheet.cell_value(info_row, 1))
        g['age_max'] = int(worksheet.cell_value(info_row, 2))
        g['time_limit'] = int(worksheet.cell_value(info_row, 3))
        g['book_name'] = worksheet.cell_value(info_row, 4)

        cur_row = 6
        if worksheet.cell_value(cur_row, name_col) != 'Names':
            raise ValueError('Invalid format')
        if worksheet.cell_value(cur_row, judge_col) != 'Judges':
            raise ValueError('Invalid format')

        participants = []
        cur_row += 1
        while cur_row < worksheet.nrows:
            val = worksheet.cell_value(cur_row, name_col)
            if val == '':
                break
            participants.append(val)
            cur_row += 1

        g['participants'] = participants

        judges = []
        cur_row = 7
        while cur_row < worksheet.nrows:
            val = worksheet.cell_value(cur_row, judge_col)
            if val == '':
                break
            judges.append(val)
            cur_row += 1
        g['judges'] = judges

        groups.append(g)
    return groups

def main():
    groups = read_groups('scoresheet_info_%d.xlsx' % year)
    for group in groups:
        wb = xlsxwriter.Workbook('symposium_%d_group_%d.xlsx'
                                 % (year, group['number']))
        f = {}
        f['clear_fmt'] = wb.add_format()
        f['base_fmt'] = wb.add_format({'border': 1})
        f['center_fmt'] = copy_fmt(wb, f['base_fmt'], {'align': 'center'})
        f['bold_fmt'] = copy_fmt(wb, f['base_fmt'], {'bold': True})
        f['cb_fmt'] = copy_fmt(wb, f['center_fmt'], {'bold': True})
        f['cb_10_fmt'] = copy_fmt(wb, f['cb_fmt'], {'font_size': 10})
        f['cb_valign_fmt'] = copy_fmt(wb, f['cb_fmt'], {'valign': 'vcenter'})
        f['rank_fmt'] = copy_fmt(wb, f['cb_fmt'], {'num_format': 1})

        create_group_worksheet(wb, group, f)
        create_final_scoresheet(wb, group, f)

if __name__ == '__main__':
    main()
