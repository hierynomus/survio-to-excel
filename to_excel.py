import json
import xlsxwriter
from collections import defaultdict


def is_int(v):
    try:
        int(v)
    except:
        return False
    return True


def parse_respondents_to_excel(json_data, output_file):
    respondent_dict = defaultdict(list)
    nr_respondents = len(json_data)
    if nr_respondents == 0:
        print("No respondents found in file")
        exit(1)

    for respoondent in json_data:
        for k, v in respoondent.iteritems():
            if k not in ["use", "datetime", "email-id"]:
                respondent_dict[k].append(v)

    for q, answers in respondent_dict.iteritems():
        if len(answers) != nr_respondents:
            print("question %s does not have %d answers, but %d" % (q, nr_respondents, len(answers)))

    workbook = xlsxwriter.Workbook(output_file)
    bold = workbook.add_format({'bold': True})
    sheet = workbook.add_worksheet("Sheet 1")
    row = 0
    for q in sorted(respondent_dict):
        sheet.write(row, 0, q, bold)
        col = 1
        for answer in respondent_dict[q]:
            if is_int(answer):
                answer = int(answer)
            sheet.write(row, col, answer)
            col += 1
        row += 1

    workbook.close()
