import sys
import xlrd
import json

#parse file using xlrd
def parse_xlsx(path):
    #open file and read columns from the worksheet row by row
    xlsx_worksheet = xlrd.open_workbook(path)
    the_sheet = xlsx_worksheet.sheet_by_index(0)
    num_rows = the_sheet.nrows
    current_row = 0

    #the controls will be entered here
    controls = {}
    controls['file'] = []

    #only enter data if there is text in the column
    while(current_row < num_rows):

        colA = the_sheet.cell(current_row, 0)
        colB = the_sheet.cell(current_row, 1)
        colC = the_sheet.cell(current_row, 2)

        controls['file'].append({
            'id': colA.value,
            'class': 'SP800-171',
            'title': colB.value,
            'paramaters': [],
            'properties': [
                {
                    'name': "label",
                    'value': colA.value,
                },
                {
                    'name': 'sort-id',
                    'value': colA.value
                }
            ],
            'links': [],
            'parts': [
                {
                    'id': colA.value + '_smt',
                    'name': 'statement',
                    'prose': colB.value,
                    'parts': []
                },
                {
                    'id': colA.value + '_gdn',
                    'name': 'guidance',
                    'prose': colC.value,
                    'links': [
                        {
                            'href': '#ID',
                            'rel': 'reference',
                            'text': colA.value
                        }
                    ]
                }
            ]
        })

        current_row+=1

    #write to the .json file (append mode) for every row in the sheet
    with open('data/parse-output.json', 'a') as outfile:
        json.dump(controls, outfile, indent=4)

#parse file with csv
def parse_csv(path):
    controls = {}
    controls['file'] = []
    controls['file'].append({
        'id': 'ID HERE',
        'class': 'SP800-171',
        'title': 'TITLE PROSE HERE',
        'paramaters': [],
        'properties': [
            {
                'name': "label",
                'value': 'ID HERE',
            },
            {
                'name': 'sort-id',
                'value': 'ID WITH 0 IN FRONT OF EACH NUM'
            }
        ],
        'links': [],
        'parts': [
            {
                'id': 'ID_smt',
                'name': 'statement',
                'prose': 'PROSE HERE (COL B)',
                'parts': []
            },
            {
                'id': 'ID_gdn',
                'name': 'guidance',
                'prose': 'PROSE HERE (COL C)',
                'links': [
                    {
                        'href': '#ID',
                        'rel': 'reference',
                        'text': 'ID'
                    }
                ]
            }
        ]
    })

    with open('data/parse-output.json', 'a') as outfile:
        json.dump(controls, outfile, indent=4)

#get file from command line argument
arg_path = sys.argv[1]

#run the correct function given file extension
if arg_path.endswith('.xlsx'):
    parse_xlsx(arg_path)
elif arg_path.endswith('.csv'):
    parse_csv(arg_path)
else:
    print("File must be of type .xlsx or .csv")