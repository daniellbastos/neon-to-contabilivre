import csv
import sys

from datetime import datetime

from openpyxl import load_workbook


def create_contabilivre_file(neon_data, out_filename):
    with open(out_filename, 'w', newline='') as csvfile:
        csv_writer = csv.writer(csvfile, delimiter=';')
        csv_writer.writerow(['data', 'descritivo', 'valor', 'identificador (opcional)'])

        for row in neon_data:
            if not isinstance(row[1].value, datetime):
                continue

            csv_writer.writerow([
                row[1].value.strftime('%d/%m/%Y'),
                row[0].value,
                row[3].value,
                row[5].value,
            ])


print('---- Starting to convert Neon to Contabilivre ----')

in_filename = sys.argv[1]
out_filename = sys.argv[2]

if not out_filename.endswith('.csv'):
    out_filename += '.csv'


wb = load_workbook(filename=in_filename, read_only=True)
ws = wb['Extrato Período']

create_contabilivre_file(ws.rows, out_filename)
print('Done')
