import pandas as pd
import xlsxwriter
import datetime
import calendar

# Path to your data file
data_file_path = 'S:/FILE_RWID/PyCharmProjects/python-fundamental/Monthly_Air_Quality_Report/DATA PARTIKULAT MATTER (PM2.5) DARI ANALYZER (22).xlsx'

# Load data
data_df = pd.read_excel(data_file_path)

# Ensure 'time' column is datetime format
data_df['time'] = pd.to_datetime(data_df['time'], errors='coerce')

# Sort and clean data
data_df.sort_values(by='time', inplace=True)
data_df.drop_duplicates(subset='time', inplace=True)
data_df = data_df[~data_df['conc'].isin([0, 99999.9])]

# Calculate statistics
data_df['date'] = data_df['time'].dt.date
avg_rh_per_date = data_df.groupby('date')['rh'].mean()
avg_at_per_date = data_df.groupby('date')['at'].mean()
max_conc_per_date = data_df.groupby('date')['conc'].max()
min_conc_per_date = data_df.groupby('date')['conc'].min()
avg_conc_per_date = data_df.groupby('date')['conc'].mean()

# Get month and year from the data for dynamic title and filename
month = data_df['time'].dt.month.mode()[0]
year = data_df['time'].dt.year.mode()[0]
month_name = calendar.month_name[month]
month_names_indonesian = {
    "January": "Januari", "February": "Februari", "March": "Maret",
    "April": "April", "May": "Mei", "June": "Juni",
    "July": "Juli", "August": "Agustus", "September": "September",
    "October": "Oktober", "November": "November", "December": "Desember"
}

# Create dynamic title
dynamic_title = f"Data Perhitungan PM 2.5 Bulan {month_names_indonesian[month_name]} Tahun {year} Stasiun Klimatologi Sumatera Utara"
formatted_title = dynamic_title.replace(" ", "_").replace(".", "")

# Create a new Excel file to save the results
output_file_path = f'S:/FILE_RWID/PyCharmProjects/python-fundamental/Monthly_Air_Quality_Report/{formatted_title}.xlsx'
workbook = xlsxwriter.Workbook(output_file_path)
worksheet = workbook.add_worksheet('Processed Data')

# Define formats for the Excel output
date_format = workbook.add_format({'num_format': 'dd'})
integer_format = workbook.add_format({'num_format': '0'})  # No decimal places for RH
decimal_format = workbook.add_format({'num_format': '0.0'})  # One decimal place for other measurements

# Write headers and data
worksheet.write('A1', dynamic_title)
headers = ['Date', 'Average RH', 'Average AT', 'Max CONC', 'Min CONC', 'Average CONC']
for i, header in enumerate(headers, start=1):
    worksheet.write(1, i, header)

# Fill in the data starting from row 3
row = 2
for date in sorted(avg_rh_per_date.index):
    worksheet.write_datetime(row, 1, datetime.datetime.combine(date, datetime.datetime.min.time()), date_format)
    worksheet.write_number(row, 2, avg_rh_per_date[date], integer_format)  # RH with no decimal places
    worksheet.write_number(row, 3, avg_at_per_date[date], decimal_format)  # One decimal place for others
    worksheet.write_number(row, 4, max_conc_per_date[date], decimal_format)
    worksheet.write_number(row, 5, min_conc_per_date[date], decimal_format)
    worksheet.write_number(row, 6, avg_conc_per_date[date], decimal_format)
    row += 1

# Close the workbook
workbook.close()

print("Data has been processed and saved to:", output_file_path)
