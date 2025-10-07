# frozen_string_literal: true

require 'caxlsx'
require 'tempfile'

module XlsxHelpers
  module_function

  # Build a temporary XLSX file and return the Tempfile instance.
  # Caller is responsible for unlinking if needed; Tempfile will clean up at GC.
  def build_temp_xlsx(rows: [[1, 2, 3], [4, 5, 6], [7, 8, 9]], header: nil, sheet_name: 'Sheet1')
    Tempfile.create(['fastsheet_spec', '.xlsx']).tap do |tmp|
      tmp.close
      Axlsx::Package.new do |package|
        package.workbook.add_worksheet(name: sheet_name) do |sheet|
          sheet.add_row(header) if header
          rows.each { |row_values| sheet.add_row(row_values) }
        end
        package.serialize(tmp.path)
      end
    end
  end

  # Build a temporary XLSX file with multiple sheets
  def build_temp_xlsx_multi_sheet
    Tempfile.create(['fastsheet_multi_spec', '.xlsx']).tap do |tmp|
      tmp.close
      create_multi_sheet_workbook(tmp.path)
    end
  end

  def create_multi_sheet_workbook(file_path)
    Axlsx::Package.new do |package|
      add_basic_sheet(package.workbook)
      add_data_sheet(package.workbook)
      add_numbers_sheet(package.workbook)
      package.serialize(file_path)
    end
  end

  def add_basic_sheet(workbook)
    workbook.add_worksheet(name: 'Sheet1') do |sheet|
      sheet.add_row %w[A1 B1]
      sheet.add_row %w[A2 B2]
    end
  end

  def add_data_sheet(workbook)
    workbook.add_worksheet(name: 'Data') do |sheet|
      sheet.add_row %w[Name Age City]
      sheet.add_row ['Alice', 30, 'NYC']
      sheet.add_row ['Bob', 25, 'LA']
    end
  end

  def add_numbers_sheet(workbook)
    workbook.add_worksheet(name: 'Numbers') do |sheet|
      sheet.add_row [1, 2, 3, 4]
      sheet.add_row [5, 6, 7, 8]
      sheet.add_row [9, 10, 11, 12]
    end
  end

  # Build a temporary XLSX file with dates, times, and formulas
  def build_temp_xlsx_with_data_types
    Tempfile.create(['fastsheet_data_types', '.xlsx']).tap do |tmp|
      tmp.close
      Axlsx::Package.new do |package|
        package.workbook.add_worksheet(name: 'DataTypes') do |sheet|
          # Header row
          sheet.add_row ['Type', 'Value', 'Formula', 'Result']

          # Date values
          sheet.add_row ['Date', Date.new(2023, 12, 25), nil, nil]
          sheet.add_row ['Date', Date.new(2024, 1, 1), nil, nil]

          # Time values (DateTime)
          sheet.add_row ['DateTime', Time.new(2023, 12, 25, 14, 30, 0), nil, nil]
          sheet.add_row ['DateTime', Time.new(2024, 6, 15, 9, 45, 30), nil, nil]

          # Numbers for formulas
          sheet.add_row ['Number', 10, nil, nil]
          sheet.add_row ['Number', 20, nil, nil]
          sheet.add_row ['Number', 5, nil, nil]

          # Basic formulas
          sheet.add_row ['Formula', nil, '=B6+B7', nil] # Sum: 10 + 20 = 30
          sheet.add_row ['Formula', nil, '=B6*B8', nil] # Product: 10 * 5 = 50
          sheet.add_row ['Formula', nil, '=B7-B8', nil] # Difference: 20 - 5 = 15
          sheet.add_row ['Formula', nil, '=B8/B8', nil] # Division: 5 / 5 = 1

          # More complex formulas
          sheet.add_row ['Formula', nil, '=SUM(B6:B8)', nil] # Sum range: 10+20+5 = 35
          sheet.add_row ['Formula', nil, '=AVERAGE(B6:B8)', nil] # Average: (10+20+5)/3 = 11.667
          sheet.add_row ['Formula', nil, '=MAX(B6:B8)', nil] # Max: 20
          sheet.add_row ['Formula', nil, '=MIN(B6:B8)', nil] # Min: 5

          # String/text formulas
          sheet.add_row ['Text', 'Hello', nil, nil]
          sheet.add_row ['Text', 'World', nil, nil]
          sheet.add_row ['Formula', nil, '=CONCATENATE(B16," ",B17)', nil] # "Hello World"
          sheet.add_row ['Formula', nil, '=LEN(B16)', nil] # Length of "Hello" = 5

          # Boolean values
          sheet.add_row ['Boolean', true, nil, nil]
          sheet.add_row ['Boolean', false, nil, nil]

          # Logical formulas
          sheet.add_row ['Formula', nil, '=B6>B8', nil] # 10 > 5 = TRUE
          sheet.add_row ['Formula', nil, '=B7<B8', nil] # 20 < 5 = FALSE
          sheet.add_row ['Formula', nil, '=IF(B6>B8,"Greater","Less")', nil] # "Greater"

          # Date formulas
          sheet.add_row ['Formula', nil, '=TODAY()', nil] # Current date
          sheet.add_row ['Formula', nil, '=NOW()', nil] # Current date and time

          # Mixed data types
          sheet.add_row ['Mixed', nil, 42.5, 'Text']
          sheet.add_row ['Mixed', Date.new(2024, 3, 15), true, nil]
        end

        package.serialize(tmp.path)
      end
    end
  end
end

RSpec.configure do |config|
  config.include XlsxHelpers
end
