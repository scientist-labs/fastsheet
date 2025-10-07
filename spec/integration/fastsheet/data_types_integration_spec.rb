# frozen_string_literal: true

require 'spec_helper'
require 'fastsheet'

RSpec.describe Fastsheet::Sheet, 'data types', :integration do
  describe 'dates and times' do
    it 'reads date values as Time objects' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Check date rows (rows 1 and 2, 0-indexed)
      date_row_1 = sheet.row(1)
      date_row_2 = sheet.row(2)

      expect(date_row_1[0]).to eq('Date')
      expect(date_row_1[1]).to be_a(Time)
      expect(date_row_1[1].year).to eq(2023)
      expect(date_row_1[1].month).to eq(12)
      expect(date_row_1[1].day).to eq(25)

      expect(date_row_2[1]).to be_a(Time)
      expect(date_row_2[1].year).to eq(2024)
      expect(date_row_2[1].month).to eq(1)
      expect(date_row_2[1].day).to eq(1)
    end

    it 'reads datetime values with time components' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Check datetime rows (rows 3 and 4, 0-indexed)
      datetime_row_1 = sheet.row(3)
      datetime_row_2 = sheet.row(4)

      expect(datetime_row_1[0]).to eq('DateTime')
      expect(datetime_row_1[1]).to be_a(Time)
      expect(datetime_row_1[1].year).to eq(2023)
      expect(datetime_row_1[1].month).to eq(12)
      expect(datetime_row_1[1].day).to eq(25)
      expect(datetime_row_1[1].hour).to eq(14)
      expect(datetime_row_1[1].min).to eq(30)
      expect(datetime_row_1[1].sec).to eq(0)

      expect(datetime_row_2[1]).to be_a(Time)
      expect(datetime_row_2[1].year).to eq(2024)
      expect(datetime_row_2[1].month).to eq(6)
      expect(datetime_row_2[1].day).to eq(15)
      expect(datetime_row_2[1].hour).to eq(9)
      expect(datetime_row_2[1].min).to eq(45)
      expect(datetime_row_2[1].sec).to eq(30)
    end

    it 'handles date edge cases' do
      # Test with a file containing edge case dates
      Tempfile.create(['edge_dates', '.xlsx']) do |tmp|
        tmp.close
        Axlsx::Package.new do |package|
          package.workbook.add_worksheet(name: 'EdgeDates') do |sheet|
            sheet.add_row ['Leap Year', Date.new(2024, 2, 29)] # Leap year date
            sheet.add_row ['Year Start', Date.new(2023, 1, 1)] # Year start
            sheet.add_row ['Year End', Date.new(2023, 12, 31)] # Year end
            sheet.add_row ['DateTime Min', Time.new(2023, 1, 1, 0, 0, 0)] # Midnight
            sheet.add_row ['DateTime Max', Time.new(2023, 12, 31, 23, 59, 59)] # End of day
          end
          package.serialize(tmp.path)
        end

        sheet = described_class.new(tmp.path)

        # Test leap year date
        leap_row = sheet.row(0)
        expect(leap_row[1]).to be_a(Time)
        expect(leap_row[1].month).to eq(2)
        expect(leap_row[1].day).to eq(29)

        # Test midnight and end of day times
        midnight_row = sheet.row(3)
        expect(midnight_row[1].hour).to eq(0)
        expect(midnight_row[1].min).to eq(0)
        expect(midnight_row[1].sec).to eq(0)

        end_of_day_row = sheet.row(4)
        expect(end_of_day_row[1].hour).to eq(23)
        expect(end_of_day_row[1].min).to eq(59)
        expect(end_of_day_row[1].sec).to eq(59)
      end
    end
  end

  describe 'formulas' do
    it 'evaluates basic arithmetic formulas' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Numbers are in rows 5, 6, 7 (10, 20, 5)
      # Basic formulas start at row 8

      # =B6+B7 (10 + 20 = 30) - row 8, column 2 (formula) and potentially calculated result
      addition_row = sheet.row(8)
      expect(addition_row[0]).to eq('Formula')
      expect(addition_row[2]).to eq('=B6+B7')
      # Note: The result might be in a different column depending on how Excel evaluates it

      # =B6*B8 (10 * 5 = 50) - row 9
      multiplication_row = sheet.row(9)
      expect(multiplication_row[2]).to eq('=B6*B8')

      # =B7-B8 (20 - 5 = 15) - row 10
      subtraction_row = sheet.row(10)
      expect(subtraction_row[2]).to eq('=B7-B8')

      # =B8/B8 (5 / 5 = 1) - row 11
      division_row = sheet.row(11)
      expect(division_row[2]).to eq('=B8/B8')
    end

    it 'handles aggregate function formulas' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Aggregate formulas start at row 12
      sum_row = sheet.row(12)
      expect(sum_row[2]).to eq('=SUM(B6:B8)')

      average_row = sheet.row(13)
      expect(average_row[2]).to eq('=AVERAGE(B6:B8)')

      max_row = sheet.row(14)
      expect(max_row[2]).to eq('=MAX(B6:B8)')

      min_row = sheet.row(15)
      expect(min_row[2]).to eq('=MIN(B6:B8)')
    end

    it 'handles text/string formulas' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Text values are in rows 16, 17 ("Hello", "World")
      # Text formulas start at row 18

      concatenate_row = sheet.row(18)
      expect(concatenate_row[2]).to eq('=CONCATENATE(B16," ",B17)')

      length_row = sheet.row(19)
      expect(length_row[2]).to eq('=LEN(B16)')
    end

    it 'handles logical formulas' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Logical formulas start around row 22
      greater_than_row = sheet.row(22)
      expect(greater_than_row[2]).to eq('=B6>B8') # 10 > 5

      less_than_row = sheet.row(23)
      expect(less_than_row[2]).to eq('=B7<B8') # 20 < 5

      if_formula_row = sheet.row(24)
      expect(if_formula_row[2]).to eq('=IF(B6>B8,"Greater","Less")')
    end

    it 'handles date function formulas' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Date function formulas
      today_row = sheet.row(25)
      expect(today_row[2]).to eq('=TODAY()')

      now_row = sheet.row(26)
      expect(now_row[2]).to eq('=NOW()')
    end

    it 'handles formula errors gracefully' do
      # Test with formulas that might cause errors
      Tempfile.create(['formula_errors', '.xlsx']) do |tmp|
        tmp.close
        Axlsx::Package.new do |package|
          package.workbook.add_worksheet(name: 'Errors') do |sheet|
            sheet.add_row ['Error Type', 'Formula']
            sheet.add_row ['Division by Zero', '=1/0']
            sheet.add_row ['Invalid Reference', '=A999+B999']
            sheet.add_row ['Invalid Function', '=INVALIDFUNC()']
            sheet.add_row ['Type Mismatch', '="text"+5']
          end
          package.serialize(tmp.path)
        end

        sheet = described_class.new(tmp.path)

        # These should not crash the library - errors should be handled gracefully
        expect { sheet.rows }.not_to raise_error

        # Error cells should be nil or have some error representation
        error_rows = sheet.rows[1..-1] # Skip header
        error_rows.each do |row|
          expect(row).to be_an(Array)
          expect(row.length).to eq(2)
          # The error values might be nil or some error representation
          # We're mainly testing that it doesn't crash
        end
      end
    end
  end

  describe 'boolean values' do
    it 'reads boolean values correctly' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Boolean values are in rows 20, 21
      true_row = sheet.row(20)
      expect(true_row[0]).to eq('Boolean')
      expect(true_row[1]).to be(true)

      false_row = sheet.row(21)
      expect(false_row[0]).to eq('Boolean')
      expect(false_row[1]).to be(false)
    end
  end

  describe 'mixed data types' do
    it 'handles rows with mixed data types' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Mixed data type rows are at the end
      mixed_row_1 = sheet.row(27) # [nil, nil, 42.5, 'Text']
      expect(mixed_row_1[0]).to eq('Mixed')
      expect(mixed_row_1[1]).to be_nil
      expect(mixed_row_1[2]).to eq(42.5)
      expect(mixed_row_1[3]).to eq('Text')

      mixed_row_2 = sheet.row(28) # [Date, true, nil]
      expect(mixed_row_2[0]).to eq('Mixed')
      expect(mixed_row_2[1]).to be_a(Time) # Date gets converted to Time
      expect(mixed_row_2[1].year).to eq(2024)
      expect(mixed_row_2[1].month).to eq(3)
      expect(mixed_row_2[1].day).to eq(15)
      expect(mixed_row_2[2]).to be(true)
      expect(mixed_row_2[3]).to be_nil
    end

    it 'preserves data types in columns' do
      file = build_temp_xlsx_with_data_types
      sheet = described_class.new(file.path)

      # Test column access with mixed data types
      first_column = sheet.column(0) # Type labels
      expect(first_column).to include('Date', 'DateTime', 'Number', 'Formula', 'Text', 'Boolean', 'Mixed')

      second_column = sheet.column(1) # Values column
      values = second_column.compact # Remove nils

      # Should contain various data types
      expect(values).to include(an_instance_of(Time)) # Dates
      expect(values).to include(an_instance_of(Integer)) # Numbers
      expect(values).to include(an_instance_of(String)) # Text
      expect(values).to include(true, false) # Booleans
    end
  end

  describe 'numeric precision' do
    it 'handles floating point numbers correctly' do
      Tempfile.create(['precision', '.xlsx']) do |tmp|
        tmp.close
        Axlsx::Package.new do |package|
          package.workbook.add_worksheet(name: 'Precision') do |sheet|
            sheet.add_row ['Type', 'Value']
            sheet.add_row ['Float', 3.14159265359]
            sheet.add_row ['Scientific', 1.23e-10]
            sheet.add_row ['Large Float', 123456789.987654321]
            sheet.add_row ['Small Float', 0.000000123]
            sheet.add_row ['Negative', -42.5]
          end
          package.serialize(tmp.path)
        end

        sheet = described_class.new(tmp.path)

        pi_row = sheet.row(1)
        expect(pi_row[1]).to be_a(Float)
        expect(pi_row[1]).to be_within(0.000001).of(3.14159265359)

        scientific_row = sheet.row(2)
        expect(scientific_row[1]).to be_a(Float)
        expect(scientific_row[1]).to be_within(1e-12).of(1.23e-10)

        large_row = sheet.row(3)
        expect(large_row[1]).to be_a(Float)
        expect(large_row[1]).to be > 123456789.0

        negative_row = sheet.row(5)
        expect(negative_row[1]).to eq(-42.5)
      end
    end
  end
end