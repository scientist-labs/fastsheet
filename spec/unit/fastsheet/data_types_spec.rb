# frozen_string_literal: true

require 'spec_helper'
require 'fastsheet/sheet'

RSpec.describe Fastsheet::Sheet, 'data types' do
  describe 'data type conversion' do
    context 'with mocked data containing various types' do
      let(:mixed_rows) do
        [
          ['Header1', 'Header2', 'Header3', 'Header4'],
          ['2023-12-25T00:00:00', 42, 'Hello World', true],
          ['2024-06-15T14:30:45', 3.14159, 'Test String', false],
          [Time.new(2024, 1, 1, 12, 0, 0), -17, '', nil]
        ]
      end

      before do
        allow(described_class).to receive(:new) do |*args|
          instance = described_class.allocate
          allow(instance).to receive(:read!) do |_file_name = nil, _sheet_selector = nil|
            instance.instance_variable_set(:@rows, mixed_rows.map(&:dup))
            instance.instance_variable_set(:@height, mixed_rows.length)
            instance.instance_variable_set(:@width, mixed_rows.first.length)
            instance.instance_variable_set(:@sheet_name, 'Sheet1')
            instance.instance_variable_set(:@sheet_index, 0)
            nil
          end
          instance.send(:initialize, *args)
          instance
        end
      end

      it 'handles mixed data types in rows' do
        sheet = described_class.new('dummy.xlsx')

        header_row = sheet.row(0)
        expect(header_row).to eq(['Header1', 'Header2', 'Header3', 'Header4'])

        data_row_1 = sheet.row(1)
        expect(data_row_1[0]).to be_a(String) # Date as string in mock
        expect(data_row_1[1]).to eq(42)
        expect(data_row_1[2]).to eq('Hello World')
        expect(data_row_1[3]).to be(true)

        data_row_2 = sheet.row(2)
        expect(data_row_2[0]).to be_a(String) # DateTime as string in mock
        expect(data_row_2[1]).to eq(3.14159)
        expect(data_row_2[2]).to eq('Test String')
        expect(data_row_2[3]).to be(false)
      end

      it 'handles nil values properly' do
        sheet = described_class.new('dummy.xlsx')

        row_with_nil = sheet.row(3)
        expect(row_with_nil[3]).to be_nil
      end

      it 'handles empty strings' do
        sheet = described_class.new('dummy.xlsx')

        row_with_empty = sheet.row(3)
        expect(row_with_empty[2]).to eq('')
      end

      it 'preserves data types in columns' do
        sheet = described_class.new('dummy.xlsx')

        # Test different columns with different data types
        first_column = sheet.column(0) # Dates/Time
        expect(first_column.length).to eq(4)
        expect(first_column).to include('Header1')

        second_column = sheet.column(1) # Numbers
        expect(second_column).to include('Header2', 42, 3.14159, -17)

        third_column = sheet.column(2) # Strings
        expect(third_column).to include('Header3', 'Hello World', 'Test String', '')

        fourth_column = sheet.column(3) # Booleans and nil
        expect(fourth_column).to include('Header4', true, false, nil)
      end
    end
  end

  describe 'edge cases' do
    context 'with edge case data' do
      let(:edge_case_rows) do
        [
          [0, '', nil, false],
          [0.0, ' ', true, true],
          [-0.0, '  whitespace  ', false, nil],
          [Float::INFINITY, "\t\n", nil, false]
        ]
      end

      before do
        allow(described_class).to receive(:new) do |*args|
          instance = described_class.allocate
          allow(instance).to receive(:read!) do |_file_name = nil, _sheet_selector = nil|
            instance.instance_variable_set(:@rows, edge_case_rows.map(&:dup))
            instance.instance_variable_set(:@height, edge_case_rows.length)
            instance.instance_variable_set(:@width, edge_case_rows.first.length)
            instance.instance_variable_set(:@sheet_name, 'Sheet1')
            instance.instance_variable_set(:@sheet_index, 0)
            nil
          end
          instance.send(:initialize, *args)
          instance
        end
      end

      it 'handles zero values correctly' do
        sheet = described_class.new('dummy.xlsx')

        first_row = sheet.row(0)
        expect(first_row[0]).to eq(0)

        second_row = sheet.row(1)
        expect(second_row[0]).to eq(0.0)

        third_row = sheet.row(2)
        expect(third_row[0]).to eq(-0.0)
      end

      it 'handles special float values' do
        sheet = described_class.new('dummy.xlsx')

        infinity_row = sheet.row(3)
        expect(infinity_row[0]).to eq(Float::INFINITY)
      end

      it 'handles various string edge cases' do
        sheet = described_class.new('dummy.xlsx')

        rows = sheet.rows

        # Empty string
        expect(rows[0][1]).to eq('')

        # Single space
        expect(rows[1][1]).to eq(' ')

        # String with surrounding whitespace
        expect(rows[2][1]).to eq('  whitespace  ')

        # Whitespace characters
        expect(rows[3][1]).to eq("\t\n")
      end

      it 'handles boolean edge cases' do
        sheet = described_class.new('dummy.xlsx')

        rows = sheet.rows

        # Various boolean combinations
        expect(rows[0][2]).to be_nil
        expect(rows[0][3]).to be(false)

        expect(rows[1][2]).to be(true)
        expect(rows[1][3]).to be(true)

        expect(rows[2][2]).to be(false)
        expect(rows[2][3]).to be_nil
      end
    end
  end

  describe 'large datasets' do
    it 'handles sheets with many rows and columns efficiently' do
      # Create a large dataset for testing
      large_rows = []
      header = (1..100).map { |i| "Col#{i}" }
      large_rows << header

      # Generate 1000 rows of mixed data
      1000.times do |row_idx|
        row_data = []
        100.times do |col_idx|
          case col_idx % 5
          when 0
            row_data << row_idx + col_idx # Integer
          when 1
            row_data << (row_idx + col_idx) * 0.1 # Float
          when 2
            row_data << "Row#{row_idx}_Col#{col_idx}" # String
          when 3
            row_data << (row_idx + col_idx).even? # Boolean
          when 4
            row_data << nil # Nil
          end
        end
        large_rows << row_data
      end

      allow(described_class).to receive(:new) do |*args|
        instance = described_class.allocate
        allow(instance).to receive(:read!) do |_file_name = nil, _sheet_selector = nil|
          instance.instance_variable_set(:@rows, large_rows)
          instance.instance_variable_set(:@height, large_rows.length)
          instance.instance_variable_set(:@width, large_rows.first.length)
          instance.instance_variable_set(:@sheet_name, 'Sheet1')
          instance.instance_variable_set(:@sheet_index, 0)
          nil
        end
        instance.send(:initialize, *args)
        instance
      end

      sheet = described_class.new('dummy.xlsx')

      expect(sheet.height).to eq(1001) # 1000 rows + header
      expect(sheet.width).to eq(100)

      # Test that we can access various parts of the large dataset
      expect(sheet.row(0)).to eq(header)
      expect(sheet.row(500)[0]).to eq(499) # row_idx starts from 0, so row 500 has idx 499
      expect(sheet.row(1000)[50]).to eq(1049) # row_idx=999, col_idx=50, case 0: 999+50=1049

      # Test column access
      first_col = sheet.column(0)
      expect(first_col[0]).to eq('Col1') # Header
      expect(first_col[1]).to eq(0) # First data row
      expect(first_col[1001]).to be_nil # Beyond data

      # Test enumeration works with large dataset
      row_count = 0
      sheet.each_row { |_row| row_count += 1 }
      expect(row_count).to eq(1001)

      column_count = 0
      sheet.each_column { |_col| column_count += 1 }
      expect(column_count).to eq(100)
    end
  end
end