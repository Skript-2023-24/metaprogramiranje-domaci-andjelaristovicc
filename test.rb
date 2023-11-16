require "google_drive"


#BITNO 
#STA JE YIELD 
#it's used inside methods for calling a block. 
#In other words, yield is pausing our method & transfering control to the block 
#so that it can do its thing & then come back with a return value.

#EACH
#method allows one loop through the elements of an array to perform operations on them.

#MAP
#The method** map** returns a new array with the results of running a block once for every element in enum.

#klasa za rad sa google sheetom
class GoogleSheet
attr_reader :worksheet

      def initialize(session, worksheet_title)
          #Inicijalizuje GS objekat sa sesijom i imenom
          @worksheet = session.spreadsheet_by_title(worksheet_title).worksheets[0]
      end
     #preuzima sve vrednosti sa worksheeta u niz
     def table_values
      (1..@worksheet.num_rows).map do |row|
        (1..@worksheet.num_cols).map { |col| @worksheet[row, col] }
      end
    end

      # preuzima određeni red na osnovu datog indeksa
      def row_by_index(index)
        (1..@worksheet.num_cols).map { |col| @worksheet[index, col] }
      end
      # Omogućava iteriranje kroz svaku ćeliju na radnom listu
      include Enumerable

      def each(&block)
        (1..@worksheet.num_rows).each do |row|
          (1..@worksheet.num_cols).each { |col| block.call(@worksheet[row, col]) }
        end
      end
    
      def merged_cell?(row, col)
        @worksheet.merged_ranges.any? { |merged_range| merged_range.include?(row, col) }
      end

end

# Klasa tabele za obavljanje operacija na GoogleSheet-u
class Table
  def initialize(worksheet)
    @worksheet = worksheet
  end

  #Biblioteka prepoznaje ukoliko postoji na bilo koji način ključna reč total ili subtotal 
  #unutar sheet-a, i ignoriše taj red

  # Preuzima vrednosti iz određene kolone, isključujući redove koji sadrže ključne reči „ukupno“ ili „subtotal“
  def column_values(column_name)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)

    return nil unless col_index

    @worksheet.rows.drop(1)
               .reject { |row| row[col_index]&.downcase&.match?(/total|subtotal/) }
               .map { |row| row[col_index] }
               .compact
  end

  # Pristupa vrednostima na određenom indeksu unutar kolone
  def value_at(column_name, index)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)

    return nil unless col_index

    row = @worksheet.rows[index + 1] # preskoci header
    row[col_index] if row
  end

  # Postavlja novu vrednost na određenom indeksu u koloni
 def set_value_at(column_name, index, value)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)

    return puts "Column '#{column_name}' not found." unless col_index

    cell = @worksheet[index + 1, col_index + 1] 

    return puts "Cell is empty or doesn't exist." unless cell && !cell.empty?

    cell.value = value
    @worksheet.save 
  end


  # Metode za direktan pristup celim kolonama koristeći imena metoda kao što su prvaKolona, drugaKolona, itd.
  # (mogu se dodati slične metode za druge kolone)
  def prvaKolona
    column_values('Prva Kolona')
  end

  def drugaKolona
    column_values('Druga Kolona')
  end

  def trecaKolona
    column_values('Treca Kolona')
  end

  # Izračunava zbir numeričkih vrednosti u navedenoj koloni
  def sum(column_name)
    column = column_values(column_name)
    # filtrira nenumeričke vrednosti pre izračunavanja zbira
    numeric_values = column.compact.map(&:to_i).select { |value| value.to_s == column_name.to_s }
    numeric_values.sum
  end

   # Izračunava avg numeričkih vrednosti u navedenoj koloni
  def avg(column_name)
    column = column_values(column_name)
    # filtrira nenumeričke vrednosti pre izračunavanja avg
    numeric_values = column.compact.map(&:to_f).select { |value| value.to_s == column_name.to_s }
    #ako postoje brojevi sum deli sa count
    if numeric_values.any?
      sum = numeric_values.sum
      count = numeric_values.size
      sum / count
    else
      nil #vratimo ako nema ni jedan broj
    end
  end

  # vrati red na osnovu vrednosti ćelije u određenoj koloni
  def row_by_cell_value(column_name, value)
    header_row = @worksheet.rows.first
    col_index = header_row.index(column_name)

    if col_index
      row = @worksheet.rows.find { |r| r[col_index] == value }
      row
    else
      nil
    end
  end
  # Izvodi operaciju mape na koloni sa numeričkim vrednostima
  def map_column(column_name, &block)
    column = column_values(column_name)
    # filtrira nenumeričke vrednosti pre izračunavanja
    numeric_values = column.compact.map(&:to_i).select { |value| value.to_s == column_name.to_s }
    numeric_values.map(&block) if numeric_values.any?
  end

  # Izvodi operaciju izbora na koloni
  def select_column(column_name, &block)
    column = column_values(column_name)
    column.select(&block) if column
  end

  # Izvodi operaciju smanjenja na koloni
  def reduce_column(column_name, initial, &block)
    column = column_values(column_name)
    column.compact.reduce(initial, &block) if column
  end

end

def main
 session = GoogleDrive::Session.from_config("config.json")
 ws = session.spreadsheet_by_key("1WtFH5iecKHTIQgWS2EXZ7pR9XM7mWOHFUE37NQJ-eSk").worksheets[0]
 gs = GoogleSheet.new(session, 'Rubyproject')
 table = Table.new(gs.worksheet)

  puts "Entire column 'Prva kolona':"
  p table.column_values('Prva kolona') #pristupamo celoj koloni

  puts "Accessing value at index 1 in 'Prva kolona':"
  p table.value_at('Prva kolona', 1) #pristupamo vrednosti u koloni na indexu 1


  puts 'Table values:'
  p gs.table_values

  puts 'Row by index:'
  p gs.row_by_index(3)

  puts 'All cells in the worksheet:'
  gs.each { |cell| puts cell }

  #pristupamo koristeci ime kolona
  prva_kolona_values = table.prvaKolona
  druga_kolona_values = table.drugaKolona
  treca_kolona_values = table.trecaKolona

puts "Values in 'Prva kolona': #{prva_kolona_values}"
puts "Values in 'Druga kolona': #{druga_kolona_values}"
puts "Values in 'Treca kolona': #{treca_kolona_values}"

# pristupamo odredjenim vrednostima
value_in_druga_kolona = table.value_at('Druga kolona', 0) # pristup prvoj vrednosti u "Druga kolona"
p value_in_druga_kolona # Output: 25

#Postavljamo novu vrednost na indeksu
table.set_value_at('Prva kolona', 1, 8) # Postavljanje '8' na indeksu 1 u "Prva kolona"

updated_prva_kolona = table.prvaKolona
p updated_prva_kolona 

sum_prva_kolona = table.sum('Prva kolona')
avg_druga_kolona = table.avg('Druga kolona')

puts "Sum of 'Prva kolona': #{sum_prva_kolona}"
puts "Average of 'Druga kolona': #{avg_druga_kolona}"

specific_row = table.row_by_cell_value('Prva kolona', '1')
puts "Row with 'Prva kolona' value '1': #{specific_row}"


mapped_treca_kolona = table.map_column('Treca kolona') { |value| value * 2 }
selected_prva_kolona = table.select_column('Prva kolona') { |value| value.to_i > 2 }
reduced_druga_kolona = table.reduce_column('Druga kolona', 0) { |sum, value| sum + value.to_i }

puts "Mapped values in 'Treca kolona': #{mapped_treca_kolona}"
puts "Selected values in 'Prva kolona' greater than 2: #{selected_prva_kolona}"
puts "Reduced value in 'Druga kolona': #{reduced_druga_kolona}"

end

main if __FILE__ == $PROGRAM_NAME