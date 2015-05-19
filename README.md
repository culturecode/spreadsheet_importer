# SpreadsheetImporter

```ruby
gem 'spreadsheet_importer'
```


## Usage
```ruby
  SpreadsheetImporter::Import.from_spreadsheet(spreadsheet) do |row|
    # Do some work on the row
  end

  # Starting at a custom offset
  SpreadsheetImporter::Import.from_spreadsheet(spreadsheet, :start_row => 5) do |row|
```

### 2D Array
```ruby
  SpreadsheetImporter::Import.from_spreadsheet([["Bob", "Hoskins"], ["Roger", "Rabbit"]])
```

### CSV
```ruby
  SpreadsheetImporter::Import.from_csv("users.csv")
```

### .xlsx
```
  SpreadsheetImporter::Import.from_xlsx("users.xlsx")
  SpreadsheetImporter::Import.from_xlsx("users.xlsx", :sheet_name => '2015') # Processing a single sheet
```
