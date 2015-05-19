$:.push File.expand_path("../lib", __FILE__)

# Maintain your gem's version:
require "spreadsheet_importer/version"

# Describe your gem and declare its dependencies:
Gem::Specification.new do |s|
  s.name        = "spreadsheet_importer"
  s.version     = SpreadsheetImporter::VERSION
  s.authors     = ["Nicholas Jakobsen", "Ryan Wallace"]
  s.email       = ["contact@culturecode.ca"]
  s.homepage    = "https://github.com/culturecode/spreadsheet_importer"
  s.summary     = "Makes it easy to import spreadsheets."
  s.description = "Makes it easy to import spreadsheets. Handles .csv and .xlsx formats as well as raw 2D arrays"
  s.license     = "MIT"

  s.files = Dir["{app,config,db,lib}/**/*", "MIT-LICENSE", "Rakefile", "README.md"]
  s.test_files = Dir["test/**/*"]

  s.add_dependency "culturecode-roo", "~> 2.0.2"
  s.add_dependency "charlock_holmes"
end
