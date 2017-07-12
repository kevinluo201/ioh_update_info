require 'active_record'
require 'roo'

# Change the following to reflect your database settings
ActiveRecord::Base.establish_connection(
  adapter:  'mysql2', # or 'postgresql' or 'sqlite3' or 'oracle_enhanced'
  host:
  database:
  username:
  password:
)

# Define your classes based on the database, as always
class Talk < ActiveRecord::Base

end

xlsx = Roo::Spreadsheet.open('./info.xlsx')
xlsx.sheets.each do |sheet|
  (2..xlsx.sheet(sheet).last_row).each do |row_num|
    row = xlsx.sheet(sheet).row row_num
    unless Talk.where(url: row[-1]).first
      Talk.create(url: row[-1],
                  school: row[0].gsub(/\[|\]/, ''),
                  major: row[4].gsub(/\[|\]/, ''))
    end
  end
end
