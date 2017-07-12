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

xlsx = Roo::Spreadsheet.open('./info_with_youtube_url.xlsx')
sheet = xlsx.sheets.first
#國內
(2..xlsx.sheet(0).last_row).each do |row_num|
  row = xlsx.sheet(sheet).row row_num
  talk = Talk.where(url: row[6]).first
  if talk
    puts "found, update youtube_url as #{row[8]}"
    talk.youtube_url = row[8]
    talk.save!
  else
    puts "not found, create youtube_url as #{row[7]}"
    Talk.create!(url: row[6],
                school: row[0].gsub(/\[|\]/, ''),
                major: row[4].gsub(/\[|\]/, ''),
                youtube_url: row[8])
  end
end
