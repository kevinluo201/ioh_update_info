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
  class << self
    attr_accessor :school, :major, :url, :youtube_id
  end

  def self.set_columns(headers)
    @school = headers['學校']
    @major = headers['系全名']
    @url = headers['完整網址']
    @youtube_id = headers['YouTube ID']
  end

  def self.create_new_talk(row)
    create!(url: row[Talk.url],
            school: row[Talk.school].gsub(/\[|\]/, ''),
            major: (row[Talk.major] ? row[Talk.major].gsub(/\[|\]/, '') : ''),
            youtube_url: row[Talk.youtube_id])
    puts "not found, create youtube_url as #{row[youtube_url]}"
  end

  def update_info(row)
    begin
      self.youtube_url = row[Talk.youtube_id]
      save!
      puts "found, update youtube_url as #{youtube_url}"
    rescue
      puts "smthing wrong with #{self.url}"
    end
  end
end
Talk.connection

headers_name = %w(學校 系全名 完整網址 YouTube\ ID)

xlsx = Roo::Spreadsheet.open('./[Editorial][Url][講座網址].xlsx')
sheet = xlsx.sheet('國內')
headers = sheet.row(1)
Talk.set_columns(headers_name.inject({}){ |memo, h| memo[h] = headers.index(h); memo })
#國內
(2..sheet.last_row).each do |row_num|
  row = sheet.row row_num
  if row[12]
    puts 'part2 skipped!'
    puts row[Talk.url]
    next
  end

  talk = Talk.where(url: row[Talk.url]).first
  if talk
    talk.update_info(row)
  else
    Talk.create_new_talk(row)
  end
end

puts '======================================================'
puts '更新國外...............................'
puts '======================================================'
sheet = xlsx.sheet('國外')
headers = sheet.row(1)
Talk.set_columns(headers_name.inject({}){ |memo, h| memo[h] = headers.index(h); memo })
#國外
(2..sheet.last_row).each do |row_num|
  row = sheet.row row_num
  talk = Talk.where(url: row[Talk.url]).first
  if talk
    talk.update_info(row)
  else
    Talk.create_new_talk(row)
  end
end

puts '======================================================'
puts '更新工作...............................'
puts '======================================================'
sheet = '工作創業'
#工作
(2..xlsx.sheet(sheet).last_row).each do |row_num|
  row = xlsx.sheet(sheet).row row_num
  talk = Talk.where(url: row[3]).first
  if talk
    puts "found, update youtube_url as #{row[5]}"
    talk.youtube_url = row[5]
    talk.save!
  else
    puts "not found, create youtube_url as #{row[5]}"
    Talk.create!(url: row[3],
                youtube_url: row[5])
  end
end

puts '======================================================'
puts '更新實習...............................'
puts '======================================================'
sheet = '實習'
#工作
(2..xlsx.sheet(sheet).last_row).each do |row_num|
  row = xlsx.sheet(sheet).row row_num
  talk = Talk.where(url: row[3]).first
  if talk
    puts "found, update youtube_url as #{row[5]}"
    talk.youtube_url = row[5]
    talk.save!
  else
    puts "not found, create youtube_url as #{row[5]}"
    Talk.create!(url: row[3],
                youtube_url: row[5])
  end
end

puts '======================================================'
puts '更新part2...............................'
puts '======================================================'
sheet = '講座 Part 2'
#國外
(2..xlsx.sheet(sheet).last_row).each do |row_num|
  row = xlsx.sheet(sheet).row row_num
  talk = Talk.where(url: row[6]).first
  if talk
    puts "found, update part2 as #{row[8]}"
    talk.part2 = row[8]
    talk.save!
  else
    puts 'not found, we got problem here!'
    puts row[6]
  end
end

puts '======================================================'
puts '更新part3...............................'
puts '======================================================'
sheet = '講座 Part 3'

#國外
(2..xlsx.sheet(sheet).last_row).each do |row_num|
  row = xlsx.sheet(sheet).row row_num
  talk = Talk.where(url: row[6]).first
  if talk
    puts "found, update part3 as #{row[8]}"
    talk.part3 = row[8]
    talk.save!
  elsif row[8].nil?
    puts "not uploaded yet"
  else
    puts 'not found, we got problem here!'
    puts row[6]
  end
end
