require 'active_record'
require 'roo'

# Change the following to reflect your database settings
ActiveRecord::Base.establish_connection(
  adapter:  'mysql2', # or 'postgresql' or 'sqlite3' or 'oracle_enhanced'
  host:     '130.211.248.191',
  database: 'ioh_talks_info',
  username: 'ioh_talk',
  password:
)

# Define your classes based on the database, as always
class Talk < ActiveRecord::Base
  def update_info(row, headers_with_index)
    attrs = {}
    headers_with_index.each do |header, index|
      col = case header
            when '學校' then 'school'
            when '學位' then 'degree'
            when '領域名', '領域名(待更新)', '網站上領域名' then 'major'
            when '講座語言' then 'language'
            when '講者國籍' then 'nationality'
            when '所有台灣求學講座', '所有海外求學講座', '講座類別' then 'talk_class'
            when '國家', '講座國家' then 'country'
            when '國家 + 學位' then 'country_degree'
            when '公司' then 'company'
            when '職稱' then 'title'
            end

      attrs[col] = row[index].gsub(/\[|\]/,'') if row[index]
    end
    self.assign_attributes attrs
  end
end
Talk.connection

xlsx = Roo::Spreadsheet.open('./[Editorial][Url][講座網址].xlsx')
puts '======================================================'
puts '更新國內...............................'
puts '======================================================'
sheet = xlsx.sheet('國內')

headers = sheet.row(1)
headers_name = %w(學校 領域名\(待更新\) 講座語言 講者國籍 所有台灣求學講座 學位)
headers_with_index = headers_name.inject({}){ |memo, h| memo[h] = headers.index(h); memo }
#國內
(2..sheet.last_row).each do |row_num|
  row = sheet.row row_num
  next unless row[headers.index('完整網址')] #未上架跳過

  talk = Talk.find_or_initialize_by(url: row[headers.index('完整網址')])
  talk.update_info(row, headers_with_index)
  begin
    talk.save!
    puts "Talk #{row_num} updated"
  rescue
    puts "Talk #{row} update failed!!!!!"
  end
end

puts '======================================================'
puts '更新國外...............................'
puts '======================================================'
sheet = xlsx.sheet('國外')

headers = sheet.row(1)
headers_name = %w(學校 學位 講座國家 國家\ +\ 學位 網站上領域名 所有海外求學講座)
headers_with_index = headers_name.inject({}){ |memo, h| memo[h] = headers.index(h); memo }
#國內
(2..sheet.last_row).each do |row_num|
  row = sheet.row row_num
  next unless row[headers.index('完整網址')] #未上架跳過

  talk = Talk.find_or_initialize_by(url: row[headers.index('完整網址')])
  talk.update_info(row, headers_with_index)
  begin
    talk.save!
    puts "Talk #{row_num} updated"
  rescue
    puts "Talk #{row} update failed!!!!!"
  end
end

puts '======================================================'
puts '更新工作/創業/實習講座...............................'
puts '======================================================'
sheet = xlsx.sheet('工作創業實習')

headers = sheet.row(1)
headers_name = %w(公司 職稱 講座國家 講座類別)
headers_with_index = headers_name.inject({}){ |memo, h| memo[h] = headers.index(h); memo }
puts headers_with_index
#國內
(2..sheet.last_row).each do |row_num|
  row = sheet.row row_num
  next unless row[headers.index('完整網址')] #未上架跳過

  talk = Talk.find_or_initialize_by(url: row[headers.index('完整網址')])
  talk.update_info(row, headers_with_index)
  begin
    talk.save!
    puts "Talk #{row_num} updated"
  rescue
    puts "Talk #{row} update failed!!!!!"
  end
end
