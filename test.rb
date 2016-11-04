require "axlsx"

class Excel
  def initialize(message, filename)
    @message = message
    @filename = filename
  end

  def contents
    package = Axlsx::Package.new
    create_workbook(package)
    stream_package(package)
  end

  private

  def create_workbook(package)
    book = package.workbook

    header, body, link = workbook_styles(book)

    book.add_worksheet(name: "Workbook Example") do |sheet|
      sheet.add_row
      sheet.add_row([nil, "Title"], style: header)
      sheet.add_row
      sheet.add_row([nil,  @message], style: body)
      sheet.add_row

      opts = hyperlink_options
      sheet.add_row([nil, opts[:display]], style: link)
      sheet.add_hyperlink(opts)

      sheet.column_widths(15, 100)
    end
  end

  def hyperlink_options
    {
      display: "Click here",
      location: "http://www.google.com",
      ref: "B6"
    }
  end

  def stream_package(package)
    package.serialize(@filename)
    package.to_stream.read
  end

  def workbook_styles(workbook)
    styles = workbook.styles
    header = styles.add_style(sz: 20, b: true)
    body   = styles.add_style(sz: 16, alignment: { wrap_text: true })
    link   = styles.add_style(sz: 14, color: Axlsx::Color.new(rgb: "FF0B0080"), u: true)
    [header, body, link]
  end
end

puts "Non-threaded"
2.times do |n|
  filename = "nonthreaded#{n}.xlsx"
  excel = Excel.new("A message #{n}", filename)
  File.open(filename, "wb") { |f| f.write(excel.contents) }
end

puts "Threaded w/o sleep"
2.times do |n|
  Thread.new do
    filename = "threadwosleep#{n}.xlsx"
    excel = Excel.new("A message #{n}", filename)
    File.open(filename, "wb") { |f| f.write(excel.contents) }
  end
end

puts "Threaded w/ sleep"
2.times do |n|
  Thread.new do
    filename = "threadedsleep#{n}.xlsx"
    excel = Excel.new("A message #{n}", filename)
    File.open(filename, "wb") { |f| f.write(excel.contents) }
  end
  sleep(5)
end
