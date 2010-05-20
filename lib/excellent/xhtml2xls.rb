class XlsCellStyle
  attr_reader :raw_style, :style

  ATTRIBUTES_MAP = {
          "font-weight" => :weight,
          "color" => :color,
          "font-size" => :size,
          "background-color" => :pattern_fg_color
  }

  def initialize(style)
    return nil unless style
    @raw_style = style.gsub(' ', '')
    hash = style2hsh
    hash[:pattern] = 1 if !hash[:pattern_fg_color].nil?
    @style = Spreadsheet::Format.new(hash)
  end

  def style2hsh
    attrs = @raw_style.split(';')
    attrs.inject({}){|acc, attr| acc.merge(encode_attribute(attr))}
  end

  def encode_attribute(attribute)
    attr, value = attribute.split(':')

    value = case attr
      when "color" then
        value.to_sym
      when "font-size" then
        value.to_i
      when "font-weight" then
        value.to_sym
      when "background-color" then
        value.to_sym
    end

    return Hash[attr, value].map_keys(ATTRIBUTES_MAP)
  end

end

class XlsCell

  attr_reader :data

  def initialize(cell, kind='normal')
    @data = cell.content || ''
    @style = XlsCellStyle.new(cell.attributes["style"].value) if !cell.attributes["style"].nil?
  end

  def style
    @style ? @style.style : nil
  end

end

class XlsRow
  attr_reader :cells

  def initialize(row)
    kind = !row.css('th').empty? ? 'header' : 'normal'
    @cells = [row.css('td').map{|c| XlsCell.new(c, kind)}, row.css('th').map{|c| XlsCell.new(c, kind)}].flatten
  end

end

class XlsTable
  attr_reader :rows

  def initialize(hsh)
    validate_hash(hsh)
    @rows = hsh.css("tr").map{|tr| XlsRow.new(tr)}
  end

  private

  def validate_hash(hsh)
    raise "Table has to have 'name' attribute and a set of rows" if hsh.attributes["name"].nil? or hsh.css("tr").empty?
  end

end

class Hsh2Xls
  require 'spreadsheet'
  require 'nokogiri'

  attr_reader :xhtml, :hsh

  def initialize(xhtml)
    html_doc  = Nokogiri::HTML(xhtml)
    @hsh      = html_doc.css("table");
  end

  private

  def render_xls
    xls = Spreadsheet::Workbook.new

    @hsh.each do |hsh|
      @sheet = xls.create_worksheet :name => hsh.attributes['name'].value
      xls_table = XlsTable.new(hsh)

      xls_table.rows.each_with_index do |row, row_idx|
        row.cells.each_with_index do |cell, cell_idx|
          render_cell(row_idx, cell_idx, cell)
        end
      end
    end

    xls_output = StringIO.new
    xls.write xls_output
    return xls_output.string
  end

  def render_cell(row_idx, cell_idx, cell)
  #TODO for unknown reason last cell in first row cannot been stylee, otherwise whole column is not displayed
  @sheet.row(row_idx).set_format(cell_idx, cell.style) if cell.style and  row_idx != 0  
    if cell.data.empty?
      @sheet[row_idx, cell_idx] = cell.data
    else
      @sheet[row_idx, cell_idx] = cell.data.match(/[\D]/) ? cell.data : cell.data.to_i
    end
  end

  public

  def output
    render_xls
  end

end