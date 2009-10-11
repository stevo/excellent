

class XlsCellStyle
  attr_reader :raw_style, :style

  ATTRIBUTES_MAP = {
    "font-weight" => :weight,
    "color" => :color,
    "font-size" => :size
  }

  def initialize(style)
    return nil unless style
    @raw_style = style.gsub(' ','')
    @style =  Spreadsheet::Format.new(style2hsh)
  end

  def style2hsh
    attrs = @raw_style.split(';')
    attrs.inject({}){|acc, attr| acc.merge(encode_attribute(attr))}
  end

  def encode_attribute(attribute)
    attr,value = attribute.split(':')

    value = case attr
      when "color" then value.to_sym
      when "font-size" then value.to_i
      when "font-weight" then value.to_sym
    end

    return Hash[attr,value].map_keys(ATTRIBUTES_MAP)
  end

end

class XlsCell

  attr_reader :data

  def initialize(cell,kind='normal')
    if cell.is_a? Hash
      cell.symbolize_keys!
      @data = cell[:content] || ''
      @style = XlsCellStyle.new(cell[:style])
    else
      @data = cell || ''
      @style = nil
    end
  end

  def style
    @style ? @style.style : nil
  end

end

class XlsRow
  attr_reader :cells

  def initialize(row)
    kind = row.keys.include?(:th) ? 'header' : 'normal'
    @cells = row.values.sum.map{|c| XlsCell.new(c,kind)}
  end

end

class XlsTable
  attr_reader :rows

  def initialize(hsh)
    @rows = hsh[:tr].map{|tr| XlsRow.new(tr)}
  end

end

class Hsh2Xls
  require 'spreadsheet'
  require 'xmlsimple'

  attr_reader :xhtml

  REQUIRED_FIRST_LEVEL_KEYS = [:name, :tr]

  def initialize(xhtml)
    @xhtml = xhtml
    @hsh =  XmlSimple.xml_in(@xhtml, {'ForceArray' => true}).symbolize_keys
    validate_hash
  end

  private

  def validate_hash
    #TODO To rewrite
    raise "Table has to have 'name' attribute and a set of rows" if !(REQUIRED_FIRST_LEVEL_KEYS-@hsh.keys ).empty?
    raise "tr is to be array" unless @hsh[:tr].kind_of?(Array)
  end

  def render_xls
    xls = Spreadsheet::Workbook.new
    @sheet = xls.create_worksheet :name => @hsh[:name]

    xls_table = XlsTable.new(@hsh)

    xls_table.rows.each_with_index do |row, row_idx|
      row.cells.each_with_index do |cell, cell_idx|
        render_cell(row_idx, cell_idx, cell)
      end
    end

    xls_output = StringIO.new
    xls.write xls_output
    return xls_output.string
  end

  def render_cell(row_idx, cell_idx, cell)
    @sheet.row(row_idx).set_format(cell_idx, cell.style) if cell.style
    @sheet[row_idx, cell_idx] = cell.data
  end

  public
  
  def output
    render_xls
  end

end