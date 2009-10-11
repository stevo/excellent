#Changes keys of hash accordig to given map
# >> {:a => 10, :c => 20}.map_keys({:a => :b})
# => {:c=>20, :b=>10}
class Hash
  def map_keys(hsh)
    returning Hash.new do |r|
      each{|k,v| r[hsh[k] || k] = v}
    end
  end
end