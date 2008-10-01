require 'singleton'

class RowGroupBatch
    
    include Singleton
    
	
    def initialize
        @my_groups = []
    end 
    
    def self.add(*x)
        instance.add(*x)
    end
    
    def self.group_rows(*x)
        instance.group_rows(*x)
    end
    
	def add(start_row_num, end_row_num)
		@my_groups << [start_row_num, end_row_num]
	end
	
    def clear
        @my_groups.clear
    end
	
    def group_pending?
        @my_groups.size > 0
    end
    
	def group_rows(sheet)
		start_row_nums = [:t_int4]
        end_row_nums = [:t_int4]
		
		@my_groups.each{|group|
            start_row_num, end_row_num = group[0], group[1]
			start_row_nums << start_row_num
			end_row_nums << end_row_num
		}

        poi_helper = :PoiHelper.jnew
		poi_helper.groupRows(sheet, start_row_nums, end_row_nums)
        clear
	end
end
