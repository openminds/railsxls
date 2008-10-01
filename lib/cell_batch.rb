require 'singleton'  
 
class CellBatch
    
    include Singleton
    
    CELL_TYPE_NUMERIC = 0
    CELL_TYPE_STRING = 1
    CELL_TYPE_FORMULA = 2
    
    def initialize
        @my_cells = []
    end
    
    def self.add(*x)
        instance.add(*x)
    end
    
    def self.write_to(*x)
        instance.write_to(*x)
    end
    
    
    def add(row_num, col_num, value, formula=false)
        @my_cells << [row_num, col_num, value, formula]
    end
    
    def clear
        @my_cells.clear
    end       
    
    def write_pending?
        @my_cells.size > 0
    end
        
    def write_to(sheet)
            row_nums = [:t_int4]
            col_nums = [:t_int4]
            cell_types = [:t_int4]
            data_types = [:t_int4]
            cell_values = [:t_string]
            @my_cells.each{|cell_data|
                row_num, col_num, value, formula = cell_data[0], cell_data[1], cell_data[2], cell_data[3] 

                row_nums << row_num
                col_nums << col_num
				if value.respond_to?('abs')
					cell_types << CELL_TYPE_NUMERIC
				elsif formula
					cell_types << CELL_TYPE_FORMULA
				else
					cell_types << CELL_TYPE_STRING
				end
                data_types << (value.respond_to?('abs') ? 1 : 2)
                cell_values << value.to_s
            }
			
            cell_update_java_object =  :PoiHelper.jnew
            cell_update_java_object.setValues(sheet, row_nums, col_nums, cell_types, data_types, cell_values)
            clear
    end
    
end

