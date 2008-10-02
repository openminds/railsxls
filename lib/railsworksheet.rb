require 'yajb/jbridge'
require 'yajb/jlambda'
require 'tempfile'

PATH_SEPARATOR =  RUBY_PLATFORM.include?("win32") ? ';' : ':' 
    
#It would be more readable to keep the jar file name with full version details
#So just look for the pattern in the jar file name ignoring the version details

POI_JAR = Dir[File.join(File.dirname(__FILE__), 'poi*jar')][0] 

JBRIDGE_OPTIONS = {
   :jvm_stdout => :never,
   :classpath =>  POI_JAR + PATH_SEPARATOR + File.dirname(__FILE__) 
}

include JavaBridge

module Worksheet
  class Handler < ActionView::TemplateHandler
	    def compile(template)
    		%Q{_set_controller_content_type(Mime::XLS)
			   controller.headers["Content-Disposition"] ||= 'attachment; filename="' + (@worksheet_name || "\#{(params[:id] || params[:action])}.xls") + '"'
			   
         Worksheet::Base.new do |workbook|
           #{template.source}
         end.process
         }
    end    
  end
  
  class Base
    @@temp_file = nil
    
    def temp_file_path
      if @@temp_file.nil?
          temp = Tempfile.new('railsworksheet-', File.join(RAILS_ROOT, 'tmp') )
          @@temp_file =  temp.path 
          temp.close
      end
      @@temp_file
    end
    
    def initialize
      jimport "java.io.*"
      jimport "org.apache.poi.hssf.usermodel.*"
      
      CellBatch.instance.clear
      RowGroupBatch.instance.clear
      yield workbook
      CellBatch.write_to(workbook.getSheetAt(0)) if CellBatch.instance.write_pending?
      RowGroupBatch.group_rows(workbook.getSheetAt(0)) if RowGroupBatch.instance.group_pending?
    end
    
    def workbook
		  @workbook ||= jnew :HSSFWorkbook
    end
    
    def process
      out = jnew :FileOutputStream, temp_file_path
      workbook.write(out)
      out.close
      
      File.open(temp_file_path, 'rb') { |file| file.read }         
    end
    
  end
end
