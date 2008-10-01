public class PoiHelper {
	public void setValues(org.apache.poi.hssf.usermodel.HSSFSheet sheet,int[] rowNumbers, 
			int[] colNumbers, int[] cellTypes, int[] dataTypes, String[] cellValues) {
				
			int CELL_TYPE_NUMERIC = 0;
			int CELL_TYPE_STRING = 1;
			int CELL_TYPE_FORMULA = 2;
			
			for (int i = 0; i < colNumbers.length; i++) {
				int rowNumber = rowNumbers[i];
				int colNumber = colNumbers[i];
				int cellType = cellTypes[i];
				int dataType = dataTypes[i];
				String cellValue = cellValues[i];
				org.apache.poi.hssf.usermodel.HSSFRow row = sheet.getRow(rowNumber);
				if (row == null) {
				  row = sheet.createRow(rowNumber);
				}
				org.apache.poi.hssf.usermodel.HSSFCell cell = row.getCell((short)colNumber);
				if (cell == null) {
				  cell = row.createCell((short)colNumber);
				}
				
				if (cellType == CELL_TYPE_FORMULA) {
					cell.setCellFormula(cellValue);
				} else if (cellType == CELL_TYPE_NUMERIC){
					cell.setCellValue(new Double(cellValue).doubleValue());					
				} else {
					cell.setCellValue(cellValue);
				}
				cell.setCellType(cellType);
			}
	}
	
	public void groupRows(org.apache.poi.hssf.usermodel.HSSFSheet sheet, 
									int[] start_row_nums, int[] end_row_nums){
		for (int i = 0; i < start_row_nums.length; i++){
			sheet.groupRow(start_row_nums[i], end_row_nums[i]);
			sheet.setRowGroupCollapsed(start_row_nums[i], true);
		}
	}
}
