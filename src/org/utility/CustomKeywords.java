package org.utility;

import com.crestech.opkey.plugin.KeywordLibrary;
import com.crestech.opkey.plugin.ResultCodes;
import com.crestech.opkey.plugin.communication.contracts.functionresult.FunctionResult;
import com.crestech.opkey.plugin.communication.contracts.functionresult.Result;
import com.crestech.opkey.plugin.utility.Validations;
import com.crestech.opkey.plugin.utility.exceptionHandler.ArgumentDataMissingException;

public class CustomKeywords implements KeywordLibrary {

	public FunctionResult Custom_ExcelGetRowOfColText(String excelReference, String sheetName, int col, String text) throws ArgumentDataMissingException {
		Validations.checkDataForBlank(0, 1, 2, 3);
		FunctionResult fr = new Utils().setExcelWorkbookAndSheet(excelReference, sheetName);
		if (fr.getOutput().trim().equalsIgnoreCase("false"))
			return Result.FAIL(ResultCodes.ERROR_CONFLICTING_CONFIGURATION).setMessage(fr.getMessage()).setOutput(false).make();
		for (int i = 1; i <= Utils.sh.getLastRowNum(); i++) {
			String cellValue = new Utils().getExcelCellValue(i, col);
			if (cellValue.equals(text))
				return Result.PASS().setOutput(i).make();
		}
		return Result.FAIL(ResultCodes.ERROR_TEXT_NOT_FOUND).setMessage("Text <" + text + "> not found in column " + col).setOutput(false).make();
	}

}
