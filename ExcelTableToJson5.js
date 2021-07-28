function ExcelTableToJSON(rngTable) {
	
	var log = ""
	log = log + "1"
	
    try {
        if (rngTable && rngTable['Rows'] && rngTable['Columns']) {
            var rowCount = rngTable.Rows.Count;
            var columnCount = rngTable.Columns.Count;
            var arr = new Array();
			var headers = new Array();
			
			for (columnLoop = 1; columnLoop <= columnCount; columnLoop++){
				headers[columnLoop - 1] = rngTable.Cells(1, columnLoop);
			}
			
            for (rowLoop = 2; rowLoop <= rowCount; rowLoop++) {
                var dict = {};
                for (columnLoop = 1; columnLoop <= columnCount; columnLoop++) {
                    var rngCell = rngTable.Cells(rowLoop, columnLoop);
					if(rngCell.Value2  != "" && rngCell.Value2  != "NULL")
					{
						dict[headers[columnLoop - 1]] = rngCell.Value2;
					}
                }
				arr[rowLoop - 2] = dict;
            }
			
			return JSON.stringify(arr);
        }
        else {
            return { error: '#Either rngTable is null or does not support Rows or Columns property!' };
        }
    }
    catch(err) {
        return {error: err.message};
    }
}