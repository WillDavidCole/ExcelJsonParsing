function ExcelTableToJSON(rngTable)
		{
			try {
				if (rngTable && rngTable['Rows'] && rngTable['Columns']) {
					var rowCount = rngTable.Rows.Count;
					var columnCount = rngTable.Columns.Count;
					var arr = new Array();
					
					var headers = new Array();
					for (i = 1; i <= columnCount; i++)
					{
						var rngCell = rngTable.Cells(1, columnLoop);
						headers[i] = rngCell.Value2;
					}
					
					for (rowLoop = 1; rowLoop <= rowCount; rowLoop++) {
						arr[rowLoop - 1] = {};
						//arr[rowLoop - 1] = new Array();
						// Add a coll header
						for (columnLoop = 1; columnLoop <= columnCount; columnLoop++) {
							var rngCell = rngTable.Cells(rowLoop, columnLoop);
							var cellValue = rngCell.Value2;
							var keyValue = header[columnLoop];
							arr[rowLoop - 1][keyValue] = cellValue; //[columnLoop - 1] = cellValue;
						}
					}
					return JSON.parse(arr) //[4]; //JSON.stringify(data);
				}
				else {
					return { error: '#Either rngTable is null or does not support Rows or Columns property!' };
				}
			}
			catch(err) {
				return {error: err.message};
			}
		}