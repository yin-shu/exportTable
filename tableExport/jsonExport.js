//把json数组导出到Excel
//由tableExport改造
//实现：支持中文，导出xml支持大量数据

function tableExport(options) {
	var defaults = {
		arrData: [],

		consoleLog: false,
		displayTableName: false,
		escape: false,
		excelstyles: ['border-bottom', 'border-top', 'border-left', 'border-right'],
		fileName: '导出数据',
		htmlContent: false,
		ignoreColumn: [],
		ignoreRow: [],
		outputMode: 'file', // file|string|base64
		tableName: 'myTableName',
		type: 'csv',
		worksheetName: 'xlsWorksheetName'
	};

	var el = this;
	var DownloadEvt = null;
	var rowIndex = 0;
	var rowspans = [];
	var trData = '';

	jQuery.extend(true, defaults, options);

	if(defaults.type == 'excel') {
		//console.log($(this).html());

		rowIndex = 0;
		var excelData = "<table>";
		if(defaults.displayTableName)
			excelData += "<tr><td style='text-align:center' colspan='"+Object.getOwnPropertyNames(defaults.arrData[0]).length+"'>" + defaults.tableName + "</td></tr>";

		getHead(defaults.arrData, rowIndex, function(cell, row, col) {
			if(cell != null) {
				trData += "<td style='text-align:center;";
				for(var styles in defaults.excelstyles) {
					if(defaults.excelstyles.hasOwnProperty(styles)) {
						//trData += defaults.excelstyles[styles] + ": 1px solid black;";
					}
				}
				trData += "'>" + cell + "</td>";
			}
		})
		if(trData.length > 0)
			excelData += "<tr>" + trData + '</tr>';
		rowIndex++;
		defaults.arrData.forEach(function(item) {
			trData = "";
			getContent(item, rowIndex, function(cell, row, col) {
				if(cell != null) {
					trData += "<td style='";
					for(var styles in defaults.excelstyles) {
						if(defaults.excelstyles.hasOwnProperty(styles)) {
							//trData += defaults.excelstyles[styles] + ": 1px solid black;";
						}
					}
					trData += "'>" + cell + "</td>";
				}
			})
			if(trData.length > 0)
				excelData += "<tr>" + trData + '</tr>';
			rowIndex++;
		})

		
		excelData += '</table>';

		if(defaults.consoleLog === true)
			console.log(excelData);

		var excelFile = "<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:x='urn:schemas-microsoft-com:office:" + defaults.type + "' xmlns='http://www.w3.org/TR/REC-html40'>";
		excelFile += '<meta http-equiv="content-type" content="application/vnd.ms-' + defaults.type + '; charset=UTF-8">';
		excelFile += '<meta http-equiv="content-type" content="application/';
		excelFile += (defaults.type === 'excel') ? 'vnd.ms-excel' : 'msword';
		excelFile += '; charset=UTF-8">';
		excelFile += "<head>";
		if(defaults.type === 'excel') {
			excelFile += "<!--[if gte mso 9]>";
			excelFile += "<xml>";
			excelFile += "<x:ExcelWorkbook>";
			excelFile += "<x:ExcelWorksheets>";
			excelFile += "<x:ExcelWorksheet>";
			excelFile += "<x:Name>";
			excelFile += defaults.worksheetName;
			excelFile += "</x:Name>";
			excelFile += "<x:WorksheetOptions>";
			excelFile += "<x:DisplayGridlines/>";
			excelFile += "</x:WorksheetOptions>";
			excelFile += "</x:ExcelWorksheet>";
			excelFile += "</x:ExcelWorksheets>";
			excelFile += "</x:ExcelWorkbook>";
			excelFile += "</xml>";
			excelFile += "<![endif]-->";
		}
		excelFile += "</head>";
		excelFile += "<body>";
		excelFile += excelData;
		excelFile += "</body>";
		excelFile += "</html>";

		if(defaults.outputMode == 'string')
			return excelFile;

		var base64data = base64encode(excelFile);

		if(defaults.outputMode === 'base64')
			return base64data;

		var extension = (defaults.type === 'excel') ? 'xls' : 'doc';
		try {
			var blob = new Blob([excelFile], { type: 'application/vnd.ms-' + defaults.type });
			//saveAs(blob, defaults.fileName + '.' + extension);

			//改动:保存方式由"另存为"改为"下载",实现导出大量数据
			//this trick will generate a temp "a" tag
			var link = document.createElement("a");
			link.id = "lnkDwnldLnk";

			//this part will append the anchor tag and remove it after automatic click
			document.body.appendChild(link);
			var csvUrl = window.webkitURL.createObjectURL(blob);
			var filename = defaults.fileName + '.' + extension;
			$("#lnkDwnldLnk")
				.attr({
					'download': filename,
					'href': csvUrl
				});

			$('#lnkDwnldLnk')[0].click();
			document.body.removeChild(link);

		} catch(e) {
			downloadFile(defaults.fileName + '.' + extension, 'data:application/vnd.ms-' + defaults.type + ';base64,' + base64data);
		}

	}

	function getHead(arr, rowIndex, callBack) {
		var col = 0;
		for(var item in arr[0]) {
			callBack(item, rowIndex, col);
			col++;
		}
	}

	function getContent(obj, rowIndex, callBack) {

		var col = 0;
		for(var cell in obj) {
			callBack(obj[cell], rowIndex, col);
			col++;
		}

	}

	function ForEachVisibleCell(tableRow, selector, rowIndex, cellcallback) {
		if(defaults.ignoreRow.indexOf(rowIndex) == -1) {
			$(tableRow).filter(':visible').find(selector).each(function(colIndex) {
				if($(this).data("tableexport-display") == 'always' ||
					($(this).css('display') != 'none' &&
						$(this).css('visibility') != 'hidden' &&
						$(this).data("tableexport-display") != 'none')) {
					if(defaults.ignoreColumn.indexOf(colIndex) == -1) {
						if(typeof(cellcallback) === "function") {
							var cs = 0; // colspan value

							// handle previously detected rowspans
							if(typeof rowspans[rowIndex] != 'undefined' && rowspans[rowIndex].length > 0) {
								for(c = 0; c <= colIndex; c++) {
									if(typeof rowspans[rowIndex][c] != 'undefined') {
										cellcallback(null, rowIndex, c);
										delete rowspans[rowIndex][c];
										colIndex++;
									}
								}
							}

							// output content of current cell
							cellcallback(this, rowIndex, colIndex);

							// handle colspan of current cell
							if($(this).is("[colspan]")) {
								cs = $(this).attr('colspan');
								for(c = 0; c < cs - 1; c++)
									cellcallback(null, rowIndex, colIndex + c);
							}

							// store rowspan for following rows
							if($(this).is("[rowspan]")) {
								var rs = parseInt($(this).attr('rowspan'));

								for(r = 1; r < rs; r++) {
									if(typeof rowspans[rowIndex + r] == 'undefined')
										rowspans[rowIndex + r] = [];
									rowspans[rowIndex + r][colIndex] = "";

									for(c = 1; c < cs; c++)
										rowspans[rowIndex + r][colIndex + c] = "";
								}
							}
						}
					}
				}
			});
		}
	}

	function escapeRegExp(string) {
		return string.replace(/([.*+?^=!:${}()|\/\\])/g, "\\$1");
	}

	function replaceAll(string, find, replace) {
		return string.replace(new RegExp(escapeRegExp(find), 'g'), replace);
	}




	
	function hyphenate(a, b, c) {
		return b + "-" + c.toLowerCase();
	}

	
	function downloadFile(filename, data) {
		var DownloadLink = document.createElement('a');

		if(DownloadLink) {
			document.body.appendChild(DownloadLink);
			DownloadLink.style = 'display: none';
			DownloadLink.download = filename;
			DownloadLink.href = data;

			if(document.createEvent) {
				if(DownloadEvt == null)
					DownloadEvt = document.createEvent('MouseEvents');

				DownloadEvt.initEvent('click', true, false);
				DownloadLink.dispatchEvent(DownloadEvt);
			} else if(document.createEventObject)
				DownloadLink.fireEvent('onclick');
			else if(typeof DownloadLink.onclick == 'function')
				DownloadLink.onclick();

			document.body.removeChild(DownloadLink);
		}
	}

	function utf8Encode(string) {
		string = string.replace(/\x0d\x0a/g, "\x0a");
		var utftext = "";
		for(var n = 0; n < string.length; n++) {
			var c = string.charCodeAt(n);
			if(c < 128) {
				utftext += String.fromCharCode(c);
			} else if((c > 127) && (c < 2048)) {
				utftext += String.fromCharCode((c >> 6) | 192);
				utftext += String.fromCharCode((c & 63) | 128);
			} else {
				utftext += String.fromCharCode((c >> 12) | 224);
				utftext += String.fromCharCode(((c >> 6) & 63) | 128);
				utftext += String.fromCharCode((c & 63) | 128);
			}
		}
		return utftext;
	}

	function base64encode(input) {
		var keyStr = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";
		var output = "";
		var chr1, chr2, chr3, enc1, enc2, enc3, enc4;
		var i = 0;
		input = utf8Encode(input);
		while(i < input.length) {
			chr1 = input.charCodeAt(i++);
			chr2 = input.charCodeAt(i++);
			chr3 = input.charCodeAt(i++);
			enc1 = chr1 >> 2;
			enc2 = ((chr1 & 3) << 4) | (chr2 >> 4);
			enc3 = ((chr2 & 15) << 2) | (chr3 >> 6);
			enc4 = chr3 & 63;
			if(isNaN(chr2)) {
				enc3 = enc4 = 64;
			} else if(isNaN(chr3)) {
				enc4 = 64;
			}
			output = output +
				keyStr.charAt(enc1) + keyStr.charAt(enc2) +
				keyStr.charAt(enc3) + keyStr.charAt(enc4);
		}
		return output;
	}

	return this;
}