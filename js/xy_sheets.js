/** xy_sheets-v1.0.1 By https://github.com/phaoer/sheets */
var oHead = document.getElementsByTagName('head')[0];

var oScript = document.createElement("script");

oScript.setAttribute("type", "text/javascript");

oScript.setAttribute("src", "/javascripts/sheets/excelJs.js");

oHead.appendChild(oScript);

if ($) {
    (function($) {
        $.fn.extend({
            insertContent: function(myValue) {
                var $t = $(this)[0];
                if (document.selection) {
                    this.focus();
                    var sel = document.selection.createRange();
                    sel.text = myValue;
                    this.focus();
                } else
                if ($t.selectionStart || $t.selectionStart == '0') {
                    var startPos = $t.selectionStart;
                    var endPos = $t.selectionEnd;
                    var scrollTop = $t.scrollTop;
                    $t.value = $t.value.substring(0, startPos) + myValue + $t.value.substring(endPos, $t.value.length);
                    this.focus();
                    $t.selectionStart = startPos + myValue.length;
                    $t.selectionEnd = startPos + myValue.length;
                    $t.scrollTop = scrollTop;
                } else {
                    this.value += myValue;
                    this.focus();
                }
            }
        });
    })(jQuery);
}

function sheets(option) {
    this.id = option.id;
    this.callback = option.fun;
    this.rows = option.rows;
    this.cols = option.cols;
}

sheets.prototype.importExcel = function() {
    var _this = this;
    document.getElementById(this.id).onchange = function() {
        var rABS = false,
            sheetData, list, arrs = colsName = colsArrs = rowsArrs = [],
            colsNum = 0,
            rowsJson = colsJson = {};
        if (!this.files) { return; }
        var f = this.files[0]; {
            var reader = new FileReader();
            var name = f.name;
            reader.onload = function(e) {
                var data = e.target.result;
                var wb;
                if (rABS) {
                    wb = XLSX.read(data, { type: 'binary' });
                } else {
                    var arr = fixdata(data);
                    wb = XLSX.read(btoa(arr), { type: 'base64' });
                }
                // document.getElementById("demo").innerHTML = JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
                sheetData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
                console.log(sheetData);
                if ((_this.rows == undefined && _this.cols == undefined)) {
                    list = sheetData;
                } else if (_this.rows != undefined && _this.cols == undefined) {    //获取列数据
                    if (_this.rows.length == undefined) {        //只获取一列
                        list = sheetData[_this.rows - 1];
                    } else {
                        if (_this.rows.length == 0) {         //获取所有列
                            var rl = [];         
                            for (var j = 0; j < sheetData.length; j++) {     //得到列的总数
                                rl.push(j + 1);
                            }
                            for (var k = 0; k < rl.length; k++) {
                                rowsArrs = [];
                                rowsArrs.push(sheetData[rl[k] - 1]);
                                rowsJson[rl[k]] = rowsArrs
                            }
                            list = rowsJson;
                        } else {                               //获取指定列
                            for (var k = 0; k < _this.rows.length; k++) {
                                rowsArrs = [];
                                rowsArrs.push(sheetData[_this.rows[k] - 1]);
                                rowsJson[_this.rows[k]] = rowsArrs
                            }
                            list = rowsJson;
                        }
                    }
                } else if (_this.rows == undefined && _this.cols != undefined) { //获取行数据
                    list = [];
                    if (_this.cols.length == undefined) { //只获取一行
                        for (var i = 0; i < sheetData.length; i++) {
                            arrs = [];
                            for (var j in sheetData[i]) {
                                arrs.push(sheetData[i][j]);
                            }
                            list.push(arrs[_this.cols - 1]);
                        }
                    } else {
                        if (_this.cols.length == 0) {     //获取所有行
                            var num = 0,
                                cl = [];
                            for (var i in sheetData[0]) {
                                num++;
                            }                                        ////得到行的总数
                            for (var j = 0; j < num; j++) {
                                cl.push(j + 1);
                            }
                            for (var k = 0; k < cl.length; k++) {
                                colsArrs = [];
                                for (var i = 0; i < sheetData.length; i++) {
                                    arrs = [];
                                    for (var j in sheetData[i]) {
                                        arrs.push(sheetData[i][j]);
                                    }
                                    colsArrs.push(arrs[cl[k] - 1]);
                                }
                                colsJson[cl[k]] = colsArrs;
                            }
                            list = colsJson;
                        } else {                          //获取指定行
                            for (var k = 0; k < _this.cols.length; k++) {
                                colsArrs = [];
                                for (var i = 0; i < sheetData.length; i++) {   //遍历数据
                                    arrs = [];
                                    for (var j in sheetData[i]) {      
                                        arrs.push(sheetData[i][j]);          //获取每条数据中每一列表头对应的值，并通过数组进行第一列到最后一列的排序，得到列的排序数组
                                    }
                                    colsArrs.push(arrs[_this.cols[k] - 1]);      //按照用户传入的所需的行在列的排序数组中取值并存入数组，得到结果数组
                                }
                                colsJson[_this.cols[k]] = colsArrs;             //按照用户传入的所需的行将结果数组对应编号返回给用户
                            }
                            list = colsJson;
                        }
                    }
                } else {
                    for (var j in sheetData[_this.rows - 1]) {
                        arrs.push(sheetData[_this.rows - 1][j]);
                    }
                    list = arrs[_this.cols - 1];
                }
                for (var l in sheetData[0]) {
                    colsNum++;
                    colsName.push(l);
                }
                var all = {
                    list: list,
                    fileName: name,
                    colsName: colsName,
                    rowsNum: sheetData.length,
                    colsNum: colsNum
                }
                eval(_this.callback + "(all)");
            };
            if (rABS) reader.readAsBinaryString(f);
            else reader.readAsArrayBuffer(f);
        }
    }
}

function fixdata(data) {
    var o = "",
        l = 0,
        w = 10240;
    for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
}