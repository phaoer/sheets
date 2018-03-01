        var oHead = document.getElementsByTagName('head')[0];

        var oScript = document.createElement("script");

        oScript.setAttribute("type", "text/javascript");

        oScript.setAttribute("src", "/javascripts/sheets/excelJs.js");

        oHead.appendChild(oScript);
        
        function sheets(option){
            console.log(option);
            this.id = option.id;
            this.callback = option.fun;
            this.rows = option.rows;
            this.cols = option.cols;
        }

        sheets.prototype.importExcel = function(){
            var _this = this;
            document.getElementById(this.id).onchange = function() {
                var rABS = false,sheetData,list,arrs=[];
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
                        document.getElementById("demo").innerHTML = JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]));
                        sheetData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
                        if(_this.rows == undefined && _this.cols == undefined){
                            list = sheetData;
                        }else if(_this.rows != undefined && _this.cols == undefined){
                            list = sheetData[_this.rows];
                        }else if(_this.rows == undefined && _this.cols != undefined){
                            list = [];
                            for(var i=0;i<sheetData.length;i++){
                               arrs=[];
                               for(var j in sheetData[i]){
                                  arrs.push(sheetData[i][j]);
                               }
                               list.push(arrs[_this.cols-1]);
                            }
                        }else{
                            for(var j in sheetData[_this.rows-1])
                            {
                                arrs.push(sheetData[_this.rows-1][j]);
                            }
                            list = arrs[_this.cols-1];
                        }
                        var all = {
                            list:list,
                            rowNum:sheetData.length,
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