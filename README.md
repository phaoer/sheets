# xy_sheets.js v1.0.0
迅游sheets处理套件，**包含excel，word，txt读写 Create By Pwh**

## 使用方法

```html
<script src="js/xy_sheets.js"></script>
    <script type="text/javascript">
        window.onload = function(){
            var data = {
                id:"file", //上传控件id
                fun:"fun", //回调函数名
                // rows:[],    //获取第几行的数据，支持数组获取多行（获取每行则不填则数组填为空）
                cols: [] //获取第几列的数据，支持数组获取多列（获取每列则数组填为空）
            }
            var sheet = new sheets(data);
            sheet.importExcel();
        }
        function fun(data){
            console.log(data);
            //data.list 返回结果集 若是表格全部数据 list为数组，包含所有列的数据[{第一列},{第二列},{第三列}...]
            // 若是指定某一行 list为数组，包含该列数据[a,b,c,d.....]，若指定多行，返回json，包含行数{1:[...,...,..],2[...,...,...],3:[...,...,...].....}
            // 若是指定某一列 同上
            // 若是指定了行列 则list返回字符串
            //data.rowsNum 数据总行数(方便以后拓展分页)
            //data.colsNum 数据总列数
        }                       
</script>
```

## v1.0.1更新
   rows,cols支持数组传入获取指定多行 若为空则是所有行和所有列

### The Relentless Pursuit of Perfection    持续更新中
