# xy_image_optimu.js v1.0.0
迅游sheets处理套件，**包含excel，word，txt读写         迅游版权所有       Create By Pwh**

## 使用方法

```html
<script src="js/xy_excel.js"></script>
    <script type="text/javascript">
        window.onload = function(){
            var data = {
                id:"file", //上传控件id
                fun:"fun", //回调函数名
                rows:2,    //获取第几行的数据（获取每行则不填）
                cols:1,    //获取第几列的数据（获取每列则不填）
            }
            var sheet = new sheets(data);
            sheet.importExcel();
        }
        function fun(data){
            console.log(data);
            //data.list 返回结果集 若是表格全部数据 list为数组，包含所有列的数据[{第一列},{第二列},{第三列}...]
                                // 若是指定某行 同上
                                // 若是指定某列 list为数组，包含该列数据[a,b,c,d.....]
                                // 若是指定了行列 则list返回字符串
            //data.rowNum 数据总行数(方便以后拓展分页)
        }                       
</script>
```

### The Relentless Pursuit of Perfection    持续更新中