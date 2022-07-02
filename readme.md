## 说明

用于导出并格式化excel。


## excel导出使用说明

```javascript
import { exportExcel } from 'formatexcel';

// header-表头  key-列键
const headKeys = [
    {'header':'序号', 'key': 'id'},
    {'header':'名字', 'key': 'name'},
];
// 表格的行数由rowData数组长度决定
const rowData = [
    {id: 1, name: 'example'},
    {id: 2, name: 'example1'},
    {id: 3, name: 'example2'},
    {id: 4, name: 'example3'},
]

// 默认使用cnopendata格式化
exportExcel(headKeys, rowData, './example.xlsx', format=youOwnFormatFunc).then(console.log)
```

## excel格式化工具

提供了一个`format_excel`命令来格式化已经存在的excel文件,格式化样式如下:

1. 更改文档的作者、最后修改者为CnOpenData
2. 首行蓝底白字加粗，字体微软雅黑
3. 添加边框，颜色蓝色
4. 最后一行空一行添加 `数据来源: CnOpenData`字样, 同样蓝底白字

```bash
Usage: format_excel [options]

Options:
  -o, --outpath <string>  转换后格式保存的文件夹, 默认保存在当前文件夹 (default: "./")
  -i, --inpath <string>   需要转换格式的excel所在文件夹
  -l, --limit <number>    限制线程数，1~10之间 (default: 5)
  -V, --version           output the version number
  -h, --help              display help for command
```