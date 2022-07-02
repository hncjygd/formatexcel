/**
 * 用于将excel文件整理成规范的格式:
 * 1. 更改文档的作者、最后修改者为CnOpenData
 * 2. 首行蓝底白字加粗，字体微软雅黑
 * 3. 添加边框，颜色蓝色
 * 4. 最后一行空一行添加 `数据来源: CnOpenData`字样, 同样蓝底白字
 */

/**
 * cnopendata内部的格式化工具
 * @param {Workbook} wb - Workbook对象 
 */

export let cnopendataFormat = function (wb) {
    // 1. 更改文档的作者、最后修改者为CnOpenData 
    wb.creator = 'CnOpenData';
    wb.lastModifiedBy = 'CnOpenData';
    const ws = wb.getWorksheet(1);
    // 2. 字体格式 蓝底白字加粗，字体微软雅黑
    const firstFont = {
        name: 'Microsoft YaHei',
        color: { argb: 'FFFFFFFF' },
        fammily: 2,
        bold: true,
    };
    // 2. 更改首行为字体格式
    ws.getRow(1).eachCell(
        { includeEmpty: true },
        (cell) => {
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF4F83BD' },
                // bgColor: { argb: 'FFFFFFFF' },
            };
            cell.font = firstFont
        }
    );
    // 3. 边框格式
    const border = {
        top: { style: 'thin', color: { argb: 'FF4F83BD' } },
        left: { style: 'thin', color: { argb: 'FF4F83BD' } },
        bottom: { style: 'thin', color: { argb: 'FF4F83BD' } },
        right: { style: 'thin', color: { argb: 'FF4F83BD' } }
    };
    // 3. 为所有单元格应用边框格式
    ws.eachRow(
        { includeEmpty: true },
        (row) => row.eachCell((cell) => cell.border = border)
    );
    // 4. 最后一行空一行添加 `数据来源: CnOpenData`字样, 同样蓝底白字  
    const cell = ws.getRow(ws.lastRow.number + 2).getCell(1);
    cell.value = '数据来源: CnOpenData';
    const lastFont = {
        name: 'Microsoft YaHei',
        color: { argb: 'FF4F83BD' },
        fammily: 2,
        bold: true,
    };
    cell.font = lastFont;
}
