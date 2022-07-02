import exceljs from 'exceljs';
import { cnopendataFormat } from './format.js';
import path, { extname, dirname } from 'path';
import { existsSync, mkdirSync } from 'fs';

/** 
 * 
 * @param {Array.<{header: string, key: string}>} headKeys - 数组元素为: {header: string, key: string} header指定表格的表头, key指定列键
 * @param {Array.<{string: any}>} rowData - 要写入的数据集，数组中元素为{key1:value2, key2:value2...} 其中key就是headKeys传入的列键, 表格的行数由数组长度决定
 * @param {string} filePath - 路径字符串,注意包含文件名的路径(可以是相对路径),如果路径文件夹不存在则会自动创建
 * @param {function(Workbook): void} [isFormat=cnopendataFormat] - 使用的格式化程序，默认是cnopendata格式化
 * @returns {Promise<void>} - 期望文件写入成功
 */
export function exportExcel(headKeys, rowData, filePath, format = cnopendataFormat) {
    if (typeof(filePath)!=='string' && extname(path) !== 'xlsx') {
        throw Error('filePath必须是文件的全路径字符串，文件后缀必须是.xlsx');
    }
    // 如果文件夹不存在将创建
    const dirPath = dirname(filePath)
    if (!existsSync(dirPath)) mkdirSync(dirPath, {recursive : true});
    const wb = new exceljs.Workbook();
    const ws = wb.addWorksheet();
    ws.columns = headKeys;
    ws.addRows(rowData);
    // 应用格式化
    format(wb);
    return wb.xlsx.writeFile(filePath);
}