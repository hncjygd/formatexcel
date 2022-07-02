#!/usr/bin/env node
import { program } from 'commander';
import fs from 'fs';
import path from 'path';
import excel from 'exceljs';
import { cnopendataFormat } from '../lib/format.js';
import Listr from 'listr';

program
    .option('-o, --outpath <string>', '转换后格式保存的文件夹, 默认保存在当前文件夹', './')
    .requiredOption('-i, --inpath <string>', '需要转换格式的excel所在文件夹')
    .option('-l, --limit <number>', '限制线程数，1~10之间', 5)
    .version('0.1.0')
    .parse(process.argv);

const options = program.opts();
// 输入路径错误提示
if (!fs.existsSync(options.inpath)) {
    console.log(options.inpath)
    console.log(`需要转换文件所在的文件夹不存在，确认指定的"-i,--input ${options.inpath}"选项是否正确`);
    process.exit(1);
}
// 输出路径错误提示
if (!options.outpath || !fs.existsSync(options.outpath)) {
    console.log(`保存转换后文件的文件夹不存在，确认指定的"-o,--output ${options.outpath}"选项是否正确`);
    process.exit(1);
}

if (!parseInt(options.limit)){
    console.log(`"-l,--limit ${options.limit}"选项必须是一个数值`);
    process.exit(1);
} else {
    const limit = parseInt(options.limit);
    if (limit > 10 || limit <= 0) {
        console.log(`"-l,--limit ${options.limit}的值超过1~10的允许范围，被强制设置为默认值5!"`);
        options.limit = 5;
    }
}

// 递归获取文件夹下所有的excel文件
function getFile(fp, all) {
    for (let f of fs.readdirSync(fp, { withFileTypes: true })) {
        const abPath = path.resolve(fp, f.name)
        if (path.extname(f.name) === '.xlsx') {
            all.push(abPath);
        }
        if (f.isDirectory()) {
            getFile(abPath, all);
        }
    }
}

async function format(filePath, toPath) {
    let wb = new excel.Workbook();
    wb = await wb.xlsx.readFile(filePath);
    cnopendataFormat(wb);
    const dir = path.dirname(toPath);
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
    await wb.xlsx.writeFile(toPath);
}

(async () => {
    const filePaths = [];
    getFile(options.inpath, filePaths);
    if (parseInt(options.limit) > filePaths) {
        console.log(`"-l,--limit ${options.limit}的值超过操过文件数量，被强制设置为与文件数量相同的线程!"`);
        options.limit = filePaths.length;
    }
    let tasks = filePaths.map((value) => 
    {
        const toPath = path.resolve(options.outpath, path.relative(options.inpath, value));
        return { title: value, task: () => format(value, toPath) }

    });
    new Listr(tasks, {concurrent: parseInt(options.limit), exitOnError: false }).run();
})();