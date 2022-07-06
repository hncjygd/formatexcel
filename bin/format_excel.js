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
    await wb.xlsx.readFile(filePath);
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
    let tasks = filePaths.map((value) => 
    {
        const toPath = path.resolve(options.outpath, path.relative(options.inpath, value));
        return { title: value, task: () => format(value, toPath) }

    });
    const limit = 10; // 超过一定数量之后Listr会出现问题
    for (let i=0; i<Math.ceil(tasks.length/limit); i++) {
        await (new Listr(tasks.slice(i*limit, (i+1)*limit)).run());
    }
    // await (new Listr(tasks).run());
})();