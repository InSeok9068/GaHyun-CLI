#! /usr/bin/env node

import { select } from '@inquirer/prompts';
import chalk from 'chalk';
import { program } from 'commander';
import figlet from 'figlet';
import { join } from 'path';
import XLSX from 'xlsx';
import { __dirname } from './utils/path.util.js';

console.log(chalk.yellow(figlet.textSync('YOON GA HYUN CLI')));

program
  .version('1.0.0')
  .description('실적 엑셀 통계 CLI')
  .action(async () => {
    const answer = await select({
      message: '무엇을 도와드릴까요?',
      choices: [
        {
          name: '실적 엑셀 통계 생성',
          value: 'stat',
        },
      ],
    });

    switch (answer) {
      case 'stat':
        console.log(XLSX.readFile(join(__dirname, '../../', 'storage/files/copy/sample.xlsm')));
        break;
    }
  });

program.parse(process.argv);
