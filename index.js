#!/usr/bin/env node

const fs = require('fs');
const fsPromises = fs.promises;
const process = require('process');
const path = require('path');

const chalk = require('chalk');
const XLSX = require('xlsx');
const yargs = require('yargs');

// Helper Functions

const getDefaultOutputPath = srcFilePath => {
  const {name} = path.parse(srcFilePath);
  const outPutFileName = `${name}.json`;

  return path.resolve(process.cwd(), outPutFileName);
};

const resolvePath = (srcFilePath, mustExist = false) => {
  const isRelative = !path.isAbsolute(srcFilePath);
  const resolvedPath = isRelative
    ? path.resolve(process.cwd(), srcFilePath)
    : srcFilePath;

  if (mustExist && !fs.existsSync(resolvedPath)) {
    throw new Error(
      `파일이 해당 경로에 존재하지 않습니다.\n입력값 : ${srcFilePath}`
    );
  }

  return resolvedPath;
};

const getWorkSheet = (workbook, sheetIndex = 0) => {
  const sheets = workbook.Sheets && Object.values(workbook.Sheets);

  if (!Array.isArray(sheets)) return;

  const worksheet = sheets[sheetIndex];

  return worksheet;
};

const sheetToArray = (worksheet, startingRow = 0) => {
  return XLSX.utils.sheet_to_json(worksheet, {
    // header: 1 옵션으로 worksheet를 2D Array로 변환한다.
    header: 1,
    // 공백인 줄은 output에 포함하지 않음.
    blankrows: false,
    range: startingRow,
  });
};

const transposeArr = arr =>
  arr.reduce((acc, row, rowIdx) => {
    row.forEach((cell, cellIdx) => {
      acc[cellIdx] = acc[cellIdx] || [];
      acc[cellIdx][rowIdx] = cell;
    });

    return acc;
  }, []);

const trimIfStr = value => {
  if (typeof value === 'string') return value.trim();

  return value;
};

const rowToMap = (rowData, propKeys) => {
  return rowData.reduce((acc, now, idx) => {
    const columnName = trimIfStr(propKeys[idx]);

    // propKey가 _이거나 지정된 index에 propKey가 없는 경우 결과에 포함시키지 않는다.
    if (!columnName || columnName === '_') return acc;

    acc[columnName] = trimIfStr(now);

    return acc;
  }, {});
};

// public api

const xlsxToJSON = (
  xlsxPath,
  propKeys,
  sheetIndex = 0,
  omitHeader = true,
  columnEntity = false
) => {
  validateParams(propKeys, sheetIndex);

  const resolvedXlsxPath = resolvePath(xlsxPath);
  const workbook = XLSX.readFile(resolvedXlsxPath, {
    cellHTML: false,
    cellFormula: false,
    cellText: false,
  });

  const worksheet = getWorkSheet(workbook, sheetIndex);
  const rowList = sheetToArray(worksheet, omitHeader ? 1 : 0);
  const targetData = columnEntity ? transposeArr(rowList) : rowList;
  const outputData = targetData.map(row => rowToMap(row, propKeys));

  return outputData;
};

exports.xlsxToJSON = xlsxToJSON;

// CLI

const log = {
  success: msg => {
    console.log(chalk`{green.bold ✅ ${msg}}`);
  },
  info: msg => {
    console.log(chalk`{cyan 📢 ${msg}}`);
  },
  error: (msg, e) => {
    console.log(chalk`{red.bold ❌ ${msg}}`);

    if (e) {
      console.log(chalk`\n{yellow 에러메시지 : ${e.message} }`);
    }
  },
};

const validateParams = (propKeys, sheetIndex) => {
  if (sheetIndex < 0) throw new Error('sheetIndex는 음수값이 될 수 없습니다.');

  const lodashRemoved = propKeys.filter(propKey => propKey !== '_');
  const lodashRemovedInSet = new Set(lodashRemoved);
  const isDistinct = lodashRemoved.length === lodashRemovedInSet.size;

  if (!isDistinct)
    throw new Error(
      'propKeys로 제공된 각 속성명들은 중복된 값이 없어야만 합니다.'
    );
};

yargs
  .command({
    command: 'gen',
    describe: 'xlsx 파일을 json으로 파싱해 저장한다.',
    builder: {
      from: {
        describe: 'xlsx 파일 경로',
        demandOption: true,
        type: 'string',
      },
      to: {
        describe: '저장할 json 파일 경로',
        type: 'string',
      },
      propKeys: {
        describe:
          'propKey로 사용될 리스트. _로 표시된 순번은 결과물에 포함하지 않는다.',
        type: 'array',
        demandOption: true,
      },
      sheetIndex: {
        describe: '변환할 시트 인덱스 (0부터 시작)',
        type: 'number',
        default: 0,
      },
      omitHeader: {
        describe: '첫번째 행 혹은 열을 생략할지 여부',
        type: 'boolean',
        default: true,
      },
      columnEntity: {
        describe: '데이터를 열 단위로 해석할지의 여부',
        type: 'boolean',
        default: false,
      },
    },
    handler: async ({
      from: xlsxPath,
      to: outputPath,
      propKeys,
      sheetIndex,
      omitHeader,
      columnEntity,
    }) => {
      log.info('xlsx 파일로부터 JSON 파일을 생성합니다...');
      log.info(`${sheetIndex} 번째 시트를 변환합니다.`);

      if (omitHeader) {
        log.info(
          `첫번째 ${columnEntity ? '열' : '행'}은 출력물에 포함시키지 않습니다.`
        );
      }

      log.info(
        `각 ${columnEntity ? '열' : '행'}들이 ${propKeys.join(
          ', '
        )} 의 속성명으로 매핑됩니다.\n`
      );

      try {
        validateParams(propKeys, sheetIndex);

        const outputData = xlsxToJSON(
          xlsxPath,
          propKeys,
          sheetIndex,
          omitHeader,
          columnEntity
        );

        const resolvedOutputPath = outputPath
          ? resolvePath(outputPath, false)
          : getDefaultOutputPath(xlsxPath);

        await fsPromises.writeFile(
          resolvedOutputPath,
          JSON.stringify(outputData, null, 2),
          {
            flag: 'w',
          }
        );

        log.success(
          `JSON 파일이 생성 되었습니다.\n\n경로 : ${resolvedOutputPath}`
        );
      } catch (e) {
        log.error('JSON 변환에 실패하였습니다.', e);
      }
    },
  })
  .parse();
