/**
 * @copyright
 * @author ngochuy
 * @create 2022-02-28
 * @update 2022-02-28
 * @Description
 * @note  Known Issue: Any merged cells to move, the results may be unpredictable
 */
'use strict';

/* eslint-disable camelcase */
const _ = require('lodash');
const fs = require('fs');
const excel = require('exceljs');

const exceljs = {
    populate: async (template, data) => {
        //  template = fs.readFileSync('./api/hooks/report/mock/Tempate01.xlsx');
        template = fs.readFileSync('./mock/Force_Ranking_Tempalte.xlsx');
        //  data = require('./mock/report_data.json');
        data = require('./mock/Force_Ranking_data.json');

        const workbook = new excel.Workbook();
        await workbook.xlsx.load(template);
        const worksheet = workbook.worksheets[0];

        populate(worksheet, data);

        await workbook.xlsx.writeFile('ReportName_' + Date.now() + '.xlsx');
        console.log('done');

        // return await workbook.xlsx.writeBuffer();
    }
};

module.exports = exceljs;

exceljs.populate();

const helper = {
    populate: (worksheet, data) => {
        let pos;

        populate_master_data(worksheet, data[0]);

        const group_temp = helper.get_group_temp(worksheet);
        if (group_temp.length) {
            // [
            //     { GroupFeildID: ['GoalPlanID'], row: Row, level: 0 },
            //     { FeildID: ['GoalPlanID', 'GroupGoalID'], row: Row, level: 'last', table_temp: Row }
            // ]
            row_pos = group_temp[0].row.number + group_temp.length;
            pos = populate_group_excel(worksheet, group_temp, row_pos, data);

            // remove temp row
            for (const rows_temp of group_temp) {
                // worksheet.spliceRows(rows_temp.row.number + 1, rows_temp.row.number);
                // worksheet.spliceRows(31, 30);  => bug
                rows_temp.row.hidden = true;
                if (rows_temp.table_temp) {
                    rows_temp.table_temp.hidden = true;
                }
            }
        } else {
            pos = populate_table_excel(worksheet, data);
        }
        return pos;
    },
    create_group_row: (worksheet, row, pos, data) => {
        const newRow = worksheet.insertRow(pos + 1, {});

        newRow.height = row.height;
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            let newCell = newRow.getCell(colNumber);
            if (cell.isMerged && cell.type === cell.Merge) {
                return;
            }
            newCell.style = cell.style;
            newCell.value = cell.value;
        });

        newRow.eachCell({ includeEmpty: true }, cell => {
            if (cell.isMerged && cell.type === cell.Merge) {
                return;
            }
            if (cell.value && typeof cell.value === 'string' && cell.value.includes('${group')) {
                // ${group[GoalPlanID]:GoalPlanName}
                const feildID = cell.value.substring(cell.value.indexOf(':') + 1, cell.value.indexOf('}'));
                display_cell_values(cell, data[feildID]);
            }

            let height = 0;
            if (cell.value && cell.alignment && cell.alignment.wrapText === true) {
                height = String(cell.value).trim().length > 20 ? String(cell.value).trim().length * 2 / 3 : 20;
            }
            if (height > newRow.height) {
                newRow.height = height;
            }
            return;
        });

        return newRow;
    },
    get_group_temp: (worksheet) => {
        let group_rows_temp = [];
        let arr_feildID = [];
        worksheet.eachRow({ includeEmpty: false }, row => {
            row.eachCell(c => {
                if (group_rows_temp.find(r => r.row.number === c.row)) {
                    return;
                }

                if (c.value && typeof c.value === 'string' && c.value.includes('${group')) {
                    // ${group[GoalPlanID]:GoalPlanName}
                    const feildID = c.value.substring(c.value.indexOf('[') + 1, c.value.indexOf(']'));
                    arr_feildID.push(feildID);
                    group_rows_temp.push({ row, GroupFeildID: [...arr_feildID], level: group_rows_temp.length });
                }
            });
        });

        if (group_rows_temp && group_rows_temp[0]) {
            group_rows_temp[group_rows_temp.length - 1].level = 'last';

            const table_temp = helper.get_table_temp(worksheet);
            group_rows_temp[group_rows_temp.length - 1].table_temp = table_temp;
        }

        return group_rows_temp;
    },
    get_table_temp: (worksheet) => {
        let table_temp;
        worksheet.eachRow({ includeEmpty: false }, row => {
            row.eachCell(c => {
                if (table_temp) {
                    return;
                }

                if (c.value && typeof c.value === 'string' && c.value.includes('${table:')) {
                    // ${table:GoalName}
                    table_temp = row;
                }
            });
        });

        return table_temp;
    },
    create_table_row: (worksheet, row, index, data) => {

        const newRow = worksheet.insertRow(index + 1, {});

        newRow.height = row.height;
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            var newCell = newRow.getCell(colNumber);

            if (cell.isMerged && cell.type === cell.Merge) {
                return;
            }
            newCell.style = cell.style;
            newCell.value = cell.value;

        });

        newRow.eachCell({ includeEmpty: true }, cell => {
            if (cell.isMerged && cell.type === cell.Merge) {
                return;
            }
            if (cell.value && typeof cell.value === 'string' && cell.value.includes('${table:')) {
                const feildID = cell.value.substring(cell.value.indexOf(':') + 1, cell.value.indexOf('}'));
                // cell.value = data[feildID] || '';
                display_cell_values(cell, data[feildID]);
            }

            let height = 0;
            if (cell.value && cell.alignment && cell.alignment.wrapText === true) {
                height = String(cell.value).trim().length > 20 ? String(cell.value).trim().length * 2 / 3 : 20;
            }
            if (height > newRow.height) {
                newRow.height = height;
            }
            return;
        });

        return newRow;
    },
};

function populate(worksheet, data) {
    let pos;

    populate_master_data(worksheet, data[0]);

    const group_temp = helper.get_group_temp(worksheet);
    if (group_temp.length) {
        // [
        //     { GroupFeildID: ['GoalPlanID'], row: Row, level: 0 },
        //     { FeildID: ['GoalPlanID', 'GroupGoalID'], row: Row, level: 'last', table_temp: Row }
        // ]
        const row_pos = group_temp[0].row.number + group_temp.length;
        pos = populate_group_excel(worksheet, group_temp, row_pos, data);

        // remove temp row
        for (const rows_temp of group_temp) {
            // worksheet.spliceRows(rows_temp.row.number + 1, rows_temp.row.number);
            // worksheet.spliceRows(31, 30);  => bug
            rows_temp.row.hidden = true;
            if (rows_temp.table_temp) {
                rows_temp.table_temp.hidden = true;
            }
        }
    } else {
        pos = populate_table_excel(worksheet, data);
    }
    return pos;
}

function populate_group_excel(worksheet, rows_temps, pos, data, group_level = 0) {
    if (!rows_temps[group_level]) {
        return pos;
    }
    let current_pos = pos;

    const group_data = _.groupBy(data, (e) => {
        const group = rows_temps[group_level].GroupFeildID.map(m => e[m]);
        return group;
    });

    for (const key in group_data) {
        // let level = rows_temps[group_level].level;
        // console.log('level', level)

        const group_data_row = group_data[key][0]; // fist object

        let new_row = helper.create_group_row(worksheet, rows_temps[group_level].row, current_pos, group_data_row);
        current_pos = new_row.number;

        if (rows_temps[group_level].level === 'last') {
            for (const table_row_data of group_data[key]) {
                const table_row = helper.create_table_row(worksheet, rows_temps[group_level].table_temp, current_pos, table_row_data);
                current_pos = table_row.number;
            }
        }
        else {
            current_pos = populate_group_excel(worksheet, rows_temps, current_pos, group_data[key], group_level += 1); // next level
        }

        if (key.includes(',')) {
            continue;
        }
        group_level = 0;
    }

    return current_pos;
}

function populate_master_data(worksheet, data) {
    worksheet.eachRow({ includeEmpty: false }, row => {
        row.eachCell(c => {
            if (c.value && typeof c.value === 'string' && c.value.includes('${') && !c.value.includes('${group') && !c.value.includes('${table')) {
                //ex: ${EmployeeName}

                const array = c.value.split('$');
                for (let index = 1; index < array.length; index++) {
                    const feildID = c.value.substring(c.value.indexOf('{') + 1, c.value.indexOf('}'));

                    const re = new RegExp(`\\$\{${feildID}\}`, 'g');
                    const value = c.value.replace(re, data[feildID] || '');

                    display_cell_values(c, value);
                    // console.log(c.value)
                }
            }
        });
    });
}

function display_cell_values(cell, values) {
    // {
    //     Null: 0,
    //     Merge: 1,
    //     Number: 2,
    //     String: 3,
    //     Date: 4,
    //     Hyperlink: 5,
    //     Formula: 6,
    //     SharedString: 7,
    //     RichText: 8,
    //     Boolean: 9,
    //     Error: 10,
    //   }

    // if (cell.value === '${DivisionID1} - ${DivisionName1}') {
    //     console.log('')
    // }
    // if (values == '90.916') {
    //     console.log('')
    // }

    // switch (cell.type) {
    //     case 2:
    //         c.numFmt = '0.00'
    //         if (values) {
    //             cell.value = parseFloat(values);
    //             break;
    //         }
    //         cell.value = 0;
    //         break;
    //     default:
    //         cell.value = values;
    // }

    switch (cell.numFmt) {
        case '0':
        case '0.00':
        case '0%':
        case '0.00%':
            if (values) {
                cell.value = parseFloat(values);
                break;
            }
            cell.value = 0;
            break;
        default:
            cell.value = values || '';
    }
    return;
}

function populate_table_excel(worksheet, data) {
    let pos;
    const table_temp = helper.get_table_temp(worksheet);
    if (!table_temp) {
        return;
    }

    for (const row_data of data) {
        const new_table_row = helper.create_table_row(worksheet, table_temp, table_temp.number, row_data);
        pos = new_table_row.number;
    }

    table_temp.hidden = true;
    return pos;
}
