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
        // template = fs.readFileSync('./mock/Tempate01.xlsx');
        template = fs.readFileSync('./mock/Mau_BC_CompetencyCaNhan_1.xlsx');
        //template = fs.readFileSync('./ReportName_1646205694408.xlsx');
        // template = fs.readFileSync('./mock/Force_Ranking_Tempalte.xlsx');

        // data = require('./mock/Tempate01_data.json');
        data = require('./mock/Mau_BC_CompetencyCaNhan_1_data.json');
        // data = require('./mock/Force_Ranking_data.json');

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

exceljs.populate().catch(err => console.error(err));

function populate(worksheet, data) {
    populate_master_data(worksheet, data[0]);

    const group_temp = get_group_temp(worksheet);
    if (group_temp.length) {
        const row_pos = group_temp[0].row.number + group_temp.length - 1;
        populate_group_excel(worksheet, group_temp, row_pos, data);
        remove_temp_rows(worksheet, group_temp)
    } else {
        populate_table_excel(worksheet, data);
    }
    return;
}
function create_group_row(worksheet, row, pos, data) {
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
            // const feildID = cell.value.substring(cell.value.indexOf(':') + 1, cell.value.indexOf('}'));
            display_cell_values(cell, data, 'group');
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
}
function get_group_temp(worksheet) {
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
        const table_temp = get_table_temp(worksheet);
        if (table_temp.length) {
            group_rows_temp.push({ row: table_temp[0].row, GroupFeildID: [], level: 'child_row' });
        }
    }

    return group_rows_temp;
}
function get_table_temp(worksheet) {
    let table_temp = [];
    worksheet.eachRow({ includeEmpty: false }, row => {
        row.eachCell(c => {
            if (table_temp.length) {
                return;
            }

            if (c.value && typeof c.value === 'string' && c.value.includes('${table:')) {
                // ${table:GoalName}
                // table_temp = row;
                table_temp.push({ row })
            }
        });
    });

    return table_temp;
}
function create_table_row(worksheet, row, index, data) {

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
            // const feildID = cell.value.substring(cell.value.indexOf(':') + 1, cell.value.indexOf('}'));
            // cell.value = data[feildID] || '';
            display_cell_values(cell, data, 'table');
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

        let new_row = create_group_row(worksheet, rows_temps[group_level].row, current_pos, group_data_row);
        current_pos = new_row.number;

        if (rows_temps[group_level + 1] && rows_temps[group_level + 1].level === 'child_row') {
            for (const table_row_data of group_data[key]) {
                const table_row = create_table_row(worksheet, rows_temps[group_level + 1].row, current_pos, table_row_data);
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
            switch (c.type) {
                case 8: //RichText
                    if (c.value && c.value.richText && _.isArray(c.value.richText)) {
                        for (const iterator of c.value.richText) {
                            if (iterator.text && typeof iterator.text === 'string' && iterator.text.includes('${') && !iterator.text.includes('${group') && !iterator.text.includes('${table')) {
                                // iterator.text = 'iterator.text';
                                display_cell_values(iterator, data, 'master-richText');
                            }
                        }
                        break;
                    }
                    c.value = 0;
                    break;
                default:
                    if (c.value && typeof c.value === 'string' && c.value.includes('${') && !c.value.includes('${group') && !c.value.includes('${table')) {
                        //ex: ${EmployeeName}

                        // const array = c.value.split('$');
                        // for (let index = 1; index < array.length; index++) {
                        //     const feildID = c.value.substring(c.value.indexOf('{') + 1, c.value.indexOf('}'));

                        //     const re = new RegExp(`\\$\{${feildID}\}`, 'g');
                        //     const value = c.value.replace(re, data[feildID] || '');

                        display_cell_values(c, data, 'master');
                        // console.log(c.value)
                        // }
                    }
                    break;
            }
        });
    });
}
function display_cell_values(cell, data, key) {

    // Excel.ValueType
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
    // ------------------------------

    let reg_exp;
    let feildID;
    let array = (key === 'master-richText') ? cell.text.split('$') : cell.value.split('$');

    for (let index = 1; index < array.length; index++) {
        switch (key) {
            case 'master':
                feildID = cell.value.substring(cell.value.indexOf('{') + 1, cell.value.indexOf('}'));
                reg_exp = new RegExp(`\\$\{${feildID}\}`, 'g');
                cell.value = cell.value.replace(reg_exp, data[feildID] || '');
                break;

            case 'group':
                feildID = cell.value.substring(cell.value.indexOf(':') + 1, cell.value.indexOf('}'));
                let groupID = cell.value.substring(cell.value.indexOf('[') + 1, cell.value.indexOf(']'));
                // reg_exp = new RegExp(`\\$\{group\[${groupID}\]\:${feildID}\}`, 'g');
                // cell.value = cell.value.replace(reg_exp, data[feildID] || '');
                cell.value = cell.value.replace(`\$\{group\[${groupID}\]\:${feildID}\}`, data[feildID] || '')
                break;

            case 'table':
                feildID = cell.value.substring(cell.value.indexOf(':') + 1, cell.value.indexOf('}'));
                reg_exp = new RegExp(`\\$\{table:${feildID}\}`, 'g');
                cell.value = cell.value.replace(reg_exp, data[feildID] || '');
                break;

            case 'master-richText':
                feildID = cell.text.substring(cell.text.indexOf('{') + 1, cell.text.indexOf('}'));
                reg_exp = new RegExp(`\\$\{${feildID}\}`, 'g');
                cell.text = cell.text.replace(reg_exp, data[feildID] || '');

                continue;

            default:
                break;
        }

        switch (cell.numFmt) {
            case '0':
            case '0.0':
            case '0.00':
            case '0.000':
            case '0.0000':
            case '0%':
            case '0.0%':
            case '0.00%':
            case `0.00\\%`:
            case `0.0000\\%`:
                if (cell.value && !isNaN(cell.value)) {
                    cell.value = parseFloat(cell.value);
                    break;
                }
                cell.value = 0;
                break;
            default:
                break
        }
    }
    return;
}
function populate_table_excel(worksheet, data) {
    let pos;
    const table_temp = get_table_temp(worksheet);
    if (!_.isArray(table_temp) || !table_temp.length) {
        return;
    }

    for (const row_data of data) {
        const new_table_row = create_table_row(worksheet, table_temp[0].row, table_temp[0].row.number, row_data);
        pos = new_table_row.number;
    }

    // table_temp.hidden = true;
    remove_temp_rows(worksheet, table_temp)
    return pos;
}
function remove_temp_rows(worksheet, row_temps) {
    if (!_.isArray(row_temps) || !row_temps.length) {
        return;
    }

    let isMerged = false;
    const pos = row_temps[0].row.number;
    const rows = worksheet.getRows(pos, worksheet.lastRow.number - pos);

    for (let index = 0; index < rows.length; index++) {
        const row = rows[index];

        if (isMerged) {
            break;
        }
        row.eachCell(c => {
            if (isMerged) {
                return;
            }

            if (c.isMerged) {
                isMerged = true;
                return;
            }
        });
    }

    if (isMerged) {
        for (const rows_temp of row_temps) {
            rows_temp.row.eachCell(c => {
                c.value = '';
            });
            rows_temp.row.hidden = true;
            if (rows_temp.table_temp) {
                rows_temp.table_temp.eachCell(c => {
                    c.value = '';
                });
                rows_temp.table_temp.hidden = true;
            }
        }
    } else {
        worksheet.spliceRows(pos, row_temps.length);
    }
}
