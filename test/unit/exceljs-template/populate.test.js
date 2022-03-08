/**
 * @copyright
 * @author ngochuy
 * @create 2022-02-28
 * @update 2022-02-28
 * @Description
 */
'use strict';

const fs = require('fs');
const exceljsTemplate = require('../../../src/index');
const ExceljsTemplateError = require('../../../src/ExceljsTemplateError');

describe('exceljs-template populate', () => {
    it('expects excel template to be Buffer', async () => {
        try {
            await exceljsTemplate.populate('non-buffer', []);
            expect('here').not.toBe('here');
        } catch (e) {
            expect(e instanceof ExceljsTemplateError).toBe(true);
            expect(e.type).toBe(ExceljsTemplateError.TYPE_INPUT);
            expect(e.message).toMatchSnapshot();
        }
    });
    it('expects data to be Array', async () => {
        try {
            await exceljsTemplate.populate(Buffer.from(''), 'non-array');
            expect('here').not.toBe('here');
        } catch (e) {
            expect(e instanceof ExceljsTemplateError).toBe(true);
            expect(e.type).toBe(ExceljsTemplateError.TYPE_INPUT);
            expect(e.message).toMatchSnapshot();
        }
    });
    it('populate a group excel template', async () => {

        const template = fs.readFileSync(`${__dirname}/../../../resources/ManagementByObjectives_MBO_Report_Tempalte.xlsx`);
        const data = require('../../../resources/ManagementByObjectives_MBO_Report.json');

        const populated = await exceljsTemplate.populate(template, data);
        expect(populated instanceof Buffer).toBe(true);
        fs.writeFileSync('ManagementByObjectives_MBO_Report_' + Date.now() + '.xlsx', populated);
    });
    it('populate a table excel template', async () => {

        const template = fs.readFileSync(`${__dirname}/../../../resources/Force_Ranking_Tempalte.xlsx`);
        const data = require('../../../resources/Force_Ranking_data.json');

        const populated = await exceljsTemplate.populate(template, data);
        expect(populated instanceof Buffer).toBe(true);
        fs.writeFileSync('Force_Ranking_Tempalte_' + Date.now() + '.xlsx', populated);
    });
});
