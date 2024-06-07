"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.downloadXlsx = exports.generateXlsx = void 0;
const helpers_1 = require("./helpers");
const main_1 = __importDefault(require("./main"));
/**
 * Generates an XLSX file by replacing data in a template file.
 * @throws {Error} Throws an error if either the templateURL or data is not provided, or if unable to fetch the template file.
 */
function generateXlsx(templateURL, data) {
    return __awaiter(this, void 0, void 0, function* () {
        if (!templateURL || !data)
            throw new Error('No templateURL or data provided');
        return yield fetch(templateURL)
            .then(response => {
            if (response.ok)
                return response.arrayBuffer();
            throw new Error(`Unable to fetch the template at: ${templateURL}`);
        })
            .then(template => (0, main_1.default)(template, data))
            .catch(e => console.error(e));
    });
}
exports.generateXlsx = generateXlsx;
/**
 * Downloads an XLSX file by replacing data in a template file and initiates the download in the browser.
 */
function downloadXlsx(templateURL, data, fileName) {
    return __awaiter(this, void 0, void 0, function* () {
        fileName = fileName || `${new Date().toISOString().substring(0, 10)} - Report.xlsx`;
        const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        const xlsxBlob = yield generateXlsx(templateURL, data);
        if (xlsxBlob)
            (0, helpers_1.downloadBlob)(xlsxBlob, fileName, mimeType);
    });
}
exports.downloadXlsx = downloadXlsx;
