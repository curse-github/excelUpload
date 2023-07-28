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
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var Colors = /** @class */ (function () {
    function Colors() {
    }
    Colors.Reset = "\x1b[0m";
    Colors.Bright = "\x1b[1m";
    Colors.Underscore = "\x1b[4m";
    Colors.Reverse = "\x1b[7m";
    //static Dim       :string = "\x1b[2m";//does not work at all
    //static Blink     :string = "\x1b[5m";//does not work at all
    //static Hidden    :string = "\x1b[8m";//does not work at all
    Colors.R = "\x1b[0m";
    Colors.B = "\x1b[1m";
    Colors.U = "\x1b[4m";
    Colors.Rev = "\x1b[7m";
    Colors.FgBlack = "\x1b[30m";
    Colors.FgRed = "\x1b[31m";
    Colors.FgGreen = "\x1b[32m";
    Colors.FgYellow = "\x1b[33m"; //does not work on powershell somehow
    Colors.FgBlue = "\x1b[34m";
    Colors.FgMagenta = "\x1b[35m";
    Colors.FgCyan = "\x1b[36m";
    Colors.FgWhite = "\x1b[37m";
    Colors.FgGray = "\x1b[90m";
    Colors.Fbla = "\x1b[30m";
    Colors.Fr = "\x1b[31m";
    Colors.Fgre = "\x1b[32m";
    Colors.Fy = "\x1b[33m"; //does not work on powershell somehow
    Colors.Fblu = "\x1b[34m";
    Colors.Fm = "\x1b[35m";
    Colors.Fc = "\x1b[36m";
    Colors.Fw = "\x1b[37m";
    Colors.Fgra = "\x1b[90m";
    Colors.BgBlack = "\x1b[40m";
    Colors.BgRed = "\x1b[41m";
    Colors.BgGreen = "\x1b[42m";
    Colors.BgYellow = "\x1b[43m";
    Colors.BgBlue = "\x1b[44m";
    Colors.BgMagenta = "\x1b[45m";
    Colors.BgCyan = "\x1b[46m";
    Colors.BgWhite = "\x1b[47m";
    Colors.BgGray = "\x1b[100m";
    Colors.Bbla = "\x1b[40m";
    Colors.Br = "\x1b[41m";
    Colors.Bgre = "\x1b[42m";
    Colors.By = "\x1b[43m";
    Colors.Bblu = "\x1b[44m";
    Colors.Bm = "\x1b[45m";
    Colors.Bc = "\x1b[46m";
    Colors.Bw = "\x1b[47m";
    Colors.Bgra = "\x1b[100m";
    return Colors;
}());
var exceljs_1 = require("exceljs");
var fs = require("fs");
var config = JSON.parse(fs.readFileSync("./config.json").toString());
var filename = config.filename;
var sheetName = config.sheetName;
var transferHandler_1 = require("./transferHandler");
function run() {
    return __awaiter(this, void 0, void 0, function () {
        var start, book, worksheet, out, Continue, headerRow, headers, i, numColumns, numEmpty, i, row, textRow, j, value, tmp, time;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    console.clear();
                    console.log(Colors.Fgra + "ExcelJs" + Colors.R);
                    start = (new Date()).getTime();
                    book = new exceljs_1.Workbook();
                    console.log(Colors.Fgra + "Reading file..." + Colors.R);
                    return [4 /*yield*/, book.xlsx.readFile(__dirname + "/" + filename + ".xlsx")];
                case 1:
                    book = _a.sent();
                    worksheet = book.getWorksheet(sheetName);
                    out = [];
                    try {
                        Continue = true;
                        headerRow = worksheet.getRow(1);
                        headers = [];
                        for (i = 1; headerRow.getCell(i).text != ""; i++)
                            headers.push(headerRow.getCell(i).text);
                        numColumns = headers.length;
                        numEmpty = 0;
                        for (i = 1; Continue; i++) {
                            row = worksheet.getRow(i);
                            textRow = [];
                            for (j = 1; j < numColumns + 1; j++) {
                                value = row.getCell(j).text;
                                if (value.match(/^\S\S\S \S\S\S \d\d \d\d\d\d \d\d:\d\d:\d\d GMT((\+|\-)\d\d\d\d)? \(.*\)$/g) != null) {
                                    tmp = (new Date(value)).valueOf();
                                    if (!Number.isNaN(tmp))
                                        value = (new Date(tmp + 1000 * 60 * 60 * 24)).toLocaleDateString();
                                }
                                textRow.push(value);
                            }
                            row.destroy();
                            if (textRow.length == 0 || textRow[0] == "") {
                                if (numEmpty == 9) {
                                    Continue = false;
                                    break;
                                }
                                numEmpty++;
                                continue;
                            }
                            else
                                numEmpty = 0;
                            out.push(textRow);
                        }
                    }
                    catch (error) { }
                    worksheet.destroy();
                    time = (new Date()).getTime() - start;
                    (0, transferHandler_1.handle)(out, __dirname + "/" + filename + ".json", time);
                    return [2 /*return*/];
            }
        });
    });
}
run().catch(function (err) { console.log(err); });
