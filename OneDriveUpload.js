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
exports.httpsRequestPromise = void 0;
//#region imports
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
/** Requests or posts from a https server and returns output.
 * @param link Link to the server to request from.
 * @param path Filepath of the file you are accessing from the server.
 * @param method Request method e.g "GET","POST","PUT".
 * @param headers An object of headers.
 * @param data Data to send if method is post.
 * @returns A Promise which will return the data returned from the request. */
function httpsRequestPromise(link, path, method, headers, data) {
    return new Promise(function (resolve, reject) {
        try {
            var curl = "";
            if (data != null) {
                //@ts-expect-error
                headers["Content-Length"] = data.length;
            }
            var options = {
                "host": link,
                "port": 443,
                "path": path,
                "method": method,
                "headers": headers
            };
            var req = https.request(options, function (res) {
                res.setEncoding('utf8');
                res.on("data", function (chunk) { curl += chunk; });
                res.on("close", function () { resolve(curl); });
            });
            req.on("error", function (err) { reject(err); });
            if (data != null) {
                req.write(data);
            }
            req.end();
        }
        catch (err) {
            reject(err);
        }
    });
}
exports.httpsRequestPromise = httpsRequestPromise;
function round3(num) { return Math.round(num * 1000) / 1000; }
var microsoft_graph_client_1 = require("@microsoft/microsoft-graph-client");
require("isomorphic-fetch");
var https = require("https");
var fs = require("fs");
var config = JSON.parse(fs.readFileSync("./config.json").toString());
var filename = config.filename;
var sheetName = config.sheetName;
var left = config.left;
var right = config.right;
//#endregion imports
var clientId = "06e160b0-cc2e-4181-bf03-7aa31e0bf435";
var clientSecretValue = "VMG8Q~PQf2zyE-OI_YlDOtLxy22dDDti8RxEZaWL";
var scope = encodeURIComponent("User.Read offline_access Files.ReadWrite");
var MicrosoftGraph = /** @class */ (function () {
    function MicrosoftGraph() {
        var _this = this;
        this.accessToken = fs.readFileSync(__dirname + "/graphTokens/access_token.txt").toString(); // https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
        this.client = microsoft_graph_client_1.Client.init({
            authProvider: (function (done) { return done(null, _this.accessToken); })
        });
    }
    MicrosoftGraph.prototype.setClient = function () {
        var _this = this;
        this.accessToken = fs.readFileSync(__dirname + "/graphTokens/access_token.txt").toString(); // https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
        this.client = microsoft_graph_client_1.Client.init({
            authProvider: (function (done) { return done(null, _this.accessToken); })
        });
    };
    MicrosoftGraph.refreshToken = function () {
        return __awaiter(this, void 0, void 0, function () {
            var params, json, _a, _b;
            return __generator(this, function (_c) {
                switch (_c.label) {
                    case 0:
                        params = "client_id=" + clientId + "&scope=" + scope;
                        params += "&refresh_token=" + (fs.readFileSync(__dirname + "/graphTokens/refresh_token.txt").toString());
                        params += "&grant_type=refresh_token&client_secret=" + clientSecretValue;
                        _b = (_a = JSON).parse;
                        return [4 /*yield*/, httpsRequestPromise("login.microsoftonline.com", "/consumers/oauth2/v2.0/token", "GET", { "Content-Type": "application/x-www-form-urlencoded" }, params)];
                    case 1:
                        json = _b.apply(_a, [_c.sent()]);
                        fs.writeFileSync(__dirname + "/graphTokens/access_token.txt", json.access_token);
                        fs.writeFileSync(__dirname + "/graphTokens/refresh_token.txt", json.refresh_token);
                        console.log(Colors.Fgra + "Access token saved to file." + Colors.R);
                        console.log(Colors.Fgra + "Refresh token saved to file.\n" + Colors.R);
                        return [2 /*return*/];
                }
            });
        });
    };
    MicrosoftGraph.prototype.apiGet = function (request) {
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2 /*return*/, new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                        var out, err_1, out;
                        return __generator(this, function (_a) {
                            switch (_a.label) {
                                case 0:
                                    _a.trys.push([0, 2, , 6]);
                                    return [4 /*yield*/, this.client.api(request).get()];
                                case 1:
                                    out = _a.sent();
                                    resolve(out);
                                    return [2 /*return*/];
                                case 2:
                                    err_1 = _a.sent();
                                    if (!(err_1.code == "InvalidAuthenticationToken")) return [3 /*break*/, 4];
                                    return [4 /*yield*/, MicrosoftGraph.refreshToken()];
                                case 3:
                                    _a.sent();
                                    this.setClient();
                                    out = this.client.api(request).get();
                                    resolve(out);
                                    return [2 /*return*/];
                                case 4:
                                    console.log(err_1.code);
                                    console.error(err_1);
                                    _a.label = 5;
                                case 5: return [3 /*break*/, 6];
                                case 6: return [2 /*return*/];
                            }
                        });
                    }); })];
            });
        });
    };
    MicrosoftGraph.prototype.getUser = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me")];
                    case 1: return [2 /*return*/, (_a.sent())];
                }
            });
        });
    };
    // oneDrive: https://onedrive.live.com/?id=root&cid=148A3EA57177D1CA
    // graph list children: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http
    MicrosoftGraph.prototype.getRoot = function () {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me/drive/root/children")];
                    case 1: return [2 /*return*/, (_a.sent()).value];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getPath = function (path) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me/drive/root" + (":" + path + ":") + "/children")];
                    case 1: return [2 /*return*/, (_a.sent()).value];
                }
            });
        });
    };
    // https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0
    MicrosoftGraph.prototype.printDir = function (dir) {
        return __awaiter(this, void 0, void 0, function () {
            var i, child;
            return __generator(this, function (_a) {
                console.log(Colors.Fgre + "/root" + Colors.R);
                for (i = 0; i < dir.length; i++) {
                    child = dir[i];
                    if (child.file == null) { //is directory
                        console.log(Colors.Fgre + "    /" + child.name + Colors.R);
                    }
                    else { //is directory
                        console.log(Colors.Fgra + "    \"" + child.name + "\": " + Colors.Fy + child.id + Colors.R);
                    }
                }
                return [2 /*return*/];
            });
        });
    };
    MicrosoftGraph.prototype.findFile = function (dir, filename) {
        return __awaiter(this, void 0, void 0, function () {
            var i;
            return __generator(this, function (_a) {
                for (i = 0; i < dir.length; i++) {
                    if (dir[i].name == filename)
                        return [2 /*return*/, dir[i]];
                }
                return [2 /*return*/];
            });
        });
    };
    MicrosoftGraph.prototype.getExcelUsedRange = function (id, sheet) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me/drive/items/" + id + "/workbook/worksheets/" + sheet + "/usedRange")];
                    case 1: // https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
                    return [2 /*return*/, (_a.sent())];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelRange = function (id, sheet) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me/drive/items/" + id + "/workbook/worksheets/" + sheet + "/range")];
                    case 1: // https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
                    return [2 /*return*/, (_a.sent())];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelRangeByAddress = function (id, sheet, range) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (range.match(/^[A-Z]+\d+:[A-Z]+\d+$/g) == null)
                            return [2 /*return*/, null];
                        return [4 /*yield*/, this.apiGet("/me/drive/items/" + id + "/workbook/worksheets/" + sheet + "/range(address=\'" + range + "\')")];
                    case 1: return [2 /*return*/, (_a.sent())];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelNames = function (id) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me/drive/items/" + id + "/workbook/names")];
                    case 1: return [2 /*return*/, (_a.sent()).value];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelRangeByName = function (id, name) {
        return __awaiter(this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0: return [4 /*yield*/, this.apiGet("/me/drive/items/" + id + "/workbook/names/" + name + "/range")];
                    case 1: // https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
                    return [2 /*return*/, (_a.sent())];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelRangeLarge = function (id, sheet, range, stepSize) {
        return __awaiter(this, void 0, void 0, function () {
            var left, top, right, bottom, out, promises, _loop_1, i;
            var _this = this;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (range.match(/^[A-Z]+\d+:[A-Z]+\d+$/g) == null)
                            return [2 /*return*/, []];
                        left = "" + range.match(/^[A-Z]+(?=\d+:[A-Z]+\d+$)/g);
                        top = parseInt("" + range.match(/(?<=^[A-Z]+)\d+(?=:[A-Z]+\d+$)/g));
                        right = "" + range.match(/(?<=^[A-Z]+\d+:)[A-Z]+(?=\d+$)/g);
                        bottom = parseInt("" + range.match(/(?<=^[A-Z]+\d+:[A-Z]+)\d+$/g));
                        if (top - bottom > 500000)
                            return [2 /*return*/, []];
                        out = [];
                        promises = [];
                        _loop_1 = function (i) {
                            promises.push(new Promise(function (resolve) { return __awaiter(_this, void 0, void 0, function () {
                                var range, j;
                                return __generator(this, function (_a) {
                                    switch (_a.label) {
                                        case 0: return [4 /*yield*/, this.getExcelRangeByAddress(id, sheet, left + i + ":" + right + Math.min((i + stepSize), bottom))];
                                        case 1:
                                            range = (_a.sent()).text;
                                            for (j = 0; j < range.length; j++) {
                                                out[i - top + j] = range[j];
                                            }
                                            resolve();
                                            return [2 /*return*/];
                                    }
                                });
                            }); }));
                        };
                        for (i = top; i < bottom; i += stepSize) {
                            _loop_1(i);
                        }
                        return [4 /*yield*/, Promise.all(promises)];
                    case 1:
                        _a.sent();
                        return [2 /*return*/, out];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelColumns = function (id, sheet, left, right, stepSize) {
        return __awaiter(this, void 0, void 0, function () {
            var start, Continue, out, i, range, numEmpty, j, time;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        start = (new Date()).getTime();
                        Continue = true;
                        out = [];
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!Continue) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getExcelRangeByAddress(id, sheet, left + (i * stepSize + 1) + ":" + right + ((i + 1) * stepSize))];
                    case 2:
                        range = (_a.sent()).text;
                        numEmpty = 0;
                        for (j = 0; j < range.length; j++) {
                            if (range[j][0] == "") {
                                if (numEmpty == 9) {
                                    Continue = false;
                                    break;
                                }
                                numEmpty++;
                                continue;
                            }
                            else
                                numEmpty = 0;
                            out.push(range[j]);
                        }
                        time = ((new Date()).getTime() - start);
                        console.log(out.length + ": " + round3(time / 1000) + "s");
                        _a.label = 3;
                    case 3:
                        i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/, out];
                }
            });
        });
    };
    MicrosoftGraph.prototype.getExcelColumnsLarge = function (id, sheet, left, right, stepSize, numSteps) {
        return __awaiter(this, void 0, void 0, function () {
            var start, Continue, out, i, range, numEmpty, j;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        start = (new Date()).getTime();
                        Continue = true;
                        out = [];
                        i = 0;
                        _a.label = 1;
                    case 1:
                        if (!Continue) return [3 /*break*/, 4];
                        return [4 /*yield*/, this.getExcelRangeLarge(id, sheet, left + (i * stepSize * numSteps + 1) + ":" + right + ((i + 1) * stepSize * numSteps), stepSize)];
                    case 2:
                        range = (_a.sent());
                        numEmpty = 0;
                        for (j = 0; j < range.length; j++) {
                            if (range[j][0] == "") {
                                if (numEmpty == 9) {
                                    Continue = false;
                                    break;
                                }
                                numEmpty++;
                                continue;
                            }
                            else
                                numEmpty = 0;
                            out.push(range[j]);
                        }
                        _a.label = 3;
                    case 3:
                        i++;
                        return [3 /*break*/, 1];
                    case 4: return [2 /*return*/, out];
                }
            });
        });
    };
    return MicrosoftGraph;
}());
var transferHandler_1 = require("./transferHandler");
function run() {
    return __awaiter(this, void 0, void 0, function () {
        var start, client, file, _a, _b, range, time;
        return __generator(this, function (_c) {
            switch (_c.label) {
                case 0:
                    console.clear();
                    console.log(Colors.Fgra + "Microsoft graph" + Colors.R);
                    start = (new Date()).getTime();
                    client = new MicrosoftGraph();
                    _b = (_a = client).findFile;
                    return [4 /*yield*/, client.getRoot()];
                case 1: return [4 /*yield*/, _b.apply(_a, [_c.sent(), filename + ".xlsx"])];
                case 2:
                    file = _c.sent();
                    return [4 /*yield*/, client.getExcelColumnsLarge(file.id, sheetName, left, right, 10000, 15)];
                case 3:
                    range = _c.sent();
                    time = (new Date()).getTime() - start;
                    (0, transferHandler_1.handle)(range, __dirname + "/" + filename + ".json", time);
                    return [2 /*return*/];
            }
        });
    });
}
run().catch(function (err) { console.log(err); });
