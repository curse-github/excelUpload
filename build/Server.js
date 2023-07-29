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
//#region imports/libs
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
function generateUUID() {
    var a = (new Date()).getTime(); //Timestamp
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var b = Math.random() * 16; //random number between 0 and 16
        b = (a + b) % 16 | 0;
        a = Math.floor(a / 16);
        return (c === 'x' ? b : (b & 0x3 | 0x8)).toString(16);
    });
}
function sha256(input) {
    return (require('crypto').createHash("sha256").update(input).digest("hex"));
}
var https = require("https");
var express = require("express");
var fs = require("fs");
//#endregion imports/libs
console.clear();
//#region Microsoft Graph
var clientId = fs.readFileSync(__dirname + "/graphTokens/ClientId.txt").toString();
var clientSecretValue = fs.readFileSync(__dirname + "/graphTokens/ClientSecret.txt").toString();
var scope = encodeURIComponent("User.Read offline_access Files.ReadWrite");
var redirect = "http://localhost:80/oauth/"; // https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/06e160b0-cc2e-4181-bf03-7aa31e0bf435/isMSAApp/true
var authUrl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?";
authUrl += "client_id=" + clientId + "&response_type=code";
authUrl += "&redirect_uri=" + encodeURIComponent(redirect);
authUrl += "&response_mode=query" + "&scope=" + scope + "&state=123456789";
var app = express();
app.use(express.urlencoded({ extended: true }));
app.use(require("body-parser").text());
var pages = {
    "/": function (req, res, send) { return res.redirect("/index.html"); },
    "/index.html": function (req, res, send) { return send(); },
    "/login": function (req, res, send) { return res.redirect(authUrl); },
    "/logout": function (req, res, send) { return send("/logout.html"); },
    "/oauth": function (req, res, send) {
        if (req.query.code != null)
            requestToken(req.query.code);
        res.redirect("/index.html");
    }
};
Object.keys(pages).forEach(function (i) {
    app.get(i, function (req, res) {
        pages[i](req, res, (function (page) { res.sendFile(__dirname + "/webpage" + (page != null ? page : i), "utf8"); }));
    });
});
app.listen(80, function () {
    console.clear();
    console.log(Colors.Fgra + "Microsoft Graph server is running at: " + Colors.Fgre + "http://localhost:80" + Colors.R);
});
var refreshTimeout;
// https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
function requestToken(code) {
    return __awaiter(this, void 0, void 0, function () {
        var params;
        return __generator(this, function (_a) {
            params = "client_id=" + clientId + "&scope=" + scope;
            params += "&code=" + code + "&redirect_uri=" + encodeURIComponent(redirect);
            params += "&grant_type=authorization_code&client_secret=" + clientSecretValue;
            httpsRequestPromise("login.microsoftonline.com", "/consumers/oauth2/v2.0/token", "POST", { "Content-Type": "application/x-www-form-urlencoded" }, params)
                .then(function (out) {
                var json = JSON.parse(out);
                fs.writeFileSync(__dirname + "/graphTokens/access_token.txt", json.access_token);
                fs.writeFileSync(__dirname + "/graphTokens/refresh_token.txt", json.refresh_token);
                console.log(Colors.Fgra + "Access token saved to file." + Colors.R);
                console.log(Colors.Fgra + "Refresh token saved to file.\n" + Colors.R);
                if (refreshTimeout != undefined)
                    clearTimeout(refreshTimeout);
                refreshTimeout = setTimeout(function () {
                    refreshToken();
                    refreshTimeout = undefined;
                }, 1000 * 60 * 60); //one hour
            });
            return [2 /*return*/];
        });
    });
}
function refreshToken() {
    return __awaiter(this, void 0, void 0, function () {
        var params;
        return __generator(this, function (_a) {
            params = "client_id=" + clientId + "&scope=" + scope;
            params += "&refresh_token=" + (fs.readFileSync(__dirname + "/graphTokens/refresh_token.txt").toString());
            params += "&grant_type=refresh_token&client_secret=" + clientSecretValue;
            httpsRequestPromise("login.microsoftonline.com", "/consumers/oauth2/v2.0/token", "GET", { "Content-Type": "application/x-www-form-urlencoded" }, params)
                .then(function (out) {
                var json = JSON.parse(out);
                fs.writeFileSync(__dirname + "/graphTokens/access_token.txt", json.access_token);
                fs.writeFileSync(__dirname + "/graphTokens/refresh_token.txt", json.refresh_token);
                console.log(Colors.Fgra + "\"Refreshed\" access token saved to file." + Colors.R);
                console.log(Colors.Fgra + "\"Refreshed\" refresh token saved to file.\n" + Colors.R);
                if (refreshTimeout != undefined)
                    clearTimeout(refreshTimeout);
                refreshTimeout = setTimeout(function () {
                    refreshToken();
                    refreshTimeout = undefined;
                }, 1000 * 60 * 60); //one hour
            });
            return [2 /*return*/];
        });
    });
}
//#endregion Microsoft Graph
//#region Excel Add-in
var addInApp = express();
addInApp.use(express.urlencoded({ extended: true }));
addInApp.use(require("body-parser").text());
var addInPages = {
    "/assets/logo-filled.png": function (req, res, send) { return send(); },
    "/assets/icon-16.png": function (req, res, send) { return send(); },
    "/assets/icon-32.png": function (req, res, send) { return send(); },
    "/assets/icon-64.png": function (req, res, send) { return send(); },
    "/assets/icon-80.png": function (req, res, send) { return send(); },
    "/assets/icon-128.png": function (req, res, send) { return send(); },
    "/d916c354314f23a46dd2.css": function (req, res, send) { return send(); },
    "/manifest.xml": function (req, res, send) { return send(); },
    "/commands.js.map": function (req, res, send) { return send(); },
    "/commands.html": function (req, res, send) { return send(); },
    "/commands.js": function (req, res, send) { return send(); },
    "/polyfill.js.map": function (req, res, send) { return send(); },
    "/polyfill.js": function (req, res, send) { return send(); },
    "/taskpane.js.map": function (req, res, send) { return send(); },
    "/taskpane.html": function (req, res, send) { return send(); },
    "/taskpane.js": function (req, res, send) { return send(); },
    "/taskpane": function (req, res, send) { return send("/taskpane.html"); }
};
Object.keys(addInPages).forEach(function (i) {
    addInApp.get(i, function (req, res) {
        addInPages[i](req, res, (function (page) { res.sendFile(__dirname + "/ExcelAddIn/plugin" + (page != null ? page : i), "utf8"); }));
    });
});
var transferHandler_1 = require("./transferHandler");
var filename = JSON.parse(fs.readFileSync(__dirname + "/config.json").toString()).filename;
var transfers = {};
addInApp.post("/wholeTransfer", function (req, res) {
    res.send("{}");
    var body = JSON.parse(req.body);
    var start = parseInt(body.start);
    var time = (new Date()).getTime() - start;
    (0, transferHandler_1.handle)(body.data, __dirname + "/" + filename + ".json", time);
});
addInApp.post("/startTransfer", function (req, res) {
    var id = generateUUID();
    var body = JSON.parse(req.body);
    var sha = body.sha;
    var start = body.start;
    if (sha != null) {
        res.send("{\"id\":\"" + id + "\"}");
        transfers[id] = {
            data: [],
            "sha": sha,
            "start": start
        };
    }
});
addInApp.post("/transfer", function (req, res) {
    var _a;
    res.send("{}");
    var body = JSON.parse(req.body);
    var id = req.query.id;
    if (id != null && transfers[id] != null) {
        (_a = transfers[id].data).push.apply(_a, body);
    }
});
addInApp.post("/endTransfer", function (req, res) {
    var id = req.query.id;
    if (id != null && transfers[id] != null) {
        var transfer = transfers[id].data;
        var sha = sha256(JSON.stringify(transfer));
        var status_1 = (sha == transfers[id].sha);
        var time = (new Date()).getTime() - transfers[id].start;
        delete transfers[id];
        if (status_1)
            (0, transferHandler_1.handle)(transfer, __dirname + "/" + filename + ".json", time);
        res.send("{\"status\":" + status_1 + "}");
    }
    else {
        res.send("{\"status\":false}");
    }
});
addInApp.listen(3000, function () {
    console.log(Colors.Fgra + "Excel Add-in server is running at: " + Colors.Fgre + "http://localhost:3000" + Colors.R);
});
//#endregion Excel Add-in
