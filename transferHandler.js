"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.handle = void 0;
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
function stringify(json) {
    return "[\n    " + json.map(function (el) { return "[" + el.map(function (el) { return "\"" + el.split("\"").join("\\\"") + "\""; }).join(",") + "]"; }).join(",\n    ") + "\n]";
}
function sha256(input) {
    return (require('crypto').createHash("sha256").update(input).digest("hex"));
}
var fs = require("fs");
function handle(transfer, outfile, time) {
    console.log(Colors.Fgra + "Received " + Colors.Fy + transfer.length + Colors.Fgra + " lines." + Colors.R);
    console.log(Colors.Fgra + "Sha256: " + Colors.Fc + sha256(JSON.stringify(transfer)) + Colors.R);
    fs.writeFile(outfile, stringify(transfer), function (e) { if (e) {
        console.log(e);
    } });
    console.log(Colors.Fgra + "Time: " + Colors.Fy + (time / 1000) + Colors.Fgra + " Seconds" + Colors.R);
}
exports.handle = handle;
