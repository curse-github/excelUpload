class Colors {
    static Reset     :string = "\x1b[0m";
    static Bright    :string = "\x1b[1m";
    static Underscore:string = "\x1b[4m";
    static Reverse   :string = "\x1b[7m";
    //static Dim       :string = "\x1b[2m";//does not work at all
    //static Blink     :string = "\x1b[5m";//does not work at all
    //static Hidden    :string = "\x1b[8m";//does not work at all
    static R  :string = "\x1b[0m";
    static B  :string = "\x1b[1m";
    static U  :string = "\x1b[4m";
    static Rev:string = "\x1b[7m";

    static FgBlack  :string = "\x1b[30m";
    static FgRed    :string = "\x1b[31m";
    static FgGreen  :string = "\x1b[32m";
    static FgYellow :string = "\x1b[33m";//does not work on powershell somehow
    static FgBlue   :string = "\x1b[34m";
    static FgMagenta:string = "\x1b[35m";
    static FgCyan   :string = "\x1b[36m";
    static FgWhite  :string = "\x1b[37m";
    static FgGray   :string = "\x1b[90m";
    static Fbla:string = "\x1b[30m";
    static Fr  :string = "\x1b[31m";
    static Fgre:string = "\x1b[32m";
    static Fy  :string = "\x1b[33m";//does not work on powershell somehow
    static Fblu:string = "\x1b[34m";
    static Fm  :string = "\x1b[35m";
    static Fc  :string = "\x1b[36m";
    static Fw  :string = "\x1b[37m";
    static Fgra:string = "\x1b[90m";

    static BgBlack  :string = "\x1b[40m" ;
    static BgRed    :string = "\x1b[41m" ;
    static BgGreen  :string = "\x1b[42m" ;
    static BgYellow :string = "\x1b[43m" ;
    static BgBlue   :string = "\x1b[44m" ;
    static BgMagenta:string = "\x1b[45m" ;
    static BgCyan   :string = "\x1b[46m" ;
    static BgWhite  :string = "\x1b[47m" ;
    static BgGray   :string = "\x1b[100m";
    static Bbla:string = "\x1b[40m" ;
    static Br  :string = "\x1b[41m" ;
    static Bgre:string = "\x1b[42m" ;
    static By  :string = "\x1b[43m" ;
    static Bblu:string = "\x1b[44m" ;
    static Bm  :string = "\x1b[45m" ;
    static Bc  :string = "\x1b[46m" ;
    static Bw  :string = "\x1b[47m" ;
    static Bgra:string = "\x1b[100m";
}

import {Workbook,Worksheet,Row} from "exceljs";
import * as fs from "fs";
const config:any = JSON.parse(fs.readFileSync("./config.json").toString())
const filename:string = config.filename;
const sheetName:string = config.sheetName;

import {handle} from "./transferHandler";
async function run() {
    console.clear();
    console.log(Colors.Fgra+"ExcelJs"+Colors.R);
    const start:number = (new Date()).getTime();
    var book:Workbook = new Workbook();
    console.log(Colors.Fgra+"Reading file..."+Colors.R);
    book = await book.xlsx.readFile(__dirname+"/"+filename+".xlsx");
    var worksheet:Worksheet = book.getWorksheet(sheetName);

    var out:string[][] = [];
    try {
        var Continue:boolean = true;
        const headerRow:Row = worksheet.getRow(1);
        const headers:string[] = [];
        for (let i = 1; headerRow.getCell(i).text!=""; i++) headers.push(headerRow.getCell(i).text);
        const numColumns:number = headers.length;
        var numEmpty:number = 0;
        for (let i = 1; Continue; i++) {
            const row:Row = worksheet.getRow(i);
            const textRow:string[] = [];
            for (let j = 1; j < numColumns+1; j++) {
                var value:any = row.getCell(j).text;
                if (value.match(/^\S\S\S \S\S\S \d\d \d\d\d\d \d\d:\d\d:\d\d GMT((\+|\-)\d\d\d\d)? \(.*\)$/g)!=null) {
                    const tmp:number = (new Date(value)).valueOf();
                    if (!Number.isNaN(tmp)) value=(new Date(tmp+1000*60*60*24)).toLocaleDateString();
                }
                textRow.push(value);
            }
            row.destroy();
            if (textRow.length==0||textRow[0]=="") {
                if (numEmpty==9) { Continue=false;break; }
                numEmpty++;continue;
            } else numEmpty=0;
            out.push(textRow);
        }
    } catch (error) {}
    worksheet.destroy();
    const time:number = (new Date()).getTime()-start;
    handle(out,__dirname+"/"+filename+".json",time);
}
run().catch((err:any)=>{console.log(err);});