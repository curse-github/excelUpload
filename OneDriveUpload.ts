//#region imports
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
/** Requests or posts from a https server and returns output.
 * @param link Link to the server to request from.
 * @param path Filepath of the file you are accessing from the server.
 * @param method Request method e.g "GET","POST","PUT".
 * @param headers An object of headers.
 * @param data Data to send if method is post.
 * @returns A Promise which will return the data returned from the request. */
export function httpsRequestPromise(link:string,path:string,method:string,headers:{[key:string]:string},data:null|string):Promise<string> {
    return new Promise<string>((resolve,reject) => {
        try {
            var curl:string = "";
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
            var req = https.request(options, (res:any) => {
                res.setEncoding('utf8');
                res.on("data", (chunk:string) => { curl += chunk; });
                res.on("close", () => { resolve(curl); });
            });
            req.on("error", (err:Error) => { reject(err); });
            if (data != null) { req.write(data); }
            req.end();
        } catch (err:any) { reject(err); }
    });
}
function round3(num:number) { return Math.round(num*1000)/1000; }

import { Client,AuthProvider,AuthProviderCallback } from "@microsoft/microsoft-graph-client";
require("isomorphic-fetch");
import * as https from "https";
import * as fs from "fs";
const config:any = JSON.parse(fs.readFileSync("./config.json").toString())
const filename:string = config.filename;
const sheetName:string = config.sheetName;
const left:string = config.left;
const right:string = config.right;
//#endregion imports

const clientId:string = "06e160b0-cc2e-4181-bf03-7aa31e0bf435";
const clientSecretValue:string = "VMG8Q~PQf2zyE-OI_YlDOtLxy22dDDti8RxEZaWL";
const scope:string = encodeURIComponent("User.Read offline_access Files.ReadWrite");

type letter = 'A'|'B'|'C'|'D'|'E'|'F'|'G'|'H'|'I'|'J'|'K'|'L'|'M'|'N'|'O'|'P'|'Q'|'R'|'S'|'T'|'U'|'V'|'W'|'X'|'Y'|'Z';
class MicrosoftGraph {// https://learn.microsoft.com/en-us/graph/auth-register-app-v2
    accessToken:string
    client:Client;
    constructor() {
        this.accessToken=fs.readFileSync(__dirname+"/graphTokens/access_token.txt").toString();// https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
        this.client = Client.init({// Initialize the Microsoft Graph client
            authProvider: ((done:AuthProviderCallback)=>done(null, this.accessToken)) as AuthProvider
        });
    }
    setClient() {
        this.accessToken=fs.readFileSync(__dirname+"/graphTokens/access_token.txt").toString();// https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
        this.client = Client.init({// Initialize the Microsoft Graph client
            authProvider: ((done:AuthProviderCallback)=>done(null, this.accessToken)) as AuthProvider
        });
    }
    static async refreshToken():Promise<void> {
        var params:string = "client_id="+clientId+"&scope="+scope;
        params+="&refresh_token="+(fs.readFileSync(__dirname+"/graphTokens/refresh_token.txt").toString());
        params+="&grant_type=refresh_token&client_secret="+clientSecretValue;
        const json:any = JSON.parse(await httpsRequestPromise("login.microsoftonline.com","/consumers/oauth2/v2.0/token","GET",{"Content-Type":"application/x-www-form-urlencoded"},params));
        fs.writeFileSync(__dirname+"/graphTokens/access_token.txt",json.access_token);
        fs.writeFileSync(__dirname+"/graphTokens/refresh_token.txt",json.refresh_token);
        console.log(Colors.Fgra+"Access token saved to file."+Colors.R);
        console.log(Colors.Fgra+"Refresh token saved to file.\n"+Colors.R);
        return;
    }
    async apiGet(request:string):Promise<any> {
        return new Promise<any>(async(resolve:(value:any)=>void)=>{
            try {
                const out:any = await this.client.api(request).get();
                resolve(out);return;
            } catch (err:any) {
                if (err.code=="InvalidAuthenticationToken") {
                    await MicrosoftGraph.refreshToken();
                    this.setClient();
                    const out:any = this.client.api(request).get();
                    resolve(out);return;
                } else {
                    console.log(err.code)
                    console.error(err)
                }
            }
        });
    }
    async getUser() {
        return (await this.apiGet("/me"));
    }
    // oneDrive: https://onedrive.live.com/?id=root&cid=148A3EA57177D1CA
    // graph list children: https://learn.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http
    async getRoot():Promise<any[]> {
        return (await this.apiGet("/me/drive/root/children")).value;
    }
    async getPath(path:string):Promise<any[]> {
        return (await this.apiGet("/me/drive/root"+(":"+path+":")+"/children")).value;
    }
    // https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0
    async printDir(dir:any[]):Promise<void> {
        console.log(Colors.Fgre+"/root"+Colors.R);
        for (let i = 0; i < dir.length; i++) {
            const child:any = dir[i];
            if (child.file==null) {//is directory
                console.log(Colors.Fgre+"    /"+child.name+Colors.R);
            } else {//is directory
                console.log(Colors.Fgra+"    \""+child.name+"\": "+Colors.Fy+child.id+Colors.R);
            }
        }
    }
    async findFile(dir:any[],filename:string) {
        for (let i = 0; i < dir.length; i++) {
            if (dir[i].name==filename) return dir[i];
        }
    }
    async getExcelUsedRange(id:string,sheet:string) {// https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
        return (await this.apiGet("/me/drive/items/"+id+"/workbook/worksheets/"+sheet+"/usedRange"));
    }
    async getExcelRange(id:string,sheet:string):Promise<any> {// https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
        return (await this.apiGet("/me/drive/items/"+id+"/workbook/worksheets/"+sheet+"/range"));
    }
    async getExcelRangeByAddress(id:string,sheet:string,range:string):Promise<any> {// https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
        if (range.match(/^[A-Z]+\d+:[A-Z]+\d+$/g)==null)return null;
        return (await this.apiGet("/me/drive/items/"+id+"/workbook/worksheets/"+sheet+"/range(address=\'"+range+"\')"));
    }
    async getExcelNames(id:string):Promise<any[]> {
        return (await this.apiGet("/me/drive/items/"+id+"/workbook/names")).value;
    }
    async getExcelRangeByName(id:string,name:string):Promise<any> {// https://learn.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http
        return (await this.apiGet("/me/drive/items/"+id+"/workbook/names/"+name+"/range"));
    }
    async getExcelRangeLarge(id:string,sheet:string,range:string,stepSize:number):Promise<string[][]> {
        if (range.match(/^[A-Z]+\d+:[A-Z]+\d+$/g)==null)return [];
        const left:string = ""+range.match(/^[A-Z]+(?=\d+:[A-Z]+\d+$)/g);
        const top:number = parseInt(""+range.match(/(?<=^[A-Z]+)\d+(?=:[A-Z]+\d+$)/g));
        const right:string = ""+range.match(/(?<=^[A-Z]+\d+:)[A-Z]+(?=\d+$)/g);
        const bottom:number = parseInt(""+range.match(/(?<=^[A-Z]+\d+:[A-Z]+)\d+$/g));
        if (top-bottom > 500000) return [];
        var out:string[][] = [];
        var promises:Promise<void>[]=[];
        for (let i = top; i < bottom; i+=stepSize) {
            promises.push(new Promise<void>(async(resolve:()=>void)=>{
                const range:any[] = (await this.getExcelRangeByAddress(id,sheet,left+i+":"+right+Math.min((i+stepSize),bottom))).text;
                for (let j = 0; j < range.length; j++) {
                    out[i-top+j]=range[j];
                }
                resolve();
            }));
        }
        await Promise.all(promises);
        return out;
    }
    async getExcelColumns(id:string,sheet:string,left:letter,right:letter,stepSize:number):Promise<string[][]> {
        var start:number = (new Date()).getTime();
        var Continue:boolean = true;
        var out:string[][] = [];
        for (let i = 0; Continue; i++) {
            const range:any[] = (await this.getExcelRangeByAddress(id,sheet,left+(i*stepSize+1)+":"+right+((i+1)*stepSize))).text;
            var numEmpty:number=0;
            for (let j = 0; j < range.length; j++) {
                if (range[j][0]=="") {
                    if (numEmpty==9) { Continue=false;break; }
                    numEmpty++;continue;
                } else numEmpty=0;
                out.push(range[j]);
            }
            const time:number = ((new Date()).getTime()-start);
            console.log(out.length+": "+round3(time/1000)+"s");
        }
        return out;
    }
    async getExcelColumnsLarge(id:string,sheet:string,left:letter,right:letter,stepSize:number,numSteps:number):Promise<string[][]> {
        var start:number = (new Date()).getTime();
        var Continue:boolean = true;
        var out:string[][] = [];
        for (let i = 0; Continue; i++) {
        const range:string[][] = (await this.getExcelRangeLarge(id,sheet,left+(i*stepSize*numSteps+1)+":"+right+((i+1)*stepSize*numSteps),stepSize));
            var numEmpty:number=0;
            for (let j = 0; j < range.length; j++) {
                if (range[j][0]=="") {
                    if (numEmpty==9) { Continue=false;break; }
                    numEmpty++;continue;
                } else numEmpty=0;
                out.push(range[j]);
            }
        }
        return out;
    }
}

import {handle} from "./transferHandler";
async function run() {
    console.clear();
    console.log(Colors.Fgra+"Microsoft graph"+Colors.R);
    const start:number = (new Date()).getTime();
    const client:MicrosoftGraph = new MicrosoftGraph();
    const file:any = await client.findFile(await client.getRoot(),filename+".xlsx");
    const range:string[][] = await client.getExcelColumnsLarge(file.id,sheetName,left as letter,right as letter,10000,15);
    const time:number = (new Date()).getTime()-start;
    handle(range,__dirname+"/"+filename+".json",time);
}
run().catch((err:any)=>{console.log(err);});