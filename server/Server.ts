//#region imports/libs
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
function generateUUID():string {
	var a = (new Date()).getTime();//Timestamp
	return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
		var b = Math.random() * 16;//random number between 0 and 16
		b = (a + b)%16 | 0;
		a = Math.floor(a/16);
		return (c === 'x' ? b : (b & 0x3 | 0x8)).toString(16);
	});
}
function sha256(input:string):string {
    return (require('crypto').createHash("sha256").update(input).digest("hex"));
}

const https = require("https");
const express = require("express");
import * as fs from "fs"

var tmp:string[] = __dirname.split("\\");tmp.pop();
const dirname:string = tmp.join("/");//one back from current path
//#endregion imports/libs

console.clear();

//#region Microsoft Graph
const clientId:string = fs.readFileSync(dirname+"/graphTokens/ClientId.txt").toString();
const clientSecretValue:string = fs.readFileSync(dirname+"/graphTokens/ClientSecret.txt").toString();
const scope:string = encodeURIComponent("User.Read offline_access Files.ReadWrite");
const redirect:string = "http://localhost:80/oauth/";// https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Authentication/appId/06e160b0-cc2e-4181-bf03-7aa31e0bf435/isMSAApp/true
var authUrl:string = "https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?";
authUrl+= "client_id="+clientId+"&response_type=code";
authUrl+= "&redirect_uri="+encodeURIComponent(redirect);
authUrl+= "&response_mode=query"+"&scope="+scope+"&state=123456789";

var app:any = express();
app.use(express.urlencoded({ extended: true }));
app.use(require("body-parser").text());
var pages:{[key:string]:any} = {
    "/":(req:any,res:any,send:any)=>res.redirect("/index.html"),
    "/index.html":(req:any,res:any,send:any)=>send(),
    "/login":(req:any,res:any,send:any)=>res.redirect(authUrl),
    "/logout":(req:any,res:any,send:any)=>send("/logout.html"),
    "/oauth":(req:any,res:any,send:any)=>{ // console.log(req.query);
        if (req.query.code!=null) requestToken(req.query.code);
        res.redirect("/index.html");
    }
};
Object.keys(pages).forEach((i) => {
    app.get(i, (req:any, res:any) => {
        pages[i](req,res,((page:string)=>{ res.sendFile(__dirname+"/webpage"+(page!=null?page:i),"utf8"); }));
    });
});
app.listen(80, function () {
    console.clear();
    console.log(Colors.Fgra+"Microsoft Graph server is running at: "+Colors.Fgre+"http://localhost:80"+Colors.R);
});

var refreshTimeout:number|undefined;
// https://learn.microsoft.com/en-us/graph/auth-v2-user?tabs=http
async function requestToken(code:string):Promise<void> {
    var params:string = "client_id="+clientId+"&scope="+scope;
    params+="&code="+code+"&redirect_uri="+encodeURIComponent(redirect);
    params+="&grant_type=authorization_code&client_secret="+clientSecretValue;
    httpsRequestPromise("login.microsoftonline.com","/consumers/oauth2/v2.0/token","POST",{"Content-Type":"application/x-www-form-urlencoded"},params)
    .then((out:string)=>{
        const json = JSON.parse(out);
        fs.writeFileSync(dirname+"/graphTokens/access_token.txt",json.access_token);
        fs.writeFileSync(dirname+"/graphTokens/refresh_token.txt",json.refresh_token);
        console.log(Colors.Fgra+"Access token saved to file."+Colors.R);
        console.log(Colors.Fgra+"Refresh token saved to file.\n"+Colors.R);
        if (refreshTimeout!=undefined) clearTimeout(refreshTimeout);
        refreshTimeout = setTimeout(() => {
            refreshToken();
            refreshTimeout=undefined;
        }, 1000*60*60) as unknown as number;//one hour
    });
    return;
}
async function refreshToken():Promise<void> {
    var params:string = "client_id="+clientId+"&scope="+scope;
    params+="&refresh_token="+(fs.readFileSync(dirname+"/graphTokens/refresh_token.txt").toString());
    params+="&grant_type=refresh_token&client_secret="+clientSecretValue;
    httpsRequestPromise("login.microsoftonline.com","/consumers/oauth2/v2.0/token","GET",{"Content-Type":"application/x-www-form-urlencoded"},params)
    .then((out:string)=>{
        const json = JSON.parse(out);
        fs.writeFileSync(dirname+"/graphTokens/access_token.txt",json.access_token);
        fs.writeFileSync(dirname+"/graphTokens/refresh_token.txt",json.refresh_token);
        console.log(Colors.Fgra+"\"Refreshed\" access token saved to file."+Colors.R);
        console.log(Colors.Fgra+"\"Refreshed\" refresh token saved to file.\n"+Colors.R);
        if (refreshTimeout!=undefined) clearTimeout(refreshTimeout);
        refreshTimeout = setTimeout(() => {
            refreshToken();
            refreshTimeout=undefined;
        }, 1000*60*60) as unknown as number;//one hour
    });
    return;
}
//#endregion Microsoft Graph

//#region Excel Add-in
const addInApp:any = express();
addInApp.use(express.urlencoded({ extended: true }));
addInApp.use(require("body-parser").text());
var addInPages:{[key:string]:((req:any,res:any,send:any)=>void)} = {
    "/assets/logo-filled.png" :(req:any,res:any,send:any)=>send(),
    "/assets/icon-16.png"     :(req:any,res:any,send:any)=>send(),
    "/assets/icon-32.png"     :(req:any,res:any,send:any)=>send(),
    "/assets/icon-64.png"     :(req:any,res:any,send:any)=>send(),
    "/assets/icon-80.png"     :(req:any,res:any,send:any)=>send(),
    "/assets/icon-128.png"    :(req:any,res:any,send:any)=>send(),
    "/d916c354314f23a46dd2.css" :(req:any,res:any,send:any)=>send(),
    "/manifest.xml"             :(req:any,res:any,send:any)=>send(),
    "/commands.js.map" :(req:any,res:any,send:any)=>send(),
    "/commands.html"   :(req:any,res:any,send:any)=>send(),
    "/commands.js"     :(req:any,res:any,send:any)=>send(),
    "/polyfill.js.map" :(req:any,res:any,send:any)=>send(),
    "/polyfill.js"     :(req:any,res:any,send:any)=>send(),
    "/taskpane.js.map" :(req:any,res:any,send:any)=>send(),
    "/taskpane.html"   :(req:any,res:any,send:any)=>send(),
    "/taskpane.js"     :(req:any,res:any,send:any)=>send(),
    "/taskpane"   :(req:any,res:any,send:any)=>send("/taskpane.html")
};
Object.keys(addInPages).forEach((i) => {
    addInApp.get(i, (req:any, res:any) => {
        addInPages[i](req, res, ((page:string) => { res.sendFile(dirname+"/ExcelAddIn/yoeman/dist"+(page!=null?page:i),"utf8"); }));
    });
});

import {handle} from "../transferHandler";
const filename:string = JSON.parse(fs.readFileSync(dirname+"/config.json").toString()).filename;

type transferObj = {data:any[][],sha:string,start:number}
var transfers:{[key:string]:transferObj}={};
addInApp.post("/wholeTransfer",(req:any, res:any)=>{
    res.send("{}");
    const body:any = JSON.parse(req.body);
    const start:number = parseInt(body.start);
    const time:number = (new Date()).getTime()-start;
    handle(body.data,dirname+"/"+filename+".json",time);
});
addInApp.post("/startTransfer",(req:any, res:any)=>{
    const id:string = generateUUID();
    const body:any = JSON.parse(req.body);
    const sha:string = body.sha;
    const start:number = body.start;
    if (sha != null) {
        res.send("{\"id\":\""+id+"\"}");
        transfers[id] = {
            data:[],
            "sha":sha,
            "start":start
        }
    }
});
addInApp.post("/transfer",(req:any, res:any)=>{
    res.send("{}");
    const body:any = JSON.parse(req.body);
    const id:string = req.query.id;
    if (id!=null&&transfers[id]!=null) {
        transfers[id].data.push(...body);
    }
});
addInApp.post("/endTransfer",(req:any, res:any)=>{
    const id:string = req.query.id;
    if (id!=null&&transfers[id]!=null) {
        const transfer = transfers[id].data;
        const sha:string=sha256(JSON.stringify(transfer));
        const status:boolean = (sha==transfers[id].sha);
        const time:number = (new Date()).getTime()-transfers[id].start
        delete transfers[id];
        if (status) handle(transfer,dirname+"/"+filename+".json",time);
        res.send("{\"status\":"+status+"}");
    } else {
        res.send("{\"status\":false}");
    }
});
addInApp.listen(3000, function () {
    console.log(Colors.Fgra+"Excel Add-in server is running at: "+Colors.Fgre+"http://localhost:3000"+Colors.R);
});
//#endregion Excel Add-in