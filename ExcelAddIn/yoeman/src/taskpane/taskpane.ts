/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
Office.onReady(function () { // Office is ready.
    $(document).ready(async()=>{ // The document is ready
        $("#get").on("click",get);
    });
});

class myRange {
    context:Excel.RequestContext;
    range:Excel.Range|undefined;

    address:string="";
    constructor(context:Excel.RequestContext) {
        this.context=context;
    }
    async update():Promise<void> {
        return new Promise<void>(async(resolve:()=>void)=>{
            if (this.range==null) { resolve();return; }
            this.range!.load("address");
            await this.context.sync();
            this.address = this.range.address;
            resolve();
        });
    }
    getSelectedRange():myRange {
        this.range = this.context.workbook.getSelectedRange();
        return this;
    }
    setValues(values:((string|number)[][])):myRange {
        if (this.range==null) return this;
        this.range.values=values;
        return this;
    }
    async getValues():Promise<any[][]> {
        return new Promise<any[][]>(async(resolve:(value:any[][])=>void)=>{
            this.range!.load("values");
            await this.context.sync();
            if (this.range==null) { resolve([]);return; }
            resolve(this.range.values);
        });
    }
    async getText():Promise<string[][]> {
        return new Promise<string[][]>(async(resolve:(value:string[][])=>void)=>{
            this.range!.load("text");
            await this.context.sync();
            if (this.range==null) { resolve([]);return; }
            resolve(this.range.text);
        });
    }
    setFill(color:string):myRange {
        if (this.range==null) return this;
        this.range.format.fill.color=color;
        return this;
    }
    clearFill():myRange {
        if (this.range==null) return this;
        this.range.format.fill.clear();
        return this;
    }
}
async function RequestContextAsync():Promise<Excel.RequestContext> {
    return new Promise<Excel.RequestContext>((resolve:(value:Excel.RequestContext)=>void)=>{
        Excel.run(async (context:Excel.RequestContext) => {
            resolve(context);
        });
    });
}
async function fetchJsonPost(url:string,data:any,headers:{[keys:string]:any}):Promise<string> {
    var tmpHeaders = headers||{};
    tmpHeaders["Content-Type"] = "application/json";
    return new Promise<string>((resolve:(value:string)=>void) => {
        fetch(url,{
            method:"POST",
            mode:"no-cors",
            headers:tmpHeaders,
            body: JSON.stringify(data)// body data type must match "Content-Type" header
        }).then((response:Response)=>response.text())
        .then((text:string)=>resolve(text));
    });
};
async function sha256(text:string):Promise<string> {
    return new Promise<string>((resolve:(value:string)=>void)=>{
        crypto.subtle
        .digest("SHA-256", new TextEncoder().encode(text))
        .then(h => {
            let hash = "",
            view = new DataView(h);
            for (let i = 0; i < view.byteLength; i += 4){
                hash += ("00000000" + view.getUint32(i).toString(16)).slice(-8);
            }
            resolve(hash);
        });
    });
}
const MAXLENGTH:number=500;
async function get() {
    try {
        var context:Excel.RequestContext = await RequestContextAsync();
        var range:myRange = new myRange(context).getSelectedRange();
        const bar:HTMLElement  = document.getElementById("bar")!;
        bar.setAttribute("style","--percent:0%");
        const statusEl:HTMLElement  = document.getElementById("status")!;
        statusEl.innerText="Processing.";

        var values:string[][] = await range.getText();
        console.log("Transfering "+values.length+" lines.");
        statusEl.innerText="Sending...";
        if (values.length<=MAXLENGTH) {
            const start:number = (new Date()).getTime();
            await fetchJsonPost("http://localhost:3000/wholeTransfer",{"data":values,"start":start},{});
            statusEl.innerText="Success!";
        } else {
            var start:number=(new Date()).getTime();
            var sha:string = await sha256(JSON.stringify(values));
            var id:string = JSON.parse(await fetchJsonPost("http://localhost:3000/startTransfer",{"sha":sha,"start":start},{})).id;
            var tmpArray:string[][] = [];
            for (let i = 0; i < values.length; i++) {
                tmpArray.push(values[i]);
                if (tmpArray.length==MAXLENGTH||i==values.length-1) {
                    await fetchJsonPost("http://localhost:3000/transfer?id="+id,tmpArray,{});
                    bar.setAttribute("style","--percent:"+((i+1)*100/values.length).toString()+"%");
                    tmpArray=[];
                }
            }
            var status:boolean = JSON.parse(await fetchJsonPost("http://localhost:3000/endTransfer?id="+id,{},{})).status;
            console.log(status?"success":"fail");
            statusEl.innerText=(status?"Success!":"Fail.");
            console.log("Time to send: "+((new Date()).getTime()-start));
        }
        bar.setAttribute("style","--percent:100%");
    } catch (err:any) {}
}