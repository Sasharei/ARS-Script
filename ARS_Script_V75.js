// == ARS Suite v75 (MULTI & Rush-Deep) ==
// Adds: multi‑composition export queue; deeper rush detection (A## or AA##), preserves all naming/SharePoint/EXIF logic.
// Based on earlier stable UI/logic.

function _esc(s){ if(s==null)return ""; s=String(s); return s.replace(/\\/g,"\\\\").replace(/"/g,'\\"').replace(/\n/g,"\\n").replace(/\r/g,"\\r"); }
function _q(s){ if(!s)return'""'; return '"' + String(s).replace(/\\/g,"\\\\").replace(/"/g,'\\"') + '"'; }
function _iso(){ var d=new Date(); function p(n){n=String(n); return n.length<2?'0'+n:n;} return d.getUTCFullYear()+"-"+p(d.getUTCMonth()+1)+"-"+p(d.getUTCDate())+"T"+p(d.getUTCHours())+":"+p(d.getUTCMinutes())+":"+p(d.getUTCSeconds())+"Z"; }
function _json(o){ var a=[],k; for(k in o) if(o.hasOwnProperty(k)){ var v=o[k],t=typeof v; a.push('"' + k + '":' + (t==="number"||t==="boolean" ? String(v) : ('"' + _esc(v) + '"'))); } return "{"+a.join(",")+"}"; }
function _waitStable(path,maxS,probeMs,passes){ var f=new File(path),w=0,last=-1,ok=0,i=probeMs||1500,m=maxS||1800,p=passes||3; while(w<m){ if(f.exists){ var sz=f.length; if(sz===last&&sz>0){ ok++; if(ok>=p) return true; } else ok=0; last=sz; } $.sleep(i); w+=(i/1000);} return false; }
function _exifCmd(exe, mp4, meta){
    var js=_json({
        AuthorInitials:meta.AuthorInitials||"", GameCode:meta.GameCode||"", VideoType:meta.VideoType||"",
        Number:meta.Number||"", Rushes:meta.Rushes||"", Resolution:meta.Resolution||"", Language:meta.Language||"",
        DurationSec:Number(meta.DurationSec||0), ARSSuiteVersion:meta.ARSSuiteVersion||"", ExportDate:meta.ExportDate||_iso(),
        Type:meta.Type||"", Intention:meta.Intention||"", Direction:meta.Direction||""
    });
    var iso=_iso(); var args=[];
    args.push(_q(exe));
    args.push("-overwrite_original_in_place");
    if (meta.AuthorInitials) args.push("-XMP-dc:creator="+_q(meta.AuthorInitials));
    args.push("-XMP-xmp:CreatorTool="+_q("ARS Suite"));
    if (meta.HumanTitle) args.push("-XMP-dc:title="+_q(meta.HumanTitle));
    args.push("-XMP-xmp:CreateDate="+_q(iso));
    args.push("-XMP-xmp:MetadataDate="+_q(iso));
    args.push("-XMP-dc:description="+_q(js));
    if (meta.Type) args.push("-XMP-dc:subject+="+_q("Type:"+meta.Type));
    if (meta.Intention) args.push("-XMP-dc:subject+="+_q("Intention:"+meta.Intention));
    if (meta.Direction) args.push("-XMP-dc:subject+="+_q("Direction:"+meta.Direction));
    if (meta.__Comment) args.push("-Comment="+_q(meta.__Comment));
    args.push(_q(mp4));
    return args.join(" ");
}

(function(){
    var win=new Window("palette","ARS Suite v75",undefined,{resizeable:true});
    win.orientation="row"; win.alignChildren="fill";
    var left=win.add("group"); left.orientation="column"; left.alignChildren="fill";
    var right=win.add("group"); right.orientation="column"; right.alignChildren="fill";

    // ----------------- RESIZE PANEL (unchanged UI) -----------------
    var rp=left.add("panel",undefined,"Resize & CTA"); rp.orientation="column"; rp.alignChildren="left";
    rp.add("statictext",undefined,"Scale Mode:"); var ddScale=rp.add("dropdownlist",undefined,["fit","fill","crop","blur","BG"]); ddScale.selection=0;
    rp.add("statictext",undefined,"Formats:");
    var fc={"Landscape":rp.add("checkbox",undefined,"Landscape (1920x1080)"),"Square":rp.add("checkbox",undefined,"Square (1080x1080)"),"Vertical":rp.add("checkbox",undefined,"Vertical (1080x1920)"),"4:5":rp.add("checkbox",undefined,"4:5 (1080x1350)"),"Custom":rp.add("checkbox",undefined,"Custom"),"Mintegral landscape":rp.add("checkbox",undefined,"Mintegral landscape (1280x720)"),"Mintegral portrait":rp.add("checkbox",undefined,"Mintegral portrait (720x1280)")}; fc["Vertical"].value=true;
    rp.add("statictext",undefined,"Landscape Align:"); var ddAlign=rp.add("dropdownlist",undefined,["Left","Center"]); ddAlign.selection=0;
    var cg=rp.add("group"); cg.orientation="row"; cg.add("statictext",undefined,"Width:"); var etW=cg.add("edittext",undefined,"1080"); etW.characters=6; cg.add("statictext",undefined,"Height:"); var etH=cg.add("edittext",undefined,"1920"); etH.characters=6;
    var cbCTA=rp.add("checkbox",undefined,"Add CTA"); var etCTA=rp.add("edittext",undefined,"PLAY NOW"); etCTA.characters=15;
    var btnResize=rp.add("button",undefined,"Resize Composition");

    // ----------------- EXPORT PANEL -----------------
    var ep=right.add("panel",undefined,"Export"); ep.orientation="column"; ep.alignChildren="left";
    ep.add("statictext",undefined,"Creator Initials:"); var ddInit=ep.add("dropdownlist",undefined,["AR","AA","MI","PG","RA","S","E","H"]); ddInit.selection=0; var etInit=ep.add("edittext",undefined,""); etInit.characters=10;
    ep.add("statictext",undefined,"Game Initials:"); var ddGame=ep.add("dropdownlist",undefined,["GR","TJ"]); ddGame.selection=0; var etGame=ep.add("edittext",undefined,""); etGame.characters=10;
    ep.add("statictext",undefined,"Video Type:"); var ddVType=ep.add("dropdownlist",undefined,["R","F","H"]); ddVType.selection=0;
    ep.add("statictext",undefined,"Language:"); var ddLang=ep.add("dropdownlist",undefined,["EN","FR","DE","ES","RU"]); ddLang.selection=0;
    ep.add("statictext",undefined,"Export Location:"); var ddLoc=ep.add("dropdownlist",undefined,["SharePoint","Custom"]); ddLoc.selection=0;
    ep.add("statictext",undefined,"Render using:"); var ddRender=ep.add("dropdownlist",undefined,["Media Encoder","Render Queue"]); ddRender.selection=1;

    ep.add("panel",undefined,"Meta");
    ep.add("statictext",undefined,"Type:"); var ddType=ep.add("dropdownlist",undefined,["Disruptive","Iteration","MarketInsp","ConceptInsp"]); ddType.selection=0;
    ep.add("statictext",undefined,"Intention:"); var ddIntent=ep.add("dropdownlist",undefined,["Win","Fail","Noob","Pro","Satisfying","Frustrating","Custom"]); ddIntent.selection=0; var etIntent=ep.add("edittext",undefined,""); etIntent.characters=18; etIntent.visible=false;
    ep.add("statictext",undefined,"Direction:"); var ddDir=ep.add("dropdownlist",undefined,["None","INF","UGC","AI","Intro","Hook","DIY","End Card","Theme","Cinematic","Gameplay","Custom"]); ddDir.selection=0; var etDir=ep.add("edittext",undefined,""); etDir.characters=18; etDir.visible=false;
    ddIntent.onChange=function(){ etIntent.visible=(ddIntent.selection && ddIntent.selection.text==="Custom"); };
    ddDir.onChange=function(){ etDir.visible=(ddDir.selection && ddDir.selection.text==="Custom"); };

    var btnExport=ep.add("button",undefined,"EXPORT (Active or Selected Comps)");

    rp.preferredSize=[360,-1]; ep.preferredSize=[360,-1]; win.minimumSize=[760,520];

    // ----------------- Helpers -----------------
    function addCTA(nc,W,H,center,label){
        var n=nc.layers.addNull(); n.name="CTA_Null";
        var btnX=(W>H)? W*0.75 : center[0];
        var btnY=(W>H)? center[1] : H - 230;
        n.property("ADBE Transform Group").property("ADBE Position").setValue([btnX,btnY]);
        var sp=n.property("ADBE Transform Group").property("ADBE Scale");
        sp.setValueAtTime(0,[100,100]); sp.setValueAtTime(0.4,[115,115]); sp.setValueAtTime(0.8,[100,100]);
        sp.expression="loopOut('cycle')";

        var sh = nc.layers.addShape(); sh.parent=n;
        var g  = sh.property("ADBE Root Vectors Group");
        var rect = g.addProperty("ADBE Vector Shape - Rect");
        var wcta = Math.min(600, 100 + (label?label.length:8)*40);
        rect.property("ADBE Vector Rect Size").setValue([wcta,130]);
        rect.property("ADBE Vector Rect Roundness").setValue(40);
        try{ g.addProperty("ADBE Vector Transform Group"); }catch(e){}
        try{ g.property("ADBE Vector Transform Group").property("ADBE Vector Position").setValue([0,0]); }catch(e){}
        var fill = g.addProperty("ADBE Vector Graphic - Fill"); fill.property("ADBE Vector Fill Color").setValue([0.2,0.9,0.2]);
        sh.property("ADBE Transform Group").property("ADBE Anchor Point").setValue([0,0]);
        sh.property("ADBE Transform Group").property("ADBE Position").setValue([0,0]);

        var tx = nc.layers.addText(label||"PLAY NOW"); tx.parent=n;
        var td = tx.property("Source Text").value; td.fontSize=80; td.fillColor=[1,1,1]; td.font="Impact"; td.justification=ParagraphJustification.CENTER_JUSTIFY;
        tx.property("Source Text").setValue(td);
        try{
            tx.property("ADBE Transform Group").property("ADBE Anchor Point").expression =
                "var r=sourceRectAtTime(time,false); [r.left + r.width/2, r.top + r.height/2];";
        }catch(e){}
        tx.property("ADBE Transform Group").property("ADBE Position").setValue([0,0]);
        tx.moveToBeginning(); sh.moveAfter(tx); n.moveAfter(sh);
    }

    function getSelectedComps(){
        var out=[]; var sel=app.project.selection; if(sel && sel.length){
            for(var i=0;i<sel.length;i++) if(sel[i] instanceof CompItem) out.push(sel[i]);
        }
        var active=app.project.activeItem; if(out.length===0 && (active instanceof CompItem)) out.push(active);
        return out;
    }

    function detectNumberFromCompOrProject(comp){
        var cm = (comp && comp.name) ? comp.name.match(/\d{3,}/) : null;
        if (cm && cm.length>0) return cm[0];
        if (!app.project.file) return null;
        var pf=app.project.file.name, m=pf.match(/(\d{3,})/g); if(!m||m.length===0) return null;
        return m[m.length-1];
    }

    function collectRushTokens(comp){
        var rushAA=[], rushExtra=[]; function pushU(a,t){ for(var i=0;i<a.length;i++) if(a[i]===t) return; a.push(t); }
        function scanName(nm){
            if(!nm) return;
            var mAA = nm.match(/\b[A-Z]{1,2}\d{2}\b/g); if(mAA) for(var j=0;j<mAA.length;j++) pushU(rushAA,mAA[j]);
            var mINF = nm.match(/\bINF\d+\b/g); if(mINF) for(var j2=0;j2<mINF.length;j2++) pushU(rushExtra,mINF[j2]);
            var mUGC = nm.match(/\bUGC\d+\b/g); if(mUGC) for(var j3=0;j3<mUGC.length;j3++) pushU(rushExtra,mUGC[j3]);
            var mHOOK = nm.match(/\bHOOK\d+\b/g); if(mHOOK) for(var j4=0;j4<mHOOK.length;j4++) pushU(rushExtra,mHOOK[j4]);
        }
        function scanComp(c, depth){ if(!c||depth>2) return; scanName(c.name); try{ for(var i=1;i<=c.numLayers;i++){ varlyr=c.layer(i); scanName(varlyr.name); if(varlyr instanceof AVLayer && varlyr.source){ scanName(varlyr.source.name); if(varlyr.source instanceof CompItem) scanComp(varlyr.source, depth+1); } } }catch(e){}
        }
        scanComp(comp,0);
        rushAA.sort();
        var rushes=""; if(rushAA.length>0) rushes=rushAA.join("-"); if(rushExtra.length>0) rushes+=(rushes!==""?"-":"")+rushExtra.join("-");
        return rushes;
    }

    function buildPlanForComp(comp, session){
        var number=detectNumberFromCompOrProject(comp);
        if(!number) throw new Error("No number found in comp/project name (need 3+ digits).");
        var initials=(session.etInit&&session.etInit.length>0)?session.etInit:session.ddInit;
        var game=(session.etGame&&session.etGame.length>0)?session.etGame:session.ddGame;
        var videoType=session.ddVType, lang=session.ddLang;
        var type2=session.ddType, intention=session.intent, direction=session.dir;

        var rushes=collectRushTokens(comp);
        var resolution=comp.width+"x"+comp.height;
        var durationSec=Math.round(comp.workAreaDuration>0?comp.workAreaDuration:comp.duration);
        var durationHuman=durationSec+"s";
        var fileName="["+initials+"]_"+game+"_"+videoType+"_"+number+(rushes?("_"+rushes):"")+"_"+resolution+"_"+lang+"_"+durationHuman+".mp4";
        fileName=fileName.replace(/__+/g,"_").replace(/_-/g,"_").replace(/_+$/,"");

        var exportRoot=session.exportRoot; var location=session.ddLoc;
        if(location==="SharePoint"){
            if(!exportRoot) throw new Error("SharePoint export folder not set.");
            var targetFolder=new Folder(exportRoot + "/" + number); if(!targetFolder.exists) targetFolder.create();
            var outFile=new File(targetFolder.fsName + "/" + fileName);
            return {comp:comp, number:number, rushes:rushes, resolution:resolution, durationSec:durationSec, durationHuman:durationHuman, fileName:fileName, outFile:outFile, initials:initials, game:game, videoType:videoType, lang:lang, type2:type2, intention:intention, direction:direction};
        } else {
            if(!exportRoot) throw new Error("Custom export folder not set.");
            var target=new Folder(exportRoot + "/" + number); if(!target.exists) target.create();
            var out=new File(target.fsName + "/" + fileName);
            return {comp:comp, number:number, rushes:rushes, resolution:resolution, durationSec:durationSec, durationHuman:durationHuman, fileName:fileName, outFile:out, initials:initials, game:game, videoType:videoType, lang:lang, type2:type2, intention:intention, direction:direction};
        }
    }

    function enqueueAndRender(plans, renderer, exiftoolPath){
        // Add all to RQ first
        for(var i=0;i<plans.length;i++){
            var pq=plans[i];
            var rqItem=app.project.renderQueue.items.add(pq.comp);
            rqItem.outputModule(1).file=pq.outFile;
        }
        // Trigger once
        if(renderer==="Media Encoder") app.project.renderQueue.queueInAME(true); else app.project.renderQueue.render();
        // EXIF per file (poll until stable)
        if(exiftoolPath){
            for(var j=0;j<plans.length;j++){
                var p=plans[j];
                _waitStable(p.outFile.fsName, 1800, 1500, 3);
                try{
                    var meta={ AuthorInitials:p.initials, GameCode:p.game, VideoType:p.videoType, Number:p.number, Rushes:p.rushes, Resolution:p.resolution, Language:p.lang, DurationSec:p.durationSec, ARSSuiteVersion:"v75", ExportDate:_iso(), Type:p.type2, Intention:p.intention, Direction:p.direction, HumanTitle:p.fileName.replace(/\.mp4$/i,"") , __Comment:("Creator: "+p.initials+" | Game: "+p.game+" | Type: "+p.videoType+" | Rushes: "+(p.rushes||"-")+" | Lang: "+p.lang+" | Number: "+p.number+" | Duration: "+p.durationHuman+" | MetaType: "+p.type2+" | Intention: "+p.intention+" | Direction: "+p.direction) };
                    var cmd=_exifCmd(exiftoolPath, p.outFile.fsName, meta);
                    var out=system.callSystem(cmd);
                }catch(ex){}
            }
        }
    }

    // ----------------- RESIZE HANDLER (kept same behavior) -----------------
    btnResize.onClick=function(){
        var comp=app.project.activeItem; if(!(comp instanceof CompItem)){ alert("[ARS] Select a composition to resize."); return; }
        app.beginUndoGroup("ARS Resize");
        var formats={"Landscape":[1920,1080],"Square":[1080,1080],"Vertical":[1080,1920],"4:5":[1080,1350],"Mintegral landscape":[1280,720],"Mintegral portrait":[720,1280]};
        if(fc["Custom"].value){ var w=parseInt(etW.text,10), h=parseInt(etH.text,10); if(!isNaN(w)&&!isNaN(h)&&w>0&&h>0) formats["Custom"]= [w,h]; }
        var mode=(ddScale.selection?ddScale.selection.text:"fit");
        var bgFile=null, bgItem=null;
        if(mode==="BG"){
            bgFile = File.openDialog("Pick background image for BG mode (PNG/JPG)");
            if(bgFile){ try{ bgItem = app.project.importFile(new ImportOptions(bgFile)); }catch(e){ bgItem=null; } }
            if(!bgItem){ alert("[ARS] No background chosen. Will use solid color."); }
        }
        var key;
        for (key in formats){
            if(!fc[key]||!fc[key].value) continue;
            var WH=formats[key], W=WH[0], H=WH[1];
            var nc=app.project.items.addComp(comp.name+"_"+W+"x"+H, W,H, comp.pixelAspect, comp.duration, comp.frameRate);
            var center=[W/2,H/2];
            var sW=W/comp.width*100, sH=H/comp.height*100;
            var scaleFit=Math.min(sW,sH), scaleFill=Math.max(sW,sH);
            if(mode==="blur"){
                var bg=nc.layers.add(comp); bg.audioEnabled=false; bg.moveToBeginning(); bg.scale.setValue([scaleFill,scaleFill]);
                try{ var bf=bg.Effects.addProperty("ADBE Gaussian Blur 2"); bf.property("Blurriness").setValue(150); bf.property("Repeat Edge Pixels").setValue(true);}catch(e){}
            } else if(mode==="BG"){
                if(bgItem){ var bgL = nc.layers.add(bgItem); bgL.moveToBeginning(); var sx=W/bgL.source.width*100, sy=H/bgL.source.height*100; var s=Math.max(sx,sy); bgL.property("ADBE Transform Group").property("ADBE Scale").setValue([s,s]); bgL.property("ADBE Transform Group").property("ADBE Position").setValue(center); }
                else { var solid=nc.layers.addSolid([0.08,0.08,0.08],"BG_Solid",W,H,1.0); solid.moveToBeginning(); }
            }
            var L=nc.layers.add(comp); L.audioEnabled=true; L.moveToBeginning();
            if(mode==="fill") L.scale.setValue([scaleFill,scaleFill]);
            else if(mode==="crop"){ var sc=sW; L.scale.setValue([sc,sc]); }
            else L.scale.setValue([scaleFit,scaleFit]);
            var pos=center; if(W>H && (key==="Landscape" || key.indexOf("Mintegral")==0)){ pos=(ddAlign.selection && ddAlign.selection.text==="Center")?center:[W*0.35, center[1]]; }
            L.property("ADBE Transform Group").property("ADBE Position").setValue(pos);
            if(cbCTA.value) addCTA(nc,W,H,center, etCTA.text);
        }
        app.endUndoGroup();
        alert("[ARS] Resize done.");
    };

    // ----------------- EXPORT HANDLER (now supports multi) -----------------
    var __sessionSharepointPath=null;
    btnExport.onClick=function(){
        var comps=getSelectedComps(); if(!comps.length){ alert("Select a composition or select comps in Project panel."); return; }
        // Collect session settings
        var renderer = (ddRender.selection?ddRender.selection.text:"Render Queue");
        var location = (ddLoc.selection?ddLoc.selection.text:"SharePoint");
        var session={
            ddInit:(ddInit.selection?ddInit.selection.text:"AR"), etInit:etInit.text||"",
            ddGame:(ddGame.selection?ddGame.selection.text:"GR"), etGame:etGame.text||"",
            ddVType:(ddVType.selection?ddVType.selection.text:"R"), ddLang:(ddLang.selection?ddLang.selection.text:"EN"),
            ddType:(ddType.selection?ddType.selection.text:"Disruptive"),
            intent: (ddIntent.selection && ddIntent.selection.text==="Custom")? (etIntent.text||"Custom") : (ddIntent.selection?ddIntent.selection.text:"Win"),
            dir: (ddDir.selection && ddDir.selection.text==="Custom")? (etDir.text||"Custom") : (ddDir.selection?ddDir.selection.text:"None"),
            ddLoc: location,
            exportRoot: null
        };
        if(location==="SharePoint"){
            if(!__sessionSharepointPath){ var f=Folder.selectDialog("Select SharePoint export folder (this session)"); if(!f) return; __sessionSharepointPath=f.fsName; }
            session.exportRoot=__sessionSharepointPath;
        } else {
            var f2=Folder.selectDialog("Choose custom export folder"); if(!f2) return; session.exportRoot=f2.fsName;
        }
        // ExifTool path resolve once
        var exiftoolPath=""; try{ if(app.settings.haveSetting("ARS","exiftoolPath")) exiftoolPath=app.settings.getSetting("ARS","exiftoolPath"); }catch(e){ exiftoolPath=""; }
        if(!exiftoolPath || !File(exiftoolPath).exists){ var pick=File.openDialog("Select exiftool.exe (for embedding XMP into MP4)"); if(pick){ exiftoolPath=pick.fsName; try{ app.settings.saveSetting("ARS","exiftoolPath",exiftoolPath);}catch(e){} } else { exiftoolPath=null; } }

        // Build plans
        var plans=[]; var errs=[];
        for(var i=0;i<comps.length;i++){
            try{ plans.push( buildPlanForComp(comps[i], session) ); }
            catch(ex){ errs.push(comps[i].name+": "+ex.toString()); }
        }
        if(!plans.length){ alert("Nothing to export.\n"+(errs.length?("Errors:\n"+errs.join("\n")):"")); return; }

        // Save sidecar project with number of the LAST plan to keep old behavior per comp (optional)
        try{ var lastN=plans[plans.length-1].number; var projFile=new File(plans[plans.length-1].outFile.path + "/" + lastN + ".aep"); app.project.save(projFile); }catch(e){}

        // Enqueue & render & embed XMP
        enqueueAndRender(plans, renderer, exiftoolPath);

        // Summary dialog
        var msg=["Export queued:"].concat(plans.map(function(p){ return "\n"+p.outFile.fsName; }));
        if(errs.length) msg.push("\n\nSkipped:"+"\n"+errs.join("\n"));
        alert(msg.join(""));
    };

    win.center(); win.show();
})();
