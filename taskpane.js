/* ===== Grid Resize Tool – Compact Edition ===== */
var CM=28.3465, MIN=0.1, gridUnitCm=0.21, apiOk=false, GTAG="DROEGE_GUIDELINE";

Office.onReady(function(info){
    if(info.host===Office.HostType.PowerPoint){
        apiOk=!!(Office.context.requirements&&Office.context.requirements.isSetSupported&&Office.context.requirements.isSetSupported("PowerPointApi","1.5"));
        initUI();
        if(!apiOk) showStatus("⚠ PowerPointApi 1.5 nicht verfügbar – Shape-Funktionen nur auf Desktop/Web","warning");
    }
});

function initUI(){
    // Grid Unit
    var gi=document.getElementById("gridUnit");
    gi.addEventListener("change",function(){var v=parseFloat(this.value);if(!isNaN(v)&&v>0){gridUnitCm=v;upPre(v);showStatus("RE: "+v.toFixed(2)+" cm","info")}});
    document.querySelectorAll(".pre").forEach(function(b){
        b.addEventListener("click",function(){var v=parseFloat(this.dataset.value);gridUnitCm=v;gi.value=v;upPre(v);showStatus("RE: "+v.toFixed(2)+" cm","info")});
    });

    // Tabs
    document.querySelectorAll(".tab").forEach(function(b){
        b.addEventListener("click",function(){
            var id=this.dataset.tab;
            document.querySelectorAll(".tab").forEach(function(t){t.classList.remove("active")});
            document.querySelectorAll(".pane").forEach(function(p){p.classList.remove("active")});
            this.classList.add("active");
            document.getElementById(id).classList.add("active");
        });
    });

    // Tab 1: Resize
    bind("shrinkWidth",function(){resize("width",-gridUnitCm)});
    bind("growWidth",function(){resize("width",gridUnitCm)});
    bind("shrinkHeight",function(){resize("height",-gridUnitCm)});
    bind("growHeight",function(){resize("height",gridUnitCm)});
    bind("shrinkBoth",function(){resize("both",-gridUnitCm)});
    bind("growBoth",function(){resize("both",gridUnitCm)});
    bind("propShrink",function(){propResize(-gridUnitCm)});
    bind("propGrow",function(){propResize(gridUnitCm)});

    // Tab 1: Match
    bind("matchWidthMax",function(){match("width","max")});
    bind("matchWidthMin",function(){match("width","min")});
    bind("matchHeightMax",function(){match("height","max")});
    bind("matchHeightMin",function(){match("height","min")});
    bind("matchBothMax",function(){match("both","max")});
    bind("matchBothMin",function(){match("both","min")});
    bind("propMatchMax",function(){propMatch("max")});
    bind("propMatchMin",function(){propMatch("min")});

    // Tab 2: Grid
    bind("snapPosition",function(){snap("position")});
    bind("snapSize",function(){snap("size")});
    bind("snapBoth",function(){snap("both")});
    bind("setSpacingH",function(){spacing("horizontal")});
    bind("setSpacingV",function(){spacing("vertical")});
    bind("showInfo",function(){shapeInfo()});
    bind("createTable",function(){createGridTable()});

    // Tab 3: Setup
    bind("setSlideSize",function(){setSlideSize()});
    bind("toggleGuidelines",function(){toggleGuides()});
    bind("copyShadowText",function(){copyShadow()});
}

function bind(id,fn){document.getElementById(id).addEventListener("click",fn)}
function upPre(v){document.querySelectorAll(".pre").forEach(function(b){b.classList.toggle("active",Math.abs(parseFloat(b.dataset.value)-v)<.001)})}

/* STATUS: dauerhaft sichtbar – Text bleibt stehen bis zur nächsten Aktion */
function showStatus(m,t){
    var e=document.getElementById("status");
    e.textContent=m;
    e.className="sts "+(t||"info");
}

function c2p(c){return c*CM}
function p2c(p){return p/CM}
function rnd(v){return Math.round(v/gridUnitCm)*gridUnitCm}

function withShapes(min,cb){
    if(!apiOk){showStatus("Nicht unterstützt (PowerPointApi 1.5 erforderlich)","error");return}
    PowerPoint.run(function(ctx){
        var sh=ctx.presentation.getSelectedShapes();sh.load("items");
        return ctx.sync().then(function(){
            if(sh.items.length<min){showStatus(min<=1?"Bitte Objekt(e) auswählen!":"Min. "+min+" Objekte auswählen!","error");return}
            return cb(ctx,sh.items);
        });
    }).catch(function(e){showStatus("Fehler: "+e.message,"error")});
}

// ===== RESIZE =====
function resize(dim,d){
    withShapes(1,function(ctx,items){
        items.forEach(function(s){s.load(["left","top","width","height"])});
        return ctx.sync().then(function(){
            var dp=c2p(Math.abs(d)),g=d>0;
            if(items.length===1){
                var s=items[0];
                if(dim==="width"||dim==="both"){var nw=g?s.width+dp:s.width-dp;if(nw>=c2p(MIN))s.width=nw}
                if(dim==="height"||dim==="both"){var nh=g?s.height+dp:s.height-dp;if(nh>=c2p(MIN))s.height=nh}
                return ctx.sync().then(function(){showStatus((dim==="both"?"W+H":dim==="width"?"W":"H")+(g?" +":"")+Math.abs(d).toFixed(2)+" cm","success")});
            }
            if(dim==="width"||dim==="both"){
                var hs=items.slice().sort(function(a,b){return a.left-b.left}),hg=[];
                for(var i=0;i<hs.length-1;i++)hg.push(hs[i+1].left-(hs[i].left+hs[i].width));
                var ok=true;for(var i=0;i<hs.length;i++)if((g?hs[i].width+dp:hs[i].width-dp)<c2p(MIN)){ok=false;break}
                if(ok){for(var i=0;i<hs.length;i++)hs[i].width=g?hs[i].width+dp:hs[i].width-dp;for(var i=1;i<hs.length;i++)hs[i].left=hs[i-1].left+hs[i-1].width+hg[i-1]}
            }
            if(dim==="height"||dim==="both"){
                var vs=items.slice().sort(function(a,b){return a.top-b.top}),vg=[];
                for(var i=0;i<vs.length-1;i++)vg.push(vs[i+1].top-(vs[i].top+vs[i].height));
                var ok=true;for(var i=0;i<vs.length;i++)if((g?vs[i].height+dp:vs[i].height-dp)<c2p(MIN)){ok=false;break}
                if(ok){for(var i=0;i<vs.length;i++)vs[i].height=g?vs[i].height+dp:vs[i].height-dp;for(var i=1;i<vs.length;i++)vs[i].top=vs[i-1].top+vs[i-1].height+vg[i-1]}
            }
            return ctx.sync().then(function(){showStatus((dim==="both"?"W+H":dim==="width"?"W":"H")+(g?" +":" −")+" Abstände OK","success")});
        });
    });
}

function propResize(d){
    withShapes(1,function(ctx,items){
        items.forEach(function(s){s.load(["left","top","width","height"])});
        return ctx.sync().then(function(){
            var dp=c2p(Math.abs(d)),g=d>0;
            if(items.length===1){
                var s=items[0],r=s.height/s.width,nw=g?s.width+dp:s.width-dp;
                if(nw>=c2p(MIN)){var nh=nw*r;if(nh>=c2p(MIN)){s.width=nw;s.height=nh}}
                return ctx.sync().then(function(){showStatus("Proportional "+(g?"+":"−")+Math.abs(d).toFixed(2)+" cm","success")});
            }
            var orig=items.map(function(s){return{shape:s,left:s.left,top:s.top,width:s.width,height:s.height,ratio:s.height/s.width}});
            var ok=true;orig.forEach(function(o){var nw=g?o.width+dp:o.width-dp;if(nw<c2p(MIN)||nw*o.ratio<c2p(MIN))ok=false});
            if(!ok){showStatus("Mindestgröße!","error");return ctx.sync()}
            var hs=orig.slice().sort(function(a,b){return a.left-b.left}),hg=[];
            for(var i=0;i<hs.length-1;i++)hg.push(hs[i+1].left-(hs[i].left+hs[i].width));
            var vs=orig.slice().sort(function(a,b){return a.top-b.top}),vg=[];
            for(var i=0;i<vs.length-1;i++)vg.push(vs[i+1].top-(vs[i].top+vs[i].height));
            orig.forEach(function(o){var nw=g?o.width+dp:o.width-dp;o.nw=nw;o.nh=nw*o.ratio;o.shape.width=nw;o.shape.height=o.nh});
            for(var i=1;i<hs.length;i++)hs[i].shape.left=hs[i-1].shape.left+hs[i-1].nw+hg[i-1];
            for(var i=1;i<vs.length;i++)vs[i].shape.top=vs[i-1].shape.top+vs[i-1].nh+vg[i-1];
            return ctx.sync().then(function(){showStatus("Proportional "+(g?"+":"−")+" Abstände OK","success")});
        });
    });
}

// ===== SNAP =====
function snap(mode){
    withShapes(1,function(ctx,items){
        items.forEach(function(s){s.load(["left","top","width","height"])});
        return ctx.sync().then(function(){
            items.forEach(function(s){
                if(mode==="position"||mode==="both"){s.left=c2p(rnd(p2c(s.left)));s.top=c2p(rnd(p2c(s.top)))}
                if(mode==="size"||mode==="both"){var nw=rnd(p2c(s.width)),nh=rnd(p2c(s.height));if(nw>=MIN)s.width=c2p(nw);if(nh>=MIN)s.height=c2p(nh)}
            });
            return ctx.sync().then(function(){showStatus("Am Raster eingerastet ✓","success")});
        });
    });
}

/* Group shapes into rows/columns by position similarity */
function groupByPosition(items,axis,tolerance){
    var groups=[],used={};
    var sorted=items.slice().sort(function(a,b){return axis==="y"?(a.top-b.top):(a.left-b.left)});
    for(var i=0;i<sorted.length;i++){
        if(used[i])continue;
        var grp=[sorted[i]];used[i]=true;
        var refPos=axis==="y"?sorted[i].top:sorted[i].left;
        for(var j=i+1;j<sorted.length;j++){
            if(used[j])continue;
            var pos=axis==="y"?sorted[j].top:sorted[j].left;
            if(Math.abs(pos-refPos)<=tolerance){grp.push(sorted[j]);used[j]=true}
        }
        groups.push(grp);
    }
    return groups;
}

function spacing(dir){
    withShapes(2,function(ctx,items){
        items.forEach(function(s){s.load(["left","top","width","height"])});
        return ctx.sync().then(function(){
            var sp=c2p(gridUnitCm);
            /* Tolerance: half a grid unit in points, min 5pt.
               Shapes within this Y/X range = same row/column */
            var tol=c2p(gridUnitCm)*0.5;
            if(tol<5)tol=5;
            if(dir==="horizontal"){
                /* Group by Y (= rows), then space each row horizontally */
                var rows=groupByPosition(items,"y",tol);
                rows.forEach(function(row){
                    if(row.length<2)return;
                    row.sort(function(a,b){return a.left-b.left});
                    for(var i=1;i<row.length;i++)row[i].left=row[i-1].left+row[i-1].width+sp;
                });
                var rc=rows.length;
                return ctx.sync().then(function(){showStatus("H-Abstand "+gridUnitCm.toFixed(2)+" cm ("+rc+" Zeile"+(rc>1?"n":"")+")" ,"success")});
            }else{
                /* Group by X (= columns), then space each column vertically */
                var cols=groupByPosition(items,"x",tol);
                cols.forEach(function(col){
                    if(col.length<2)return;
                    col.sort(function(a,b){return a.top-b.top});
                    for(var i=1;i<col.length;i++)col[i].top=col[i-1].top+col[i-1].height+sp;
                });
                var cc=cols.length;
                return ctx.sync().then(function(){showStatus("V-Abstand "+gridUnitCm.toFixed(2)+" cm ("+cc+" Spalte"+(cc>1?"n":"")+")" ,"success")});
            }
        });
    });
}

function shapeInfo(){
    withShapes(1,function(ctx,items){
        items.forEach(function(s){s.load(["name","left","top","width","height"])});
        return ctx.sync().then(function(){
            var el=document.getElementById("infoDisplay"),html="";
            items.forEach(function(s,idx){
                if(items.length>1)html+='<div style="font-weight:700;margin-top:'+(idx>0?'6':'0')+'px;color:#e94560">'+(s.name||'Obj '+(idx+1))+'</div>';
                html+='<div class="info-item"><span class="info-label">W:</span><span class="info-value">'+p2c(s.width).toFixed(2)+' cm</span></div>';
                html+='<div class="info-item"><span class="info-label">H:</span><span class="info-value">'+p2c(s.height).toFixed(2)+' cm</span></div>';
                html+='<div class="info-item"><span class="info-label">X:</span><span class="info-value">'+p2c(s.left).toFixed(2)+' cm</span></div>';
                html+='<div class="info-item"><span class="info-label">Y:</span><span class="info-value">'+p2c(s.top).toFixed(2)+' cm</span></div>';
            });
            el.innerHTML=html;el.classList.add("visible");
            showStatus("Info geladen ✓","info");
        });
    });
}

// ===== MATCH =====
function match(dim,mode){
    withShapes(2,function(ctx,items){
        items.forEach(function(s){s.load(["width","height"])});
        return ctx.sync().then(function(){
            var ws=items.map(function(s){return s.width}),hs=items.map(function(s){return s.height});
            var tw=mode==="max"?Math.max.apply(null,ws):Math.min.apply(null,ws);
            var th=mode==="max"?Math.max.apply(null,hs):Math.min.apply(null,hs);
            items.forEach(function(s){if(dim==="width"||dim==="both")s.width=tw;if(dim==="height"||dim==="both")s.height=th});
            return ctx.sync().then(function(){showStatus("Angeglichen → "+mode+" ✓","success")});
        });
    });
}

function propMatch(mode){
    withShapes(2,function(ctx,items){
        items.forEach(function(s){s.load(["width","height"])});
        return ctx.sync().then(function(){
            var ws=items.map(function(s){return s.width});
            var tw=mode==="max"?Math.max.apply(null,ws):Math.min.apply(null,ws);
            items.forEach(function(s){var r=s.height/s.width;s.width=tw;s.height=tw*r});
            return ctx.sync().then(function(){showStatus("Prop. angeglichen → "+mode+" ✓","success")});
        });
    });
}

// ===== TABLE =====
function createGridTable(){
    var cols=parseInt(document.getElementById("tableColumns").value);
    var rows=parseInt(document.getElementById("tableRows").value);
    var cw=parseFloat(document.getElementById("tableCellWidth").value);
    var ch=parseFloat(document.getElementById("tableCellHeight").value);
    if(isNaN(cols)||isNaN(rows)||cols<1||rows<1){showStatus("Ungültige Spalten/Zeilen!","error");return}
    if(isNaN(cw)||isNaN(ch)||cw<1||ch<1){showStatus("Ungültige Zellgröße!","error");return}
    if(cols>15){showStatus("Max 15 Spalten!","warning");return}
    if(rows>20){showStatus("Max 20 Zeilen!","warning");return}

    PowerPoint.run(function(ctx){
        var sel=ctx.presentation.getSelectedSlides();sel.load("items");
        return ctx.sync().then(function(){
            if(sel.items.length>0)return buildTable(ctx,sel.items[0],cols,rows,cw,ch);
            var slides=ctx.presentation.slides;slides.load("items");
            return ctx.sync().then(function(){
                if(!slides.items.length){showStatus("Keine Folie!","error");return ctx.sync()}
                return buildTable(ctx,slides.items[0],cols,rows,cw,ch);
            });
        });
    }).catch(function(e){showStatus("Fehler: "+e.message,"error")});
}

function buildTable(ctx,slide,cols,rows,cw,ch){
    var wCm=cw*gridUnitCm,hCm=ch*gridUnitCm,sp=gridUnitCm;
    var wPt=c2p(wCm),hPt=c2p(hCm),spPt=c2p(sp);
    var x0=c2p(8*gridUnitCm),y0=c2p(17*gridUnitCm);
    for(var r=0;r<rows;r++)for(var c=0;c<cols;c++){
        var s=slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);
        s.left=x0+(c*(wPt+spPt));s.top=y0+(r*(hPt+spPt));s.width=wPt;s.height=hPt;
        s.fill.setSolidColor("FFFFFF");s.lineFormat.color="808080";s.lineFormat.weight=0.3;
        s.name="TC_"+r+"_"+c;
    }
    return ctx.sync().then(function(){
        showStatus(cols+"×"+rows+" Tabelle erstellt ✓","success");
    });
}

// ===== SETUP =====
function setSlideSize(){
    var tw=785.5,th=547;
    PowerPoint.run(function(ctx){
        var ps=ctx.presentation.pageSetup;ps.load(["slideWidth","slideHeight"]);
        return ctx.sync().then(function(){ps.slideWidth=tw;return ctx.sync()}).then(function(){ps.slideHeight=th;return ctx.sync()}).then(function(){
            showStatus("Format: 27,711 × 19,297 cm ✓","success");
        });
    }).catch(function(e){showStatus("Fehler: "+e.message,"error")});
}

function toggleGuides(){
    PowerPoint.run(function(ctx){
        var masters=ctx.presentation.slideMasters;masters.load("items");
        return ctx.sync().then(function(){
            if(!masters.items.length){showStatus("Kein Master!","error");return ctx.sync()}
            var m0=masters.items[0],sh=m0.shapes;sh.load("items");
            return ctx.sync().then(function(){
                for(var i=0;i<sh.items.length;i++)sh.items[i].load("name");
                return ctx.sync().then(function(){
                    var existing=[];
                    for(var i=0;i<sh.items.length;i++)if(sh.items[i].name&&sh.items[i].name.indexOf(GTAG)===0)existing.push(sh.items[i]);
                    return existing.length>0?removeGuides(ctx,masters.items):addGuides(ctx,masters.items);
                });
            });
        });
    }).catch(function(e){showStatus("Fehler: "+e.message,"error")});
}

function addGuides(ctx,masters){
    var pos=[
        {t:"vertical",u:8},{t:"vertical",u:126},
        {t:"horizontal",u:5},{t:"horizontal",u:9},{t:"horizontal",u:15},{t:"horizontal",u:17},{t:"horizontal",u:86}
    ];
    var ps=ctx.presentation.pageSetup;ps.load(["slideWidth","slideHeight"]);
    return ctx.sync().then(function(){
        var sw=ps.slideWidth,sh=ps.slideHeight;
        masters.forEach(function(master){
            pos.forEach(function(g){
                var pt=Math.round(c2p(g.u*gridUnitCm)),s;
                if(g.t==="vertical"){s=master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);s.left=pt;s.top=0;s.width=1;s.height=sh}
                else{s=master.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle);s.left=0;s.top=pt;s.width=sw;s.height=1}
                s.name=GTAG+"_"+g.t+"_"+g.u;s.fill.setSolidColor("FF0000");s.lineFormat.visible=false;
            });
        });
        return ctx.sync().then(function(){showStatus("Hilfslinien eingeblendet ✓","success")});
    });
}

function removeGuides(ctx,masters){
    var proms=[];
    masters.forEach(function(master){
        var sh=master.shapes;sh.load("items");
        proms.push(ctx.sync().then(function(){
            for(var i=0;i<sh.items.length;i++)sh.items[i].load("name");
            return ctx.sync().then(function(){
                for(var i=0;i<sh.items.length;i++)if(sh.items[i].name&&sh.items[i].name.indexOf(GTAG)===0)sh.items[i].delete();
            });
        }));
    });
    return Promise.all(proms).then(function(){return ctx.sync().then(function(){showStatus("Hilfslinien entfernt ✓","success")})});
}

function copyShadow(){
    var t="Schatten-Standardwerte:\nFarbe: Schwarz\nTransparenz: 75 %\nGröße: 100 %\nWeichzeichnen: 4 pt\nWinkel: 90°\nAbstand: 1 pt";
    if(navigator.clipboard&&navigator.clipboard.writeText)navigator.clipboard.writeText(t).then(function(){showStatus("Kopiert ✓","success")}).catch(function(){showStatus("Kopieren fehlgeschlagen","error")});
    else showStatus("Zwischenablage nicht verfügbar","error");
}
