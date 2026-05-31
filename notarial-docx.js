/* Oficial Notarial IA — generador de Word (.docx) compartido.
   Uso:  NotarialDocx.download({title:'DATOS PARA ...', filename:'Datos_...'}, bodyHtml)
   - Si la librería docx (window.docx) está cargada -> .docx con estilo.
   - Si no -> cae a HTML compatible con Word (.doc). El botón nunca se rompe. */
(function(){
  var NAVY="1A3A5C", GOLD="C9A84C", INK="2C2C2C", GREY="6B6258";

  function showOk(){
    var ok=document.getElementById('ok');
    if(ok){ ok.style.display='block'; ok.scrollIntoView({behavior:'smooth'}); }
  }
  function trigger(blob, name){
    var a=document.createElement('a');
    a.href=URL.createObjectURL(blob); a.download=name; a.click();
    showOk();
  }

  // ---- FALLBACK: HTML compatible con Word ----
  function legacy(meta, bodyHtml){
    var w='<html xmlns:o=\'urn:schemas-microsoft-com:office:office\' xmlns:w=\'urn:schemas-microsoft-com:office:word\'>'+
      '<head><meta charset=\'utf-8\'><style>body{font-family:\'Times New Roman\',serif;font-size:12pt;line-height:1.9;margin:3cm}'+
      'h1{text-align:center;font-size:14pt;margin-bottom:20pt}h2{font-size:12pt;margin-top:16pt;border-bottom:1px solid #aaa;padding-bottom:3pt}p{margin:3pt 0}'+
      '</style></head><body><h1>'+(meta.top||'NOTARÍA')+'<br>'+meta.title+'</h1>'+bodyHtml+'</body></html>';
    trigger(new Blob([w],{type:'application/msword'}), meta.filename+'.doc');
  }

  // ---- PARSER HTML -> bloques ----
  function alertType(styleStr){
    var s=(styleStr||'').toLowerCase();
    if(s.indexOf('c0392b')>-1 || s.indexOf('fff0ee')>-1) return 'red';
    if(s.indexOf('e8f0ff')>-1 || s.indexOf('90a8e8')>-1) return 'blue';
    return 'amber';
  }
  function isSubHeading(p){
    var st=p.querySelector('strong');
    if(!st) return false;
    var t=p.textContent.trim();
    return t===st.textContent.trim() && /:\s*$/.test(t);
  }
  function isFooter(p){
    var s=(p.getAttribute('style')||'').toLowerCase();
    return s.indexOf('#777')>-1 || /documento generado/i.test(p.textContent);
  }
  function kv(p){
    // devuelve [label,value] o null
    var span=p.querySelector('span');
    var strong=p.querySelector('strong');
    var label='', value='';
    if(span && /:\s*$/.test(span.textContent.trim())){           // cluster-2: <span>Label:</span> value
      label=span.textContent.replace(/:\s*$/,'').trim();
      value=p.textContent.replace(span.textContent,'').replace(/\u00a0/g,' ').trim();
    } else if(strong){                                            // cluster-1: Label: <strong>Value</strong>
      value=strong.textContent.trim();
      label=p.textContent.replace(strong.textContent,'').replace(/:\s*$/,'').trim();
    } else {
      var txt=p.textContent.replace(/\u00a0/g,' ').trim(); if(!txt) return null;
      var i=txt.indexOf(':');
      if(i>-1){ label=txt.slice(0,i).trim(); value=txt.slice(i+1).trim(); }
      else { label=''; value=txt; }
    }
    if(!value) return null;
    return [label, value];
  }

  function build(meta, bodyHtml){
    var D=window.docx;
    var doc=new D.Document({
      creator:"Oficial Notarial IA",
      styles:{ default:{ document:{ run:{ font:"Calibri", size:19, color:INK } } } },
      sections:[{
        properties:{ page:{ margin:{ top:1418, bottom:1418, left:1418, right:1418 } } },
        footers:{ default:new D.Footer({ children:[ new D.Paragraph({ alignment:D.AlignmentType.CENTER,
          children:[ new D.TextRun({text:"Documento generado por Oficial Notarial IA · "+new Date().toLocaleDateString('es-ES')+"   ·   Página ",size:15,color:GREY,font:"Calibri"}),
                     new D.TextRun({children:[D.PageNumber.CURRENT],size:15,color:GREY,font:"Calibri"}) ] }) ] }) },
        children: blocks(meta, bodyHtml, D)
      }]
    });
    D.Packer.toBlob(doc).then(function(b){ trigger(b, meta.filename+'.docx'); });
  }

  function blocks(meta, bodyHtml, D){
    var out=[];
    out.push(new D.Paragraph({alignment:D.AlignmentType.CENTER,spacing:{after:40},
      children:[new D.TextRun({text:(meta.top||"NOTARÍA"),bold:true,color:NAVY,size:30,font:"Cambria",characterSpacing:60})]}));
    out.push(new D.Paragraph({alignment:D.AlignmentType.CENTER,spacing:{after:80},
      children:[new D.TextRun({text:meta.title,bold:true,color:GOLD,size:22,font:"Cambria"})]}));
    out.push(new D.Paragraph({alignment:D.AlignmentType.JUSTIFIED,spacing:{after:220},
      children:[new D.TextRun({text:"Documento preparado a partir de los datos aportados por el interesado para la preparación del expediente notarial. Verifíquese su exactitud antes de la firma.",italics:true,color:GREY,size:18,font:"Calibri"})]}));

    var rows=[];
    function flush(){
      if(!rows.length) return;
      var noB={style:D.BorderStyle.NONE,size:0,color:"FFFFFF"};
      var hair={style:D.BorderStyle.SINGLE,size:2,color:"E0D8CC"};
      out.push(new D.Table({ width:{size:100,type:D.WidthType.PERCENTAGE},
        borders:{top:noB,left:noB,right:noB,bottom:hair,insideHorizontal:hair,insideVertical:noB},
        rows: rows.map(function(r){
          var cells=[];
          if(r[0]){
            cells.push(new D.TableCell({width:{size:38,type:D.WidthType.PERCENTAGE},margins:{top:60,bottom:60,left:80,right:120},
              children:[new D.Paragraph({children:[new D.TextRun({text:r[0],color:GREY,size:19,font:"Calibri"})]})]}));
            cells.push(new D.TableCell({width:{size:62,type:D.WidthType.PERCENTAGE},margins:{top:60,bottom:60,left:120,right:80},
              children:[new D.Paragraph({children:[new D.TextRun({text:r[1],bold:true,color:INK,size:19,font:"Calibri"})]})]}));
          } else {
            cells.push(new D.TableCell({columnSpan:2,margins:{top:60,bottom:60,left:80,right:80},
              children:[new D.Paragraph({children:[new D.TextRun({text:r[1],color:INK,size:19,font:"Calibri"})]})]}));
          }
          return new D.TableRow({children:cells});
        })
      }));
      rows=[];
    }
    function heading(t){ flush(); out.push(new D.Paragraph({spacing:{before:280,after:120},
      border:{bottom:{color:GOLD,style:D.BorderStyle.SINGLE,size:8,space:4}},
      children:[new D.TextRun({text:t.toUpperCase(),bold:true,color:NAVY,size:22,font:"Cambria",characterSpacing:16})]})); }
    function sub(t){ flush(); out.push(new D.Paragraph({spacing:{before:150,after:60},
      children:[new D.TextRun({text:t,bold:true,color:NAVY,size:20,font:"Calibri"})]})); }
    function alert(text,type){ flush();
      var c={red:["C0392B","FDECEA"],amber:["C9A84C","FBF5E6"],blue:["6F8FD6","EAF0FB"]}[type]||["C9A84C","FBF5E6"];
      out.push(new D.Paragraph({ spacing:{before:80,after:80}, indent:{left:120},
        shading:{ fill:c[1] },
        border:{ left:{ style:D.BorderStyle.SINGLE, size:18, color:c[0], space:10 } },
        children:[new D.TextRun({text:text,size:18,color:INK,font:"Calibri"})] })); }

    var dom=new (window.DOMParser)().parseFromString('<body>'+bodyHtml+'</body>','text/html');
    var nodes=dom.body.childNodes;
    for(var i=0;i<nodes.length;i++){
      var n=nodes[i];
      if(n.nodeType!==1) continue; // solo elementos
      var tag=n.tagName.toUpperCase();
      if(tag==='H2'){ heading(n.textContent.trim()); }
      else if(tag==='DIV'){ alert(n.textContent.replace(/\s+/g,' ').trim(), alertType(n.getAttribute('style'))); }
      else if(tag==='P'){
        if(isFooter(n)) continue;
        if(isSubHeading(n)){ sub(n.textContent.replace(/:\s*$/,'').trim()); continue; }
        var r=kv(n); if(r) rows.push(r);
      }
      // BR y otros: ignorar
    }
    flush();
    return out;
  }

  window.NotarialDocx={ download:function(meta, bodyHtml){
    try{
      if(window.docx && window.docx.Packer){ build(meta, bodyHtml); }
      else { legacy(meta, bodyHtml); }
    }catch(e){ console.error('NotarialDocx fallo, uso fallback:',e); legacy(meta, bodyHtml); }
  }};
})();
