(this["webpackJsonpexcel-to-word-converter-frontend"]=this["webpackJsonpexcel-to-word-converter-frontend"]||[]).push([[0],{98:function(e,n,t){"use strict";t.r(n);var c=t(0),a=t.n(c),o=t(8),r=t.n(o),s=t(134),l=t(135),i=t(139),d=t(137),p=t(130),j=t(136),u=t(141),b=t(127),h=t(138),m=t(10);const x=Object(b.a)((e=>({input:{display:"none"}})));var O=function(e){let{onFileChange:n}=e;const t=x();return Object(m.jsxs)("div",{children:[Object(m.jsx)("input",{accept:".xlsx,.xls",className:t.input,id:"contained-button-file",type:"file",onChange:e=>{const t=e.target.files[0];n(t)}}),Object(m.jsx)("label",{htmlFor:"contained-button-file",children:Object(m.jsx)(p.a,{variant:"contained",color:"default",component:"span",children:"\u0412\u044b\u0431\u0440\u0430\u0442\u044c Excel \u0444\u0430\u0439\u043b"})})]})},g=t(61),f=t.n(g);const v=Object(b.a)((e=>({container:{marginTop:e.spacing(4),display:"flex",flexDirection:"column",alignItems:"center"},button:{marginTop:e.spacing(2)},checkbox:{marginTop:e.spacing(2)},discountInput:{marginTop:e.spacing(2),width:"100%"}})));function w(e){return Object(m.jsx)(h.a,{elevation:6,variant:"filled",...e})}var y=function(){const e=v(),[n,t]=Object(c.useState)(null),[a,o]=Object(c.useState)(!1),[r,b]=Object(c.useState)(null),[h,x]=Object(c.useState)(!1),[g,y]=Object(c.useState)(""),[k,C]=Object(c.useState)(!1),S=(e,n)=>{"clickaway"!==n&&b(null)};return Object(m.jsxs)(s.a,{className:e.container,children:[Object(m.jsx)(l.a,{variant:"h4",gutterBottom:!0,children:"\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0435\u0440 Excel \u0432 Word"}),Object(m.jsx)(O,{onFileChange:e=>{console.log("\u0424\u0430\u0439\u043b \u0432\u044b\u0431\u0440\u0430\u043d:",e.name),t(e)}}),Object(m.jsxs)("div",{className:e.checkbox,children:[Object(m.jsx)(i.a,{checked:h,onChange:e=>x(e.target.checked),color:"primary"}),Object(m.jsx)(l.a,{component:"span",children:"\u0414\u043e\u0431\u0430\u0432\u0438\u0442\u044c \u0441\u043a\u0438\u0434\u043a\u0443"})]}),h&&Object(m.jsx)(d.a,{className:e.discountInput,label:"\u041f\u0440\u043e\u0446\u0435\u043d\u0442 \u0441\u043a\u0438\u0434\u043a\u0438",type:"number",value:g,onChange:e=>y(e.target.value),variant:"outlined",size:"small"}),Object(m.jsxs)("div",{className:e.checkbox,children:[Object(m.jsx)(i.a,{checked:k,onChange:e=>C(e.target.checked),color:"primary"}),Object(m.jsx)(l.a,{component:"span",children:"\u0421\u0434\u0435\u043b\u0430\u0442\u044c \u0441\u043e\u043a\u0440\u0430\u0449. \u041a\u041f"})]}),Object(m.jsx)(p.a,{variant:"contained",color:"primary",onClick:async()=>{if(n){o(!0),b(null);try{console.log("\u041d\u0430\u0447\u0430\u043b\u043e \u043a\u043e\u043d\u0432\u0435\u0440\u0442\u0430\u0446\u0438\u0438 \u0444\u0430\u0439\u043b\u0430:",n.name);const e=await(async(e,n,t)=>{const c=new FormData;c.append("file",e),c.append("originalFileName",e.name),null!==n&&c.append("discountPercentage",n),c.append("makeShortVersion",t);try{return await f.a.post("https://\u0430\u0440\u0445\u0438\u043e-\u043a\u043e\u043c\u043c\u0435\u0440\u0447\u0435\u0441\u043a\u043e\u0435.\u0440\u0444/api/convert",c,{headers:{"Content-Type":"multipart/form-data"},responseType:"arraybuffer",withCredentials:!0})}catch(r){if(r.response){const n=(new TextDecoder).decode(r.response.data);throw new Error(n)}throw r}})(n,h?g:null,k);console.log("\u041e\u0442\u0432\u0435\u0442 \u043f\u043e\u043b\u0443\u0447\u0435\u043d:",e);const t=new Blob([e.data],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}),c=window.URL.createObjectURL(t),a=document.createElement("a");a.href=c,a.download="converted.docx",a.click(),window.URL.revokeObjectURL(c),console.log("\u0424\u0430\u0439\u043b \u0443\u0441\u043f\u0435\u0448\u043d\u043e \u0441\u043a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u043d \u0438 \u0441\u043a\u0430\u0447\u0430\u043d")}catch(r){var e;console.error("\u041e\u0448\u0438\u0431\u043a\u0430 \u043f\u0440\u0438 \u043a\u043e\u043d\u0432\u0435\u0440\u0442\u0430\u0446\u0438\u0438:",r),b((null===(e=r.response)||void 0===e?void 0:e.data)||"\u041f\u0440\u043e\u0438\u0437\u043e\u0448\u043b\u0430 \u043e\u0448\u0438\u0431\u043a\u0430 \u043f\u0440\u0438 \u043a\u043e\u043d\u0432\u0435\u0440\u0442\u0430\u0446\u0438\u0438 \u0444\u0430\u0439\u043b\u0430")}finally{o(!1)}}else b("\u041f\u043e\u0436\u0430\u043b\u0443\u0439\u0441\u0442\u0430, \u0432\u044b\u0431\u0435\u0440\u0438\u0442\u0435 \u0444\u0430\u0439\u043b Excel")},disabled:!n||a,className:e.button,children:a?Object(m.jsx)(j.a,{size:24}):"\u041a\u043e\u043d\u0432\u0435\u0440\u0442\u0438\u0440\u043e\u0432\u0430\u0442\u044c"}),Object(m.jsx)(u.a,{open:!!r,autoHideDuration:6e3,onClose:S,children:Object(m.jsx)(w,{onClose:S,severity:"error",children:r})})]})};r.a.render(Object(m.jsx)(a.a.StrictMode,{children:Object(m.jsx)(y,{})}),document.getElementById("root"))}},[[98,1,2]]]);
//# sourceMappingURL=main.e5d1523d.chunk.js.map