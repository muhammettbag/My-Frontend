const C=`
      .toast-notification {
        min-width: 350px;
        padding: 16px 20px;
        border-radius: 6px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
        font-size: 14px;
        font-weight: 500;
        animation: slideIn 0.3s ease-out;
      }
      
      .toast-notification.success {
        background-color: #28a745;
        color: white;
        border-left: 4px solid #207535;
      }
      
      .toast-notification.error {
        background-color: #dc3545;
        color: white;
        border-left: 4px solid #a02830;
      }
      
      .toast-notification-bar {
        margin-top: 8px;
        width: 100%;
        height: 3px;
        background-color: rgba(255, 255, 255, 0.3);
        border-radius: 2px;
        overflow: hidden;
      }
      
      .toast-notification-progress {
        height: 100%;
        background-color: rgba(255, 255, 255, 0.7);
        animation: progress 4s linear forwards;
      }
      
      @keyframes slideIn {
        from {
          transform: translateX(400px);
          opacity: 0;
        }
        to {
          transform: translateX(0);
          opacity: 1;
        }
      }
      
      @keyframes progress {
        from {
          width: 100%;
        }
        to {
          width: 0%;
        }
      }
    `,A=document.createElement("style");A.textContent=C;document.head.appendChild(A);const l=(n,e="success")=>{const t=document.getElementById("toast-container"),o=document.createElement("div");o.className=`toast-notification ${e}`;const r=document.createElement("div");r.textContent=n,r.style.marginBottom="8px";const a=document.createElement("div");a.className="toast-notification-bar";const i=document.createElement("div");i.className="toast-notification-progress",a.appendChild(i),o.appendChild(r),o.appendChild(a),t.appendChild(o),setTimeout(()=>{o.style.animation="slideOut 0.3s ease-out forwards",setTimeout(()=>{o.remove()},300)},4e3)};let s=1,f=10,p=1,h=null,T=[];const x=document.getElementById("makaleForm"),E=document.getElementById("makaleFormBaslik"),d=document.getElementById("makaleFormSubmitBtn"),v=document.getElementById("makaleTitle"),B=document.getElementById("makaleBadge"),$=document.getElementById("makaleContent"),L=document.getElementById("makaleImageUrl"),g=document.getElementById("makaleCategoryId"),k=document.getElementById("makaleAuthorId"),u=document.getElementById("makaleImportFileInput"),m=document.getElementById("makaleImportButton");let P=[],U=[];const z=`${API_BASE_URL}/Article/add`,H=`${API_BASE_URL}/Article/Update`,Y=`${API_BASE_URL}/Article/ExportArticles`,D=`${API_BASE_URL}/Article/Import`,G=async()=>{try{const n=await fetch(`${API_BASE_URL}/Category/GetAll`);if(!n.ok)throw new Error(`HTTP ${n.status}`);const e=await n.json();let t=[];Array.isArray(e)?t=e:e?.data&&Array.isArray(e.data)?t=e.data:e?.$values&&Array.isArray(e.$values)?t=e.$values:e?.items&&Array.isArray(e.items)?t=e.items:e?.Items&&Array.isArray(e.Items)&&(t=e.Items),t.length>0?(P=t,N(t),console.log("Kategoriler yüklendi:",t.length)):console.warn("Kategoriler yüklendi ama liste boş")}catch(n){console.error("Kategoriler yüklenemedi:",n)}},_=async()=>{try{const n=await fetch(`${API_BASE_URL}/Author/GetAll`);if(!n.ok)throw new Error(`HTTP ${n.status}`);const e=await n.json();let t=[];Array.isArray(e)?t=e:e?.data&&Array.isArray(e.data)?t=e.data:e?.$values&&Array.isArray(e.$values)?t=e.$values:e?.items&&Array.isArray(e.items)?t=e.items:e?.Items&&Array.isArray(e.Items)&&(t=e.Items),t.length>0?(U=t,R(t),console.log("Yazarlar yüklendi:",t.length)):console.warn("Yazarlar yüklendi ama liste boş")}catch(n){console.error("Yazarlar yüklenemedi:",n)}},N=n=>{const e=[];e.push('<option value="">Kategori Seçiniz</option>'),n.forEach((t,o)=>{const r=t.id||t.categoryId||t.Id||t.CategoryId||"",a=t.name||t.Name||t.title||t.Title||"Bilinmiyen Kategori";r&&e.push(`<option value="${r}">${o+1}. ${a}</option>`)}),g.innerHTML=e.join("")},R=n=>{const e=[];e.push('<option value="">Yazar Seçiniz</option>'),n.forEach((t,o)=>{const r=t.id||t.authorId||t.Id||t.AuthorId||"",a=t.name||t.Name||t.firstName||t.FirstName||t.fullName||t.FullName||"Bilinmiyen Yazar";r&&e.push(`<option value="${r}">${o+1}. ${a}</option>`)}),k.innerHTML=e.join("")},I=()=>{h=null,E&&(E.textContent="Makale Ekle"),d&&(d.textContent="Kaydet",d.classList.remove("btn-warning"),d.classList.add("btn-success")),g&&(g.value=""),k&&(k.value="")};window.makaleleriExcelIndir=async function(){const n=document.getElementById("makaleExportButton");n&&(n.disabled=!0,n.textContent="İndiriliyor...");try{const e=await fetch(Y,{method:"GET"});if(!e.ok){const a=await e.text();throw new Error(`HTTP ${e.status}: ${a||e.statusText}`)}const t=await e.blob();if(!t||t.size===0)throw new Error("Boş dosya döndü.");const o=URL.createObjectURL(t),r=document.createElement("a");r.href=o,r.download="Makaleler.xlsx",document.body.appendChild(r),r.click(),r.remove(),URL.revokeObjectURL(o),l("Makaleler Excel olarak indirildi.","success")}catch(e){console.error("Makale Excel indirme hatası:",e),l(`Excel indirilemedi: ${e.message}`,"error")}finally{n&&(n.disabled=!1,n.textContent="Excel İndir")}};window.makaleleriYenile=function(){G(),_(),makaleleriGetir(s)};window.makaleExcelSec=function(){if(!u){l("Dosya seçici bulunamadı.","error");return}u.click()};window.makaleExcelYukle=async function(n){if(!n){l("Lütfen bir Excel dosyası seçin.","error");return}const e=(n.name||"").toLowerCase();if(!e.endsWith(".xlsx")&&!e.endsWith(".xls")){l("Sadece .xlsx veya .xls dosyası yükleyebilirsiniz.","error");return}m&&(m.disabled=!0,m.textContent="Yükleniyor...");try{const t=new FormData;t.append("file",n);const o=await fetch(D,{method:"POST",body:t}),r=await o.text();if(!o.ok)throw new Error(r||`HTTP ${o.status}: ${o.statusText}`);l(r||"Makaleler başarıyla içeri aktarıldı.","success"),makaleleriYenile()}catch(t){console.error("Makale Excel içeri aktarma hatası:",t),l(`Excel yüklenemedi: ${t.message}`,"error")}finally{m&&(m.disabled=!1,m.textContent="Excel Yükle"),u&&(u.value="")}};u&&u.addEventListener("change",n=>{const e=n.target?.files?.[0];e&&makaleExcelYukle(e)});document.addEventListener("DOMContentLoaded",()=>{makaleleriYenile()});x.addEventListener("submit",async n=>{n.preventDefault();const e={title:v.value.trim(),badge:B.value.trim()||null,content:$.value.trim(),imageUrl:L.value.trim(),categoryId:g.value,authorId:k.value};if(!e.title){l("Başlık zorunludur.","error");return}if(!e.imageUrl){l("Görsel URL zorunludur.","error");return}if(!e.categoryId){l("Kategori zorunludur.","error");return}if(!e.authorId){l("Yazar zorunludur.","error");return}const t=!!h,o=t?{...e,id:h}:e,r=t?H:z,a=t?"PUT":"POST";try{const i=await fetch(r,{method:a,headers:{"Content-Type":"application/json"},body:JSON.stringify(o)});if(!i.ok){const c=await i.text();throw new Error(`HTTP ${i.status}: ${c||i.statusText}`)}l(t?"Makale başarıyla güncellendi!":"Makale başarıyla eklendi!","success"),x.reset(),I(),listeyiGoster()}catch(i){console.error("Makale kaydetme hatası:",i),l(`Makale kaydedilemedi: ${i.message}`,"error")}});window.formuGoster=function(){document.getElementById("makale-liste-view").style.display="none",document.getElementById("makale-form-view").style.display="block",document.getElementById("makaleForm").reset(),I()};window.listeyiGoster=function(){document.getElementById("makale-form-view").style.display="none",document.getElementById("makale-liste-view").style.display="block",I(),makaleleriGetir(s)};window.makaleDuzenle=function(n){const e=T.find(r=>{const a=r?.id||r?.articleId||r?.Id||r?.ArticleId;return String(a)===String(n)});if(!e){l("Düzenlenecek makale bulunamadı.","error");return}h=e.id||e.articleId||e.Id||e.ArticleId,v.value=e.title||e.Title||"",B.value=e.badge||e.Badge||"",$.value=e.content||e.Content||"",L.value=e.imageUrl||e.ImageUrl||"";const t=e.categoryId||e.CategoryId,o=e.authorId||e.AuthorId;t&&(g.value=t),o&&(k.value=o),E&&(E.textContent="Makale Düzenle"),d&&(d.textContent="Güncelle",d.classList.remove("btn-success"),d.classList.add("btn-warning")),document.getElementById("makale-liste-view").style.display="none",document.getElementById("makale-form-view").style.display="block"};window.makalaSil=async function(n){if(confirm("Bu makaleyi silmek istediğinizden emin misiniz?"))try{const t=await fetch(`${API_BASE_URL}/Article/Delete/${n}`,{method:"DELETE",headers:{"Content-Type":"application/json"}});if(!t.ok){const o=await t.text();throw new Error(`HTTP ${t.status}: ${o||t.statusText}`)}l("Makale başarıyla silindi!","success"),makaleleriGetir(s)}catch(t){console.error("Makale silme hatası:",t),l(`Makale silinemedi: ${t.message}`,"error")}};window.sayfayaGit=function(n){n<1||n>p||(s=n,makaleleriGetir(s))};const b=()=>{const n=document.getElementById("makalePagination");if(!n||p<=1){n&&(n.innerHTML="");return}let e="";for(let t=1;t<=p;t++)e+=`
          <li class="page-item ${t===s?"active":""}">
          <a class="page-link" href="#" onclick="sayfayaGit(${t}); return false;">${t}</a>
          </li>
        `;n.innerHTML=e};window.makaleleriGetir=function(n=s){s=n;const e=document.getElementById("makaleTablosu"),t=document.getElementById("makalePagination");e.innerHTML='<tr><td colspan="4" class="text-center">Yükleniyor...</td></tr>',t&&(t.innerHTML="");const o=[`${API_BASE_URL}/Article/GetAll`],r=async a=>{const i=await fetch(`${a}?page=${s}&pageSize=${f}`);if(i.status===404)return null;if(!i.ok){const c=await i.text();throw new Error(`HTTP ${i.status}: ${c||i.statusText}`)}const y=await i.text();if(!y)return[];try{return JSON.parse(y)}catch{throw new Error("API JSON döndürmedi. Yanıt formatını kontrol et.")}};(async()=>{for(const a of o){const i=await r(a);if(i!==null)return i}throw new Error("HTTP 404: Uygun liste endpointi bulunamadı.")})().then(a=>{e.innerHTML="";const i=Array.isArray(a)?a:a?.items||a?.Items||[];T=i,f=a?.pageSize||a?.PageSize||f,s=a?.page||a?.Page||s;const y=a?.totalCount||a?.TotalCount||i.length;if(p=a?.totalPages||a?.TotalPages||Math.max(1,Math.ceil(y/f)),!i||i.length===0){e.innerHTML='<tr><td colspan="4" class="text-center">Henüz makale yok.</td></tr>',b();return}i.forEach((c,M)=>{const S=(s-1)*f+M+1,w=c.id||c.articleId||c.Id||c.ArticleId||"";e.innerHTML+=`
                        <tr>
                <td>${S}</td>
                            <td>${c.title}</td>
                            <td>${c.badge||"-"}</td>
                            <td>
                        <button class="btn btn-warning btn-sm" onclick="makaleDuzenle('${w}')">Düzenle</button>
                                <button class="btn btn-danger btn-sm" onclick="makalaSil('${w}')">Sil</button>
                            </td>
                        </tr>
                    `}),b()}).catch(a=>{console.error("Hata:",a),p=1,e.innerHTML=`<tr><td colspan="4" class="text-danger text-center">Makale listesi alınamadı: ${a.message}</td></tr>`})};
