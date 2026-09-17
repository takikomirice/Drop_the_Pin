// Only the selected background is attached. Extending this catalog does not
// change pin storage, the map center, or the zoom level.
export const backgrounds = [
  {id:'osm',name:'通常地図',scope:'世界',attribution:'© <a href="https://www.openstreetmap.org/copyright" target="_blank" rel="noopener">OpenStreetMap</a> contributors',url:'https://tile.openstreetmap.org/{z}/{x}/{y}.png',maxNativeZoom:19},
  {id:'japan-terrain',name:'陰影起伏図',scope:'世界',attribution:'<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>（陰影起伏図）',url:'https://cyberjapandata.gsi.go.jp/xyz/hillshademap/{z}/{x}/{y}.png',minZoom:2,maxNativeZoom:16,fallback:'terrain',note:'日本の山や谷の凹凸を見る地図。範囲外は世界版。拡大しすぎると粗くなります。'}
];
// The world relief is an underlay, not a separate choice in the menu.
const underlays = [
  {id:'terrain',attribution:'<a href="https://maps.gsi.go.jp/development/ichiran.html" target="_blank" rel="noopener">地理院タイル</a>（陰影起伏図・全球版）',url:'https://cyberjapandata.gsi.go.jp/xyz/earthhillshade/{z}/{x}/{y}.png',maxNativeZoom:8}
];

function element(tag, className, text) {
  const el=document.createElement(tag);
  if(className) el.className=className;
  if(text!=null) el.textContent=text;
  return el;
}
function button(label,className) {
  const el=element('button',className,label); el.type='button'; return el;
}
export function attach(map, config={}) {
  const L=window.L;
  let active=null, pending=null, timeout=null, generation=0;
  let popup=null, disposeCard=null, revealGeneration=0;
  const pins=L.markerClusterGroup({
    maxClusterRadius:45,showCoverageOnHover:false,animate:false,
    spiderfyOnMaxZoom:true,spiderfyDistanceMultiplier:1.6,
    iconCreateFunction(cluster) {
      const count=cluster.getChildCount();
      return L.divIcon({className:'dtp-cluster',iconSize:[44,44],html:'<span aria-label="'+count+'件のピン。拡大して表示">'+count+'</span>'});
    }
  }).addTo(map);
  const toolbar=document.getElementById('map-toolbar');
  const toggle=button('地図 ▾','dtp-map-toggle');
  toggle.setAttribute('aria-label','地図を切り替え');
  toggle.setAttribute('aria-expanded','false');
  toggle.setAttribute('aria-controls','background-map-options');
  const menu=element('div','dtp-map-control');menu.id='background-map-options';menu.hidden=true;
  toolbar.append(toggle,menu);
  function closeMenu(restoreFocus=false) {
    menu.hidden=true;toggle.setAttribute('aria-expanded','false');
    if(restoreFocus)toggle.focus();
  }
  toggle.onclick=()=>{menu.hidden=!menu.hidden;toggle.setAttribute('aria-expanded',String(!menu.hidden));};
  function outsideClick(event) {if(!toolbar.contains(event.target))closeMenu();}
  function escapeMenu(event) {if(event.key==='Escape'&&!menu.hidden){event.stopPropagation();closeMenu(true);}}
  document.addEventListener('pointerdown',outsideClick);
  toolbar.addEventListener('keydown',escapeMenu);
  let select,status,retry;
  {
    select=element('select'); select.setAttribute('aria-label','背景地図');
    backgrounds.forEach(bg=>select.append(new Option(bg.name,bg.id)));
    status=element('div','dtp-basemap-status');status.setAttribute('role','status');
    status.id='background-map-status';select.setAttribute('aria-describedby',status.id);
    retry=button('再試行','dtp-map-retry');retry.hidden=true;
    retry.onclick=()=>switchBackground(select.value);
    select.onchange=()=>switchBackground(select.value);
    menu.append(select,status,retry);
  }
  L.control.scale({position:'bottomright',imperial:false}).addTo(map);
  const mapElement=map.getContainer();
  const panel=document.getElementById(config.panelId||'side-panel');
  function bottomInset() {
    if(!panel) return 0;
    const area=mapElement.getBoundingClientRect(), box=panel.getBoundingClientRect();
    return box.width>=area.width-2 && box.left<=area.left+2
      ? Math.max(0,Math.min(area.height,area.bottom-box.top)) : 0;
  }
  function positionControls() {
    mapElement.querySelectorAll('.leaflet-bottom').forEach(corner=>{corner.style.bottom=bottomInset()+'px';});
  }
  const resizeObserver=new ResizeObserver(positionControls);
  resizeObserver.observe(mapElement);if(panel)resizeObserver.observe(panel);
  const classObserver=new MutationObserver(positionControls);
  classObserver.observe(document.body,{attributes:true,attributeFilter:['class']});
  if(panel)panel.addEventListener('transitionend',positionControls);
  positionControls();
  function cancelPending() {
    clearTimeout(timeout); timeout=null;
    if(pending) {pending.watch.off();map.removeLayer(pending.layer);pending=null;}
  }
  function switchBackground(id) {
    const bg=backgrounds.find(item=>item.id===id);
    if(!bg) return;
    const request=++generation;
    cancelPending();retry.hidden=true;
    select.value=id;
    if(active && active.id===id) {status.textContent=bg.note||bg.name;return;}
    function makeLayer(item,options={}) {
      return L.tileLayer(item.url,{attribution:item.attribution,maxZoom:19,minZoom:item.minZoom||0,maxNativeZoom:item.maxNativeZoom,...options});
    }
    const japanBounds=L.latLngBounds([[20,122],[46,154]]);
    let layer=makeLayer(bg,bg.fallback?{bounds:japanBounds}:{}),watch=layer;
    if(bg.fallback) {
      const fallback=makeLayer(underlays.find(item=>item.id===bg.fallback));
      layer=L.layerGroup([fallback,layer]);
      if(map.getZoom()<bg.minZoom||!map.getBounds().overlaps(japanBounds))watch=fallback;
    }
    pending={id,layer,watch};let failed=false,loaded=0;
    status.textContent=bg.name+'を読み込み中…';
    function fail() {
      if(request!==generation || !pending) return;
      cancelPending(); status.textContent=bg.name+'を読み込めません。'+(active?'元の地図を表示しています。':'再試行してください。');retry.hidden=false;
    }
    function ready() {
      if(request!==generation || !pending) return;
      if(failed && (!bg.fallback||!loaded)) {fail();return;}
      clearTimeout(timeout);watch.off('load',ready);watch.off('tileerror',onError);watch.off('tileload',onTileLoad);
      if(active) map.removeLayer(active.layer);
      active={id,layer};pending=null;
      status.textContent=bg.name+(bg.note?' · '+bg.note:'');
      toggle.title=bg.name;
    }
    function onError() {failed=true;}
    function onTileLoad() {loaded++;}
    watch.on('load',ready);watch.on('tileerror',onError);watch.on('tileload',onTileLoad);timeout=setTimeout(fail,12000);
    layer.addTo(map);
    if(!watch.isLoading()) ready();
  }
  function closePin() {
    revealGeneration++;
    if(popup) map.closePopup(popup);
  }
  function openPin(pin,marker) {
    if(!pin) return false;
    closePin();
    const request=++revealGeneration;
    const hasMarker=!!marker&&pins.hasLayer(marker);
    const located=pin.lat!=null&&pin.lng!=null&&Number.isFinite(Number(pin.lat))&&Number.isFinite(Number(pin.lng));
    const position=hasMarker?marker.getLatLng():(located?[Number(pin.lat),Number(pin.lng)]:map.getCenter());
    const reveal=()=>{
      if(request!==revealGeneration || (hasMarker&&!pins.hasLayer(marker))) return;
      if(hasMarker) marker.closeTooltip();
      if(config.beforeOpen) config.beforeOpen();
      const card=element('article','dtp-pin-card');
      const title=String(pin.title||'').trim()||'(無題)';
      const media=element('div','dtp-pin-media');
      const audioOnly=pin.hasAudio&&!(pin.fileId||pin.imageUrl);
      if(audioOnly) media.classList.add('dtp-pin-media-audio');
      const photoButton=button('','dtp-pin-photo');photoButton.setAttribute('aria-label','写真を拡大');photoButton.disabled=true;
      const img=element('img','protected-photo');img.alt=title+'の写真';img.draggable=false;img.hidden=true;photoButton.append(img);
      const photoStatus=element('div','dtp-photo-status','写真を読み込み中…');photoStatus.setAttribute('role','status');
      const retryPhoto=button('写真を再読み込み','dtp-photo-retry');retryPhoto.hidden=true;
      const infoButton=button('ⓘ','dtp-info-toggle');infoButton.setAttribute('aria-label','詳細情報');infoButton.setAttribute('aria-expanded','false');
      const info=element('div','dtp-pin-info');info.hidden=true;
      const description=element('p','',pin.description||'説明はありません。');
      info.append(element('h3','dtp-info-title',title),description);
      if(!located) info.append(element('p','dtp-pin-unplaced','未配置（場所はまだ指定されていません）'));
      if(config.timeText&&!config.mountDetails) info.append(element('p','dtp-info-time',config.timeText(pin)));
      if(pin.status) info.append(element('p','',pin.status));
      if(Array.isArray(pin.tags)&&pin.tags.length) info.append(element('p','',pin.tags.map(tag=>'#'+tag).join(' ')));
      (Array.isArray(pin.links)?pin.links:[]).forEach(value=>{
        try {
          const url=new URL(value);if(!['https:','http:'].includes(url.protocol)) return;
          const a=element('a','',value);a.href=url.href;a.target='_blank';a.rel='noopener noreferrer';info.append(a);
        } catch(_) {}
      });
      infoButton.onclick=()=>{
        const expanded=info.hidden;info.hidden=!expanded;
        infoButton.setAttribute('aria-expanded',String(expanded));
        infoButton.setAttribute('aria-label',expanded?'写真に戻る':'詳細情報');
        infoButton.textContent=expanded?'←':'ⓘ';photoButton.tabIndex=expanded?-1:0;
        if(audioOnly) {media.classList.toggle('is-expanded',expanded);if(popup) popup.update();}
      };
      let source='',loader=null;
      function showSource(url) {
        source=url;img.onload=()=>{img.hidden=false;photoStatus.hidden=true;photoButton.disabled=false;};
        img.onerror=()=>{photoButton.disabled=true;photoStatus.hidden=false;photoStatus.textContent='写真を表示できませんでした';retryPhoto.hidden=false;};
        img.src=url;
      }
      function renderPhoto(view) {
        if(view.status==='ready') {showSource(view.objectUrl);return;}
        source='';photoButton.disabled=true;img.hidden=true;photoStatus.hidden=false;
        photoStatus.textContent=view.status==='error'?'写真を表示できませんでした':'写真を読み込み中…';
        retryPhoto.hidden=view.status!=='error';
      }
      photoButton.onclick=()=>{if(source) config.openPhoto(source,title,photoButton);};
      media.append(photoButton,photoStatus,retryPhoto,info,infoButton);
      card.append(media,element('div','dtp-pin-title',(located?'':'未配置 · ')+title));
      const audioSlot=element('div','dtp-pin-audio');audioSlot.hidden=!pin.hasAudio;card.append(audioSlot);
      let unmountDetails=null;
      L.DomEvent.disableClickPropagation(card);L.DomEvent.disableScrollPropagation(card);
      popup=L.popup({className:'dtp-photo-popup',maxWidth:300,minWidth:220,autoPanPaddingTopLeft:[16,100],autoPanPaddingBottomRight:[16,bottomInset()+24],offset:[0,-32]}).setLatLng(position).setContent(card);
      const thisPopup=popup;
      const cleanup=()=>{
        if(config.closePhoto) config.closePhoto(photoButton);
        img.onload=img.onerror=null;if(loader) loader.destroy();
        if(unmountDetails) {const unmount=unmountDetails;unmountDetails=null;unmount();}
        card.remove();
        if(popup===thisPopup) {popup=null;disposeCard=null;}
      };
      disposeCard=cleanup;thisPopup.on('remove',cleanup);thisPopup.openOn(map);
      if(config.mountDetails) unmountDetails=config.mountDetails(pin,info,audioSlot);
      thisPopup.update();
      if(audioOnly) {photoStatus.hidden=true;photoButton.hidden=true;}
      else if(!(pin.fileId||pin.imageUrl)) {photoStatus.textContent='写真なし';}
      else if(config.photoLoader) {
        loader=config.photoLoader(renderPhoto);
        retryPhoto.onclick=()=>{retryPhoto.hidden=true;loader.retry();};
        loader.open(pin.id,title);
      } else {
        retryPhoto.onclick=()=>{retryPhoto.hidden=true;showSource(pin.imageUrl);};showSource(pin.imageUrl);
      }
      infoButton.focus({preventScroll:true});
    };
    if(hasMarker) pins.zoomToShowLayer(marker,reveal);else reveal();
    return true;
  }
  function clearPins() {closePin();pins.clearLayers();}
  map.on('unload',()=>{generation++;cancelPending();if(disposeCard) disposeCard();resizeObserver.disconnect();classObserver.disconnect();if(panel)panel.removeEventListener('transitionend',positionControls);document.removeEventListener('pointerdown',outsideClick);toolbar.removeEventListener('keydown',escapeMenu);toggle.remove();menu.remove();});
  switchBackground('osm');
  return {pins,openPin,closePin,clearPins};
}
