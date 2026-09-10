const {test, expect} = require('@playwright/test');
const png = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jD1sAAAAASUVORK5CYII=';
const photo = {ok:true, mimeType:'image/png', base64:png, byteLength:Buffer.from(png,'base64').length};
const pin = {id:'map-a',title:'谷の写真',description:'説明 <img src=x onerror=alert(1)>',eventAt:'2025-07-20T09:16:45',lat:35,lng:139,color:'#22643d',icon:'default',status:'未対応',tags:['植物'],links:[],fileId:'photo-a',imageUrl:'data:image/png;base64,'+png,hasAudio:false};
async function mapSelector(page) {
  const select=page.getByRole('combobox',{name:'背景地図'});
  if(!await select.isVisible()) await page.getByRole('button',{name:'地図を切り替え',exact:true}).click();
  return select;
}
async function setup(page,shared=false,pins=[pin]) {
  await page.route('https://**/*', route=>route.fulfill({contentType:'image/png',body:Buffer.from(png,'base64')}));
  await page.goto(shared?'/map-shared':'/map-edit');
  await page.evaluate(({pins,shared,photo})=>{
    const api=shared?window.__productionShared:window.__productionEdit;
    api.state.pins=pins;
    if(shared) api.renderSharedMap();
    else {
      window.__gasMock.enqueue('getPinPhotoData',{response:photo});
      api.map.setView([35,139],15);
      api.renderPins();
    }
  },{pins,shared,photo});
  // The production panels finish their 240 ms opening/resize transition.
  await page.waitForTimeout(300);
}
for(const shared of [false,true]) {
  test(`many pins are added in one cluster batch (${shared?'shared':'edit'})`,async({page})=>{
    await setup(page,shared);
    const result=await page.evaluate(shared=>{
      const proto=L.MarkerClusterGroup.prototype;
      const addOne=proto.addLayer,addMany=proto.addLayers;
      let single=0,batch=0,total=0;
      proto.addLayer=function(...args){single++;return addOne.apply(this,args);};
      proto.addLayers=function(items,...args){batch++;total+=items.length;return addMany.call(this,items,...args);};
      const api=shared?window.__productionShared:window.__productionEdit;
      const base=api.state.pins[0];
      api.state.pins=Array.from({length:300},(_,i)=>({...base,id:'bulk-'+i,lat:35+(i%30)*0.001,lng:139+Math.floor(i/30)*0.001}));
      const started=performance.now();
      try {if(shared)api.renderSharedMap();else api.renderPins();}
      finally {proto.addLayer=addOne;proto.addLayers=addMany;}
      return {single,batch,total,elapsedMs:performance.now()-started};
    },shared);
    console.log('cluster batch '+(shared?'shared':'edit')+': '+JSON.stringify(result));
    expect(result.single).toBe(0);
    expect(result.batch).toBe(1);
    expect(result.total).toBe(300);
  });
  test(`standard map choices are limited to the standard and combined relief maps (${shared?'shared':'edit'})`,async({page})=>{
    await setup(page,shared);
    const select=await mapSelector(page);
    await expect(select.locator('option')).toHaveText(['通常地図','陰影起伏図']);
    await expect(page.locator('.dtp-map-control')).not.toContainText('地図を追加したい方へ');
    await page.keyboard.press('Escape');
    if(shared) await page.getByRole('button',{name:'使い方',exact:true}).click();
    else {
      await page.getByRole('button',{name:'その他の操作',exact:true}).click();
      await page.getByRole('menuitem',{name:'使い方',exact:true}).click();
    }
    const help=page.getByRole('dialog',{name:'使い方',exact:true});
    await expect(help.getByRole('heading',{name:'地図を追加したい方へ',exact:true})).toBeVisible();
    await expect(help).toContainText('ご自身のアプリをカスタマイズ');
    await expect(help).toContainText('利用条件・料金');
    await help.getByRole('button',{name:'閉じる',exact:true}).click();
    await expect(help).not.toBeVisible();
  });
  test(`map pin opens photo popup, info is independent from enlargement (${shared?'shared':'edit'})`,async({page})=>{
    await setup(page,shared);
    await page.locator('.leaflet-marker-icon').first().click();
    const card=page.locator('.dtp-pin-card');
    await expect(card).toBeVisible();
    await expect(card.locator('.dtp-pin-title')).toHaveText(pin.title);
    await expect(card.locator('.dtp-pin-info')).toBeHidden();
    await card.getByRole('button',{name:'詳細情報',exact:true}).click();
    await expect(card.locator('.dtp-pin-info')).toBeVisible();
    await expect(card.locator('.dtp-pin-info')).toContainText(pin.description);
    await expect(card.locator('.dtp-pin-info img')).toHaveCount(0);
    await expect(page.locator(shared?'#shared-photo-viewer-overlay':'#photo-viewer-overlay')).not.toHaveClass(/open/);
    await card.getByRole('button',{name:'写真に戻る',exact:true}).click();
    await expect(card.getByRole('button',{name:'写真を拡大',exact:true})).toBeEnabled();
    await card.getByRole('button',{name:'写真を拡大',exact:true}).click();
    await expect(page.locator(shared?'#shared-photo-viewer-overlay':'#photo-viewer-overlay')).toHaveClass(/open/);
  });
}
test('clustering retains every pin and a list selection reveals a clustered pin',async({page})=>{
  await setup(page,false,[pin,{...pin,id:'map-b',title:'隣の写真',lng:139.00001}]);
  await expect(page.locator('.dtp-cluster')).toHaveCount(1);
  await expect(page.locator('.dtp-cluster')).toContainText('2');
  await page.evaluate(()=>window.__productionEdit.focusPin(window.__productionEdit.state.pins[1]));
  await expect(page.locator('.dtp-pin-title')).toHaveText('隣の写真');
  expect(await page.evaluate(()=>window.__productionEdit.state.pins.length)).toBe(2);
});
test('base maps preserve the center, zoom and pins in both directions',async({page})=>{
  await setup(page);
  const before=await page.evaluate(()=>{const m=window.__productionEdit.map; return [m.getCenter().lat,m.getCenter().lng,m.getZoom()];});
  await (await mapSelector(page)).selectOption('japan-terrain');
  await expect(page.locator('.dtp-basemap-status')).toContainText('範囲外は世界版');
  expect(await page.evaluate(()=>{const m=window.__productionEdit.map; return [m.getCenter().lat,m.getCenter().lng,m.getZoom()];})).toEqual(before);
  await (await mapSelector(page)).selectOption('osm');
  await expect(page.locator('.dtp-basemap-status')).toHaveText('通常地図');
  expect(await page.evaluate(()=>{const m=window.__productionEdit.map; return [m.getCenter().lat,m.getCenter().lng,m.getZoom()];})).toEqual(before);
  await expect(page.locator('.leaflet-marker-icon')).toHaveCount(1);
  await expect(page.locator('.leaflet-control-scale')).toBeVisible();
});

test('failed route load is visible and retries without discarding photos',async({page})=>{
  await setup(page);
  await page.evaluate(async pin=>{
    window.__gasMock.enqueue('getMapData',{response:[pin]});
    window.__gasMock.enqueue('getRouteGroups',{response:{ok:false,error:'network error'}});
    await window.__productionEdit.loadMapData();
    window.__gasMock.enqueue('getRouteGroups',{response:[]});
  },pin);
  await expect(page.locator('#route-load-error')).toBeVisible();
  await page.getByRole('button',{name:'ルートを再読み込み'}).click();
  await expect(page.locator('#route-load-error')).toBeHidden();
  expect(await page.evaluate(()=>window.__productionEdit.state.pins.length)).toBe(1);
});

test('failed basemap switch retains the previous map and permits retry',async({page})=>{
  await setup(page);
  await page.route('https://cyberjapandata.gsi.go.jp/**',route=>route.abort());
  await (await mapSelector(page)).selectOption('japan-terrain');
  await expect(page.locator('.dtp-basemap-status')).toContainText('元の地図');
  await expect(page.locator('.dtp-map-retry')).toBeVisible();
  await expect(page.locator('.leaflet-tile-pane img[src*="tile.openstreetmap.org"]').first()).toBeVisible();
  await page.unroute('https://cyberjapandata.gsi.go.jp/**');
  await page.locator('.dtp-map-retry').click();
  await expect(page.locator('.dtp-basemap-status')).toContainText('範囲外は世界版');
  await expect(page.locator('.leaflet-tile-pane img[src*="/hillshademap/"]').first()).toBeVisible();
  await expect(page.locator('.dtp-map-retry')).toBeHidden();
});

test('GPX failure retains existing tracks, partial failures warn, and retry clears the warning',async({page})=>{
  await setup(page);
  await page.evaluate(async()=>{
    const api=window.__productionEdit;
    api.state.tracks=[{trackId:'old-track',name:'既存のGPX'}];
    window.__gasMock.enqueue('getTracks',{response:{ok:false,error:'network error'}});
    await api.loadTracks().catch(()=>{});
  });
  await expect(page.locator('#track-load-error')).toBeVisible();
  expect(await page.evaluate(()=>window.__productionEdit.state.tracks[0].trackId)).toBe('old-track');
  await page.evaluate(()=>window.__gasMock.enqueue('getTracks',{response:{ok:true,tracks:[],warnings:[{code:'TRACK_SEGMENTS_CORRUPTED',trackId:'broken'}]}}));
  await page.getByRole('button',{name:'GPXを再読み込み'}).click();
  await expect(page.locator('#track-load-message')).toContainText('1件');
  await page.evaluate(()=>window.__gasMock.enqueue('getTracks',{response:{ok:true,tracks:[],warnings:[]}}));
  await page.getByRole('button',{name:'GPXを再読み込み'}).click();
  await expect(page.locator('#track-load-error')).toBeHidden();
});

test('closing a loading photo prevents late responses reopening or replacing it',async({page})=>{
  await setup(page);
  await page.evaluate(()=>{window.__gasMock.clear();window.__gasMock.enqueue('getPinPhotoData',{defer:'old-photo'});});
  await page.locator('.leaflet-marker-icon').first().click();
  await expect(page.locator('.dtp-photo-status')).toHaveText('写真を読み込み中…');
  await page.locator('.dtp-photo-popup .leaflet-popup-close-button').click();
  await page.evaluate(photo=>window.__gasMock.resolve('old-photo',photo),photo);
  await expect(page.locator('.dtp-pin-card')).toHaveCount(0);
});

test('mobile photo card and info controls stay inside the visible map',async({page})=>{
  await page.setViewportSize({width:375,height:812});
  await setup(page);
  await (await mapSelector(page)).click();
  await page.keyboard.press('Escape');
  await page.locator('.leaflet-marker-icon').first().click();
  const card=page.locator('.dtp-pin-card');
  await expect(card).toBeVisible();
  await card.getByRole('button',{name:'詳細情報',exact:true}).click();
  const box=await card.boundingBox();
  expect(box.x).toBeGreaterThanOrEqual(0);
  expect(box.x+box.width).toBeLessThanOrEqual(375);
  expect(box.y).toBeGreaterThanOrEqual(0);
  expect(box.y+box.height).toBeLessThanOrEqual(812);
  await page.screenshot({path:'test-results/map-mobile.png'});
});

test('no-photo pin remains readable and exposes the existing operations',async({page})=>{
  await setup(page,false,[{...pin,fileId:'',imageUrl:''}]);
  await page.locator('.leaflet-marker-icon').first().click();
  await expect(page.locator('.dtp-photo-status')).toHaveText('写真なし');
  await page.getByRole('button',{name:'詳細情報',exact:true}).click();
  await page.locator('.dtp-pin-card').getByRole('button',{name:'その他の操作',exact:true}).click();
  await expect(page.locator('#pin-detail-overlay')).toHaveClass(/open/);
  await expect(page.locator('.dtp-pin-card')).toHaveCount(0);
});

for(const shared of [false,true]) for(const width of [375,1280]) {
  test(`Japan maps use topbar controls and bottom-right scale (${shared?'shared':'edit'}, ${width})`,async({page})=>{
    await page.setViewportSize({width,height:812});
    await setup(page,shared);
    if(!shared) await page.evaluate(()=>window.__productionEdit.renderAccessMode());
    const bar=page.locator(shared?'#shared-topbar':'#topbar');
    await expect(bar.getByRole('button',{name:'地図を切り替え',exact:true})).toBeVisible();
    const before=await page.evaluate(shared=>{
      const m=shared?window.__productionShared.ensureSharedMap():window.__productionEdit.map;
      return [m.getCenter().lat,m.getCenter().lng,m.getZoom()];
    },shared);
    const select=await mapSelector(page);
    await select.selectOption('japan-terrain');
    await expect(page.locator('.leaflet-tile-pane img[src*="/xyz/hillshademap/"]').first()).toBeVisible();
    expect(await page.evaluate(shared=>{
      const m=shared?window.__productionShared.ensureSharedMap():window.__productionEdit.map;
      return [m.getCenter().lat,m.getCenter().lng,m.getZoom()];
    },shared)).toEqual(before);
    const scale=await page.locator('.leaflet-control-scale').boundingBox();
    const attribution=await page.locator('.leaflet-control-attribution').boundingBox();
    const bounds=await page.locator(shared?'#shared-map':'#map').boundingBox();
    expect(scale.x+scale.width).toBeGreaterThan(bounds.x+bounds.width-30);
    expect(scale.y+scale.height).toBeLessThanOrEqual(attribution.y+1);
    const picker=await page.locator('.dtp-map-control').boundingBox();
    expect(picker.x).toBeGreaterThanOrEqual(0);
    expect(picker.x+picker.width).toBeLessThanOrEqual(width);
    for(const button of await bar.getByRole('button').all()) {
      if(!await button.isVisible())continue;
      const box=await button.boundingBox();
      expect(box.x).toBeGreaterThanOrEqual(0);
      expect(box.x+box.width).toBeLessThanOrEqual(width);
    }
    await page.keyboard.press('Escape');
    await expect(select).toBeHidden();
    await expect(bar.getByRole('button',{name:'地図を切り替え',exact:true})).toBeFocused();
  });
}

for(const zoom of [1,12]) test(`Relief uses the world underlay overseas without domestic requests (zoom ${zoom})`,async({page})=>{
  await setup(page);
  const domesticRequests=[];
  page.on('request',request=>{if(request.url().includes('/xyz/hillshademap/'))domesticRequests.push(request.url());});
  await page.evaluate(zoom=>window.__productionEdit.map.setView([50,-100],zoom,{animate:false}),zoom);
  const select=await mapSelector(page);
  await select.selectOption('japan-terrain');
  await expect(page.locator('.dtp-basemap-status')).toContainText('範囲外は世界版');
  await expect(page.locator('.leaflet-tile-pane img[src*="/earthhillshade/"]').first()).toBeVisible();
  expect(await page.locator('.leaflet-tile-pane img[src*="/earthhillshade/"]').first().getAttribute('src')).toContain('/earthhillshade/'+Math.min(zoom,8)+'/');
  expect(domesticRequests).toHaveLength(0);
  expect(await page.evaluate(()=>window.__productionEdit.map.getZoom())).toBe(zoom);
});

test('Relief overlays detailed Japan tiles on world relief at maximum zoom',async({page})=>{
  await setup(page);
  await page.evaluate(()=>window.__productionEdit.map.setView([35,139],19,{animate:false}));
  await (await mapSelector(page)).selectOption('japan-terrain');
  await expect(page.locator('.dtp-basemap-status')).toContainText('範囲外は世界版');
  for(const [layer,zoom] of [['hillshademap',16],['earthhillshade',8]]) {
    const tile=page.locator('.leaflet-tile-pane img[src*="/'+layer+'/"]').first();
    await expect(tile).toBeVisible();
    expect(await tile.getAttribute('src')).toContain('/'+layer+'/'+zoom+'/');
  }
  expect(await page.evaluate(()=>window.__productionEdit.map.getZoom())).toBe(19);
});

test('failed Japan tiles preserve the preceding map and allow another selection',async({page})=>{
  await setup(page);
  const select=await mapSelector(page);
  await page.route('https://cyberjapandata.gsi.go.jp/xyz/hillshademap/**',route=>route.abort());
  await select.selectOption('japan-terrain');
  await expect(page.locator('.dtp-basemap-status')).toContainText('元の地図');
  await expect(page.locator('.dtp-map-retry')).toBeVisible();
  await expect(page.locator('.leaflet-tile-pane img[src*="tile.openstreetmap.org"]').first()).toBeVisible();
  await select.selectOption('osm');
  await expect(page.locator('.dtp-basemap-status')).toHaveText('通常地図');
  await expect(page.locator('.dtp-map-retry')).toBeHidden();
});
