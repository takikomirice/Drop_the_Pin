const {test,expect}=require('@playwright/test');
const pin={id:'startup-pin',title:'読み込んだピン',lat:35,lng:139,tags:[],links:[],fileId:'',imageUrl:'',hasAudio:false};
const methods=['getAppSettings','getMapData','getRouteGroups','getTracks'];
async function begin(page) {
  await page.goto('/production-edit');
  await page.evaluate(methods=>{
    window.__gasMock.clear();
    for(const method of methods) window.__gasMock.enqueue(method,{defer:method});
    window.__gasMock.enqueue('ensureMediaDriveStructure',{response:{ok:true}});
    window.__productionEdit.state.initializing=true;
    window.__startupPromise=window.__productionEdit.initializeApp();
  },methods);
}
test('startup requests independent data together and gates editing until all settle',async({page})=>{
  await begin(page);
  expect(await page.evaluate(()=>window.__gasMock.calls.map(c=>c.method).sort())).toEqual([...methods].sort());
  await expect(page.locator('#startup-load-status')).toBeVisible();
  await expect(page.locator('#pin-add-btn')).toBeDisabled();
  await page.evaluate(pin=>{
    window.__gasMock.resolve('getTracks',{ok:true,tracks:[]});
    window.__gasMock.resolve('getMapData',[pin]);
    window.__gasMock.resolve('getRouteGroups',[]);
  },pin);
  await expect(page.getByRole('button',{name:'読み込んだピンの詳細を表示',exact:true})).toBeVisible();
  expect(await page.evaluate(()=>window.__productionEdit.state.initializing)).toBe(true);
  await expect(page.locator('#pin-add-btn')).toBeDisabled();
  await page.evaluate(async()=>{window.__gasMock.resolve('getAppSettings',{ok:true});await window.__startupPromise;});
  await expect(page.locator('#startup-load-status')).toBeHidden();
  await expect(page.locator('#pin-add-btn')).toBeEnabled();
  expect(await page.evaluate(()=>window.__runtime.unhandledRejections)).toEqual([]);
});
test('parallel startup handles independent failures and exposes a persistent pin retry',async({page})=>{
  await begin(page);
  expect(await page.evaluate(()=>window.__gasMock.calls.map(c=>c.method).sort())).toEqual([...methods].sort());
  await page.evaluate(async()=>{
    for(const method of ['getAppSettings','getMapData','getRouteGroups','getTracks']) window.__gasMock.reject(method,'permission denied');
    await window.__startupPromise;
  });
  await expect(page.locator('#startup-load-status')).toContainText('読み込めませんでした');
  await expect(page.locator('#startup-load-retry')).toBeVisible();
  await expect(page.locator('#route-load-error')).toBeVisible();
  await expect(page.locator('#track-load-error')).toBeVisible();
  expect(await page.evaluate(()=>window.__runtime.unhandledRejections)).toEqual([]);
});
