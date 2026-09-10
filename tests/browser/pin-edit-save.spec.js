const { test, expect } = require('@playwright/test');

const pin = {id:'edit-pin-a',title:'before',description:'keep input',eventAt:'2025-07-20T09:16:45',lat:null,lng:null,color:'#e53935',icon:'default',status:'未対応',tags:[],links:[],fileId:'',hasAudio:false};
async function openEditor(page) {
  await page.goto('/production-edit');
  await page.evaluate(pin => {
    window.__productionEdit.state.pins = [pin,{...pin,id:'edit-pin-b',title:'other'}];
    window.__productionEdit.openPinEditor(pin);
    window.__gasMock.enqueue('updatePinDetails',{defer:'edit-save'});
  },pin);
  await page.locator('#edit-title').fill('after');
}

test('pending edit blocks double submit, Escape, switching pins and preview; success releases it', async ({page}) => {
  await openEditor(page);
  await page.locator('#edit-save').click();
  await expect(page.locator('#edit-save')).toBeDisabled();
  await expect(page.locator('#edit-save')).toHaveText('保存中…');
  await expect(page.locator('#edit-cancel')).toBeDisabled();
  await expect(page.locator('#edit-title')).toBeDisabled();
  await page.keyboard.press('Escape');
  await expect(page.locator('#edit-overlay')).toHaveClass(/open/);
  expect(await page.evaluate(() => {
    const api=window.__productionEdit;
    document.getElementById('edit-save').dispatchEvent(new MouseEvent('click'));
    document.getElementById('edit-toggle').dispatchEvent(new MouseEvent('click'));
    api.openPinEditor(api.state.pins[1]);
    return {id:api.state.editingPinId,preview:api.state.previewMode,pending:api.hasPendingMutationWork(),calls:window.__gasMock.calls.filter(c=>c.method==='updatePinDetails').length};
  })).toEqual({id:pin.id,preview:false,pending:true,calls:1});
  await page.evaluate(()=>window.__gasMock.resolve('edit-save',{ok:true,updatedAt:'saved'}));
  await expect(page.locator('#edit-overlay')).not.toHaveClass(/open/);
  expect(await page.evaluate(()=>window.__productionEdit.hasPendingMutationWork())).toBe(false);
  expect(await page.evaluate(()=>window.__productionEdit.state.pins[0].title)).toBe('after');
});

test('failed edit keeps draft and restores controls for a single retry', async ({page}) => {
  await openEditor(page);
  await page.locator('#edit-save').click();
  await page.evaluate(()=>window.__gasMock.reject('edit-save','temporary test failure'));
  await expect(page.locator('#edit-save')).toBeEnabled();
  await expect(page.locator('#edit-title')).toHaveValue('after');
  await expect(page.locator('#edit-event-at')).toHaveValue('2025-07-20T09:16:45');
  expect(await page.evaluate(()=>window.__productionEdit.hasPendingMutationWork())).toBe(false);
  await page.keyboard.press('Escape');
  await page.evaluate(()=>window.__gasMock.enqueue('updatePinDetails',{defer:'retry'}));
  await page.locator('#edit-save').click();
  await page.evaluate(()=>window.__gasMock.resolve('retry',{ok:true,updatedAt:'retried'}));
  await expect(page.locator('#edit-overlay')).not.toHaveClass(/open/);
  expect(await page.evaluate(()=>window.__gasMock.calls.filter(c=>c.method==='updatePinDetails').length)).toBe(2);
});

for (const samePin of [false,true]) test(`a stale response cannot close a newer session (same pin: ${samePin})`, async ({page}) => {
  await openEditor(page);
  await page.locator('#edit-save').click();
  await page.evaluate(samePin=> {
    // Simulate a replaced view/session while I/O is in flight; normal UI blocks it.
    window.__productionEdit.state.editingPinId=samePin ? 'edit-pin-a' : 'edit-pin-b';
    window.__productionEdit.state.pinEditSession += 1;
    window.__gasMock.resolve('edit-save',{ok:true,updatedAt:'old-response'});
  },samePin);
  await expect(page.locator('#edit-save')).toBeEnabled();
  expect(await page.evaluate(()=>window.__productionEdit.state.editingPinId)).toBe(samePin ? 'edit-pin-a' : 'edit-pin-b');
  await expect(page.locator('#edit-overlay')).toHaveClass(/open/);
});
